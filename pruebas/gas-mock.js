// Simulación mínima de los servicios de Google Apps Script para probar Code.gs en Node.
const vm = require('vm');
const fs = require('fs');

function crearEntorno(opciones = {}) {
  let ahora = opciones.ahora ? new Date(opciones.ahora).getTime() : Date.now();
  const RealDate = Date;
  class FakeDate extends RealDate {
    constructor(...a) { if (a.length === 0) super(ahora); else super(...a); }
    static now() { return ahora; }
  }
  const registro = { correos: [], fetch: [], logs: [], triggers: [] };

  const encadenable = () => new Proxy({}, { get: (t, k) => k === 'build' ? () => ({}) : () => encadenable() });

  class Rango {
    constructor(hoja, r, c, nr, nc) { Object.assign(this, { hoja, r, c, nr, nc }); }
    getValues() {
      const out = [];
      for (let i = 0; i < this.nr; i++) {
        const fila = [];
        for (let j = 0; j < this.nc; j++) {
          const v = (this.hoja.datos[this.r - 1 + i] || [])[this.c - 1 + j];
          fila.push(v === undefined ? '' : v);
        }
        out.push(fila);
      }
      return out;
    }
    getValue() { return this.getValues()[0][0]; }
    setValues(vals) {
      if (vals.length !== this.nr || vals[0].length !== this.nc) throw new Error(`Dimensiones no coinciden: datos ${vals.length}x${vals[0].length}, rango ${this.nr}x${this.nc}`);
      vals.forEach((f, i) => f.forEach((v, j) => this.hoja.escribir(this.r + i, this.c + j, v)));
      return this;
    }
    setValue(v) { this.hoja.escribir(this.r, this.c, v); return this; }
    clearContent() { for (let i = 0; i < this.nr; i++) for (let j = 0; j < this.nc; j++) this.hoja.escribir(this.r + i, this.c + j, ''); return this; }
  }
  // Métodos de formato: no hacen nada pero devuelven el rango (encadenables)
  ['setNumberFormat','setFontWeight','setBackground','setFontColor','setWrap','setHorizontalAlignment','insertCheckboxes','clearDataValidations','setDataValidation']
    .forEach(m => { Rango.prototype[m] = function () { return this; }; });

  class Hoja {
    constructor(nombre) { this.nombre = nombre; this.datos = []; this.maxRows = 1000; this.maxCols = 26; }
    escribir(r, c, v) {
      while (this.datos.length < r) this.datos.push([]);
      if (typeof v === 'string' && v.startsWith("'")) v = v.slice(1);       // como Sheets: ' fuerza texto y no se guarda
      this.datos[r - 1][c - 1] = v;
    }
    getName() { return this.nombre; }
    getLastRow() {
      for (let i = this.datos.length - 1; i >= 0; i--) if ((this.datos[i] || []).some(v => v !== '' && v !== undefined)) return i + 1;
      return 0;
    }
    getMaxRows() { return this.maxRows; }
    getMaxColumns() { return this.maxCols; }
    getRange(a, b, c, d) {
      if (typeof a === 'string') return new Rango(this, 1, 1, 1, 1);   // rangos A1 solo se usan para formato
      return new Rango(this, a, b, c || 1, d || 1);
    }
    getDataRange() { const n = this.getLastRow(); const w = Math.max(0, ...this.datos.map(f => f.length)); return new Rango(this, 1, 1, n, w); }
    appendRow(fila) { const r = this.getLastRow() + 1; fila.forEach((v, j) => this.escribir(r, j + 1, v)); }
    insertColumnAfter(col) { this.datos.forEach(f => f.splice(col, 0, '')); }
    setFrozenRows() {} setConditionalFormatRules() {} setColumnWidth() {}
  }

  class Libro {
    constructor() { this.hojas = [new Hoja('Hoja 1')]; }
    getSheetByName(n) { return this.hojas.find(h => h.nombre === n) || null; }
    insertSheet(n) { const h = new Hoja(n); this.hojas.push(h); return h; }
    getSheets() { return this.hojas.slice(); }
    deleteSheet(h) { this.hojas = this.hojas.filter(x => x !== h); }
    setSpreadsheetTimeZone() {} setActiveSheet() {}
  }
  const libro = new Libro();
  const props = {};

  const pad = n => String(n).padStart(2, '0');
  const ctx = {
    console: { log: (...a) => registro.logs.push(a.join(' ')), error: (...a) => registro.logs.push('ERROR ' + a.join(' ')) },
    Date: FakeDate, Math, JSON, String, Number, Object, Array, Error, RegExp, Intl, parseInt, parseFloat, isNaN,
    SpreadsheetApp: {
      getActiveSpreadsheet: () => libro,
      flush: () => {},
      newDataValidation: encadenable,
      newConditionalFormatRule: encadenable,
    },
    PropertiesService: { getScriptProperties: () => ({
      getProperty: k => (k in props ? props[k] : null),
      setProperty: (k, v) => { props[k] = String(v); },
      setProperties: o => Object.entries(o).forEach(([k, v]) => { props[k] = String(v); }),
    }) },
    LockService: { getScriptLock: () => ({ tryLock: () => !opciones.lockFalla, waitLock: () => { if (opciones.lockFalla) throw new Error('Lock timeout: another process was holding the lock for too long.'); }, releaseLock: () => {} }) },
    Utilities: { formatDate: (d, tz, fmt) => {
      const p = Object.fromEntries(new Intl.DateTimeFormat('en-CA', { timeZone: tz, year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', hourCycle: 'h23' })
        .formatToParts(d).map(x => [x.type, x.value]));
      return fmt.replace('yyyy', p.year).replace('MM', p.month).replace('dd', p.day).replace('HH', p.hour).replace('mm', p.minute);
    } },
    ContentService: { MimeType: { JSON: 'json' }, createTextOutput: s => ({ texto: s, setMimeType() { return this; } }) },
    CacheService: { getScriptCache: () => ({ get: k => (registro.cache || {})[k] || null, put: (k, v) => { (registro.cache = registro.cache || {})[k] = v; } }) },
    MailApp: { sendEmail: (...a) => registro.correos.push(a) },
    UrlFetchApp: { fetch: (...a) => { registro.fetch.push(a); return { getContentText: () => '{"ok":true,"result":[]}' }; } },
    ScriptApp: {
      getProjectTriggers: () => registro.triggers,
      deleteTrigger: t => { registro.triggers = registro.triggers.filter(x => x !== t); },
      newTrigger: fn => { const t = { fn, getHandlerFunction: () => fn }; const c = new Proxy({}, { get: (_, k) => k === 'create' ? () => { registro.triggers.push(t); return t; } : () => c }); return { timeBased: () => c }; },
    },
  };
  vm.createContext(ctx);
  const codigo = fs.readFileSync(opciones.archivo || __dirname + '/../Code.gs', 'utf8');
  vm.runInContext(codigo + '\n;this.__exports = { CONFIG, COL, configurarHojas, doGet, doPost, leerMenu, archivarPedidos, instalarArchivoDiario, enlaceWhatsApp, limpiar, repararEnlacesWhatsApp, leerContadores_, marcarPruebasComoRechazadas };', ctx);

  return {
    ...ctx.__exports, libro, props, registro,
    fecha: iso => new FakeDate(iso),   // fechas creadas "dentro" de Apps Script, como las que devuelve Sheets
    setAhora: iso => { ahora = new RealDate(iso).getTime(); },
    avanzar: ms => { ahora += ms; },
    post: body => JSON.parse(ctx.__exports.doPost({ postData: { contents: typeof body === 'string' ? body : JSON.stringify(body) } }).texto),
    get: q => JSON.parse(ctx.__exports.doGet({ parameter: q || {} }).texto),
    hoja: n => libro.getSheetByName(n),
  };
}
module.exports = { crearEntorno };
