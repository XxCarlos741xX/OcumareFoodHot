/**
 * =====================================================================
 *  BACKEND DE PEDIDOS — Ocumare Food Hot (Google Apps Script, gratis)
 * =====================================================================
 *  Qué hace:
 *   • GET  ?accion=menu  → entrega a la web el menú de la hoja "Menu"
 *   • POST (pedido)      → valida, recalcula precios, guarda en "Pedidos"
 *                          y avisa al equipo por Telegram/correo
 *   • Cada madrugada archiva los pedidos cerrados en "Historial AAAA"
 *
 *  Numeración de pedidos:
 *   • "N° del día": 1, 2, 3… se reinicia cada día. Es el que ve el cliente
 *     y el que se canta en la cocina.
 *   • "ID": número único que nunca se repite (aunque se borren filas).
 *     Sirve para buscar un pedido o cuadrar cuentas.
 *
 *  Instalación (una sola vez):
 *   1. Extensiones → Apps Script en la hoja. Pega este archivo y guarda.
 *   2. Ejecuta "configurarHojas" y acepta los permisos.
 *   3. Ejecuta "instalarArchivoDiario" (programa el archivo automático).
 *   4. Implementar → Nueva implementación → "Aplicación web"
 *        Ejecutar como: Yo  ·  Quién tiene acceso: Cualquier usuario
 *      Copia la URL /exec y pégala en API_URL de index.html.
 *
 *  ⚠️ Si cambias ESTE código después: Implementar → Gestionar
 *     implementaciones → ✏️ → Versión: "Nueva versión". La URL no cambia.
 *     (Cambiar el menú o los estados en la hoja NO requiere esto.)
 * =====================================================================
 */

// ======================= ✏️ CONFIGURACIÓN =======================
const CONFIG = {
  TELEGRAM_TOKEN: '',        // Token que te da @BotFather. Ej: '7123456789:AAH...'
  TELEGRAM_CHAT_ID: '',      // ID del grupo del equipo. Usa la función verChatId() para obtenerlo
  EMAIL_AVISO: '',           // Opcional: un correo que también reciba cada pedido (máx. 100/día)
  MONEDA: '$',
  CODIGO_PAIS: '58',         // Código por defecto para el enlace de WhatsApp (58 = Venezuela)
  ZONA_HORARIA: 'America/Caracas',  // Define cuándo empieza "el día" para el N° del día
  HORA_ARCHIVO: 4,           // Hora (0-23) en que se archivan los pedidos cerrados
  MAX_UNIDADES: 30,          // Máximo por producto en un pedido (anti-bromas)
};
// ================================================================

const HOJA_PEDIDOS = 'Pedidos';
const HOJA_MENU = 'Menu';
const PREFIJO_HISTORIAL = 'Historial ';
const ESTADOS = ['Nuevo', 'Preparando', 'En camino', 'Entregado', 'Rechazado'];
const ESTADOS_CERRADOS = ['Entregado', 'Rechazado'];
const COLORES_ESTADO = ['#FFF3B0', '#FFD8A8', '#C5E1FF', '#C8F2C8', '#F5C2C0'];
const COLUMNAS_PEDIDOS = ['N° del día', 'ID', 'Fecha', 'Estado', 'Cliente', 'Teléfono', 'Entrega',
                          'Dirección', 'Productos', 'Total', 'Pago', 'Notas', 'WhatsApp cliente'];
// Posición (1 = columna A) de cada dato en "Pedidos" e "Historial"
const COL = { DIA: 1, ID: 2, FECHA: 3, ESTADO: 4, CLIENTE: 5, TELEFONO: 6, ENTREGA: 7, DIRECCION: 8,
              PRODUCTOS: 9, TOTAL: 10, PAGO: 11, NOTAS: 12, WHATSAPP: 13 };
const COLUMNAS_MENU = ['id', 'categoria', 'nombre', 'descripcion', 'precio', 'disponible', 'imagen', 'destacado'];


/* ---------------------------------------------------------------
   Crea o actualiza las hojas. Se puede ejecutar las veces que sea:
   no borra pedidos ni menú, y migra la hoja de la versión anterior.
   --------------------------------------------------------------- */
function configurarHojas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone(CONFIG.ZONA_HORARIA);

  // --- Hoja Pedidos ---
  const p = ss.getSheetByName(HOJA_PEDIDOS) || ss.insertSheet(HOJA_PEDIDOS);
  if (p.getLastRow() === 0) p.appendRow(COLUMNAS_PEDIDOS);
  migrarNumeracion_(p);
  darFormatoPedidos_(p);

  // --- Hoja Menu ---
  let m = ss.getSheetByName(HOJA_MENU);
  if (!m) {
    m = ss.insertSheet(HOJA_MENU);
    m.appendRow(COLUMNAS_MENU);
    // Menú real tomado de las historias destacadas de Instagram (@ocumarefoodhot)
    const ejemplo = [
      ["HB1", "Hamburguesas", "Combo dúo mini", "2 hamburguesas mini de carne o pollo, vegetales y 3 salsas básicas, con papas a la francesa", 4.99, true, "img/duo-mini.jpg", false],
      ["HB2", "Hamburguesas", "Tradicional", "Pan de orégano, 150 g de pollo o carne, queso Kraft, vegetales y 3 salsas básicas, con papas a la francesa o rústicas", 4.99, true, "img/tradicional.jpg", false],
      ["HB3", "Hamburguesas", "Pollo crispy", "120 g de milanesa rebozada en panko, lechuga, tomate, cebollas caramelizadas, salsa tártara y queso Kraft, con papas a la francesa", 6.99, true, "img/pollo-crispy.jpg", false],
      ["HB4", "Hamburguesas", "Ocumare Gourmet", "250 g de carne rellena de mozzarella envuelta en tocineta, vegetales, salsa big mac, pepinillos y queso Kraft, con papas a la francesa", 8.99, true, "img/ocumare-gourmet.jpg", true],
      ["HB5", "Hamburguesas", "Poseidón", "100 g de camarón al ajillo y 100 g de pulpo laqueado, lechuga, salsa spicy y queso crema, con papas a la francesa o rústicas", 8.99, true, "img/poseidon.jpg", false],
      ["HB6", "Hamburguesas", "Especial", "150 g de pollo o carne, queso Kraft, vegetales, 3 salsas, huevo, jamón y tocineta, con papas a la francesa o rústicas", 8.99, true, "img/especial.jpg", false],
      ["HB7", "Hamburguesas", "Mixta", "120 g de pollo y 150 g de carne, queso Kraft, vegetales, 3 salsas, huevo, jamón y tocineta, con papas a la francesa o rústicas", 9.99, true, "img/mixta.jpg", false],
      ["HB8", "Hamburguesas", "Cheese burger", "Doble carne de 150 g, triple queso Kraft, pepinillos, sweet relish y salsa big mac, con papas a la francesa o rústicas", 9.99, true, "img/cheese-burger.jpg", false],
      ["PP1", "Perros y pepitos", "Perro normal", "Pan de orégano, 1 salchicha, ensalada, papitas y 3 salsas básicas", 2.5, true, "img/perro-normal.jpg", false],
      ["PP2", "Perros y pepitos", "Perro especial", "Salchicha, ensalada, papitas, 3 salsas básicas y queso amarillo", 2.99, true, "img/perro-especial.jpg", false],
      ["PP3", "Perros y pepitos", "Perro gourmet", "Salchicha, ensalada, papitas, salsa tártara, queso amarillo, pepinillo, sweet relish y tocineta", 4.99, true, "img/perro-gourmet.jpg", false],
      ["PP4", "Perros y pepitos", "Perro polaco", "Salchicha polaca, ensalada, papitas, 3 salsas básicas y queso amarillo", 4.99, true, "img/perro-polaco.jpg", false],
      ["PP5", "Perros y pepitos", "Ternerito", "Pan de orégano con 150 g de pollo o carne, ensalada, salsa tártara, queso amarillo y tocineta", 7, true, "img/ternerito.jpg", false],
      ["PP6", "Perros y pepitos", "Pepito normal", "Pan de orégano de 40 cm, 300 g de carne o pollo, lechuga, tomate, cebolla, 3 salsas, papitas, queso amarillo y jamón, con papas fritas", 16, true, "img/pepito-normal.jpg", false],
      ["PP7", "Perros y pepitos", "Pepito especial", "Pan de 40 cm, 300 g de carne o pollo, vegetales, 3 salsas, papitas, queso amarillo, jamón, tocineta, huevo y 2 chorizos, con papas fritas", 21, true, "", false],
      ["PP8", "Perros y pepitos", "Pepito playero", "Pan de orégano, 250 g de pollo y 250 g de camarones rebozados, lechuga y salsa tártara, con 200 g de papas rústicas", 25, true, "img/pepito-playero.jpg", true],
      ["PO1", "Pollo", "Pollo frito 1 pieza", "1 pieza rebozada con 100 g de papas fritas, 2 arepitas y ensalada", 5, true, "img/pollo-1.jpg", false],
      ["PO2", "Pollo", "Pollo frito 2 piezas", "2 piezas rebozadas con 150 g de papas fritas, 5 arepitas y ensalada", 9, true, "img/pollo-2.jpg", false],
      ["PO3", "Pollo", "Pollo frito 3 piezas", "3 piezas rebozadas con 250 g de papas fritas, 8 arepitas y ensalada", 13, true, "img/pollo-3.jpg", false],
      ["PO4", "Pollo", "Tender de pollo", "120 g de milanesa en tiras rebozadas con 150 g de papas fritas", 5, true, "img/tender.jpg", false],
      ["PO5", "Pollo", "Alitas 16 piezas", "Alitas rebozadas en salsa BBQ y mostaza miel, 200 g de papas fritas, 6 arepitas y ensalada", 14, true, "img/alitas.jpg", true],
      ["PO6", "Pollo", "Club house", "250 g de pollo, 2 huevos, jamón, queso amarillo, lechuga, tomate, cebolla y 3 salsas en pan blanco, con 300 g de papas fritas", 17, true, "img/club-house.jpg", false],
      ["PO7", "Pollo", "Patacón", "2 ruedas de patacón rellenas con 150 g de proteína, vegetales, 3 salsas, jamón, queso amarillo y 1 huevo", 9, true, "", false],
      ["PA1", "Parrillas", "Parrilla 1 persona", "120 g de pollo, 120 g de bistec, 1 chorizo y 200 g de papas fritas, con ensalada", 15, true, "img/parrilla-1.jpg", false],
      ["PA2", "Parrillas", "Parrilla para 2", "250 g de pollo, 250 g de bistec, 2 chorizos, 300 g de papas fritas y 6 arepitas, con ensalada", 25, true, "img/parrilla-2.jpg", false],
      ["PA3", "Parrillas", "Parrilla mar y tierra", "250 g de pollo, 250 g de bistec, 250 g de mariscos y 300 g de papas fritas, con ensalada", 28, true, "img/parrilla-mar-tierra.jpg", true],
      ["PE1", "Pescados y mariscos", "Pargo rosado", "Pargo frito crujiente con arroz, tostones y ensalada rallada", 7, true, "img/pargo-rosado.jpg", true],
      ["PE2", "Pescados y mariscos", "Pargo blanco", "Pargo frito con 3 tostones, arroz y ensalada", 7, true, "img/pargo-blanco.jpg", false],
      ["PE3", "Pescados y mariscos", "Jurel", "Rueda de jurel a la plancha bañada en ajillo, con arroz, 3 tostones y ensalada", 8, true, "img/jurel.jpg", false],
      ["PE4", "Pescados y mariscos", "Merluza", "Filet de merluza rebozado o al ajillo, con 3 tostones, arroz y ensalada", 8, true, "img/merluza.jpg", false],
      ["PE5", "Pescados y mariscos", "Camarones al ajillo", "300 g de camarón salteado en ajo, mantequilla y vino blanco, con 200 g de papas fritas", 14, true, "img/camarones-ajillo.jpg", false],
      ["PS1", "Pastas y risottos", "Pasta boloñesa", "Vermicelli con 200 g de salsa boloñesa, queso pecorino y 2 tajadas", 5.99, true, "img/pasta-bolonesa.jpg", false],
      ["PS2", "Pastas y risottos", "Pasta cremosa de pollo", "Linguine con tocineta crocante, pollo a las finas hierbas y crema de leche, con pan al ajillo y pecorino", 7.99, true, "img/pasta-cremosa-pollo.jpg", false],
      ["PS3", "Pastas y risottos", "Pasta marinera", "Linguine con 350 g de mariscos salteados en vino blanco y bisque, con parmesano y pan de ajo", 14, true, "img/pasta-marinera.jpg", false],
      ["PS4", "Pastas y risottos", "Pasta frutos del mar", "Linguine con 350 g de mariscos en vino blanco, finas especias y crema, con parmesano y pan de ajo", 14, true, "img/pasta-frutos-mar.jpg", false],
      ["PS5", "Pastas y risottos", "Risotto de pollo", "Risotto cremoso con pollo, finas hierbas y tocineta en crema de leche, con pecorino y pan de ajo", 7.99, true, "", false],
      ["PS6", "Pastas y risottos", "Risotto de mariscos", "Risotto cremoso con mariscos flameados en vino blanco y crema, con pecorino y pan de ajo", 14, true, "img/risotto-mariscos.jpg", false],
      ["EN1", "Ensaladas", "César con pollo", "120 g de pollo a la plancha, lechuga, crotones, tocineta, pecorino y aderezo César", 8, true, "img/cesar-pollo.jpg", false],
      ["EN2", "Ensaladas", "César con camarón", "120 g de camarón al ajillo, lechuga, crotones, tocineta, pecorino y aderezo César", 10, true, "img/cesar-camaron.jpg", false],
      ["SO1", "Sopas", "Sopa de res", "Sopa grande de costilla con 2 arepitas", 6, true, "img/sopa-res.jpg", false],
      ["SO2", "Sopas", "Sopa de pescado", "Sopa grande de pescado con 2 arepitas", 6, true, "img/sopa-pescado.jpg", false],
      ["SO3", "Sopas", "Fosforera", "Sopa grande de mariscos con 2 arepitas", 10, true, "img/fosforera.jpg", false],
      ["AL1", "Almuerzos", "Ragú de carne", "200 g de ragú de carne con papas y zanahoria, con 3 contornos", 4.99, true, "img/ragu.jpg", false],
      ["AL2", "Almuerzos", "Bistec encebollado", "120 g de bistec con especias y cebolla salteada, con 3 contornos", 4.99, true, "img/bistec-encebollado.jpg", false],
      ["AL3", "Almuerzos", "Chuleta de cerdo", "Chuleta fresca con ajo, especias y vino blanco, con 3 contornos", 4.99, true, "img/chuleta-cerdo.jpg", false],
      ["AL4", "Almuerzos", "Chuleta ahumada", "Chuleta ahumada al ajillo con 3 contornos", 4.99, true, "img/chuleta-ahumada.jpg", false],
      ["AL5", "Almuerzos", "Pollo guisado", "1 pieza de pollo guisado con papas y zanahorias, con 3 contornos", 4.99, true, "img/pollo-guisado.jpg", false],
      ["AL6", "Almuerzos", "Milanesa de pollo", "120 g de milanesa con ajo, orégano, pimienta, mantequilla y vino blanco, con 3 contornos", 4.99, true, "img/milanesa-pollo.jpg", false],
      ["AL7", "Almuerzos", "Pabellón criollo", "Arroz blanco, caraotas, carne mechada, 2 tajadas y queso llanero", 5.99, true, "img/pabellon.jpg", false],
      ["AL8", "Almuerzos", "Milanesa parmesana", "Milanesa de pollo con salsa napolitana, mozzarella y pecorino gratinados, con 3 contornos", 5.99, true, "img/milanesa-parmesana.jpg", false]
    ];
    m.getRange(2, 1, ejemplo.length, ejemplo[0].length).setValues(ejemplo);
  }
  m.setFrozenRows(1);
  m.getRange(1, 1, 1, COLUMNAS_MENU.length).setFontWeight('bold').setBackground('#E9B949');
  m.getRange('F2:F200').insertCheckboxes();   // ✅ disponible / ⬜ agotado
  m.getRange('H2:H200').insertCheckboxes();   // ✅ sale en "Lo más pedido 🔥"
  m.getRange('E2:E200').setNumberFormat('#,##0.00');

  // Borra la hoja vacía que Google crea por defecto
  ['Hoja 1', 'Sheet1', 'Hoja1'].forEach(n => {
    const h = ss.getSheetByName(n);
    if (h && h.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(h);
  });
  ss.setActiveSheet(p);
  console.log('Listo ✅ Hojas configuradas. Último ID: ' + leerContadores_().ultimoId);
}

// Formato común de "Pedidos" y de cada "Historial AAAA"
function darFormatoPedidos_(h) {
  h.getRange(1, 1, 1, COLUMNAS_PEDIDOS.length).setValues([COLUMNAS_PEDIDOS])
    .setFontWeight('bold').setBackground('#E8811E').setFontColor('#FFFFFF');
  h.setFrozenRows(1);
  const filas = h.getMaxRows() - 1;
  h.getRange(2, 1, filas, h.getMaxColumns()).clearDataValidations();
  h.getRange(2, COL.ESTADO, filas, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(ESTADOS, true).setAllowInvalid(false).build());
  const rango = h.getRange(2, 1, filas, COLUMNAS_PEDIDOS.length);
  const letraEstado = String.fromCharCode(64 + COL.ESTADO);
  h.setConditionalFormatRules(ESTADOS.map((e, i) =>
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$' + letraEstado + '2="' + e + '"')
      .setBackground(COLORES_ESTADO[i]).setRanges([rango]).build()));
  h.getRange(2, COL.DIA, filas, 1).setHorizontalAlignment('center').setFontWeight('bold');
  h.getRange(2, COL.ID, filas, 1).setHorizontalAlignment('center').setFontColor('#777777');
  h.getRange(2, COL.FECHA, filas, 1).setNumberFormat('dd/mm/yyyy hh:mm');
  h.getRange(2, COL.TELEFONO, filas, 1).setNumberFormat('@');   // texto: no perder ceros ni el +
  h.getRange(2, COL.TOTAL, filas, 1).setNumberFormat('#,##0.00');
  h.getRange(1, COL.PRODUCTOS, h.getMaxRows(), 1).setWrap(true);
  h.setColumnWidth(COL.DIA, 90);
  h.setColumnWidth(COL.ID, 60);
  h.setColumnWidth(COL.FECHA, 130);
  h.setColumnWidth(COL.PRODUCTOS, 260);
}

// Migra la hoja de la v1/v2 (columna única "N°") a la numeración doble
function migrarNumeracion_(p) {
  const encabezado = p.getRange(1, 1, 1, 2).getValues()[0];
  const props = PropertiesService.getScriptProperties();
  if (encabezado[0] === 'N°' && encabezado[1] === 'Fecha') {
    p.insertColumnAfter(1);                         // nueva columna B = ID
    const n = p.getLastRow() - 1;
    if (n > 0) {
      const ids = [];
      for (let i = 1; i <= n; i++) ids.push([i]);
      p.getRange(2, COL.ID, n, 1).setValues(ids);
      const ultimaFecha = p.getRange(n + 1, COL.FECHA).getValue();
      const numeroDia = p.getRange(n + 1, COL.DIA).getValue();
      props.setProperties({
        ULTIMO_ID: String(n),
        DIA: ultimaFecha instanceof Date ? Utilities.formatDate(ultimaFecha, CONFIG.ZONA_HORARIA, 'yyyy-MM-dd') : '',
        N_DIA: String(Number(numeroDia) || 0),
      });
    }
    console.log('Migrado a numeración doble (' + n + ' pedidos).');
  }
  // Si el contador no existe (hoja nueva o propiedades borradas), se reconstruye con el ID más alto
  if (!props.getProperty('ULTIMO_ID')) props.setProperty('ULTIMO_ID', String(maximoIdGuardado_()));
}

function maximoIdGuardado_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let max = 0;
  ss.getSheets().forEach(h => {
    const nombre = h.getName();
    if (nombre !== HOJA_PEDIDOS && nombre.indexOf(PREFIJO_HISTORIAL) !== 0) return;
    const n = h.getLastRow() - 1;
    if (n < 1) return;
    h.getRange(2, COL.ID, n, 1).getValues().forEach(r => { max = Math.max(max, Number(r[0]) || 0); });
  });
  return max;
}

function leerContadores_() {
  const p = PropertiesService.getScriptProperties();
  return { ultimoId: Number(p.getProperty('ULTIMO_ID') || 0), dia: p.getProperty('DIA') || '', nDia: Number(p.getProperty('N_DIA') || 0) };
}

// Reserva el siguiente ID único y el siguiente número del día (llamar dentro del lock)
function siguienteNumero_() {
  const props = PropertiesService.getScriptProperties();
  let c = leerContadores_();
  if (!c.ultimoId) c.ultimoId = maximoIdGuardado_();   // red de seguridad
  const hoy = Utilities.formatDate(new Date(), CONFIG.ZONA_HORARIA, 'yyyy-MM-dd');
  const id = c.ultimoId + 1;
  const nDia = c.dia === hoy ? c.nDia + 1 : 1;
  props.setProperties({ ULTIMO_ID: String(id), DIA: hoy, N_DIA: String(nDia) });
  return { id, nDia };
}


/* ---------------------- Web: leer menú ---------------------- */
function doGet(e) {
  const accion = (e && e.parameter && e.parameter.accion) || 'menu';
  if (accion === 'menu') return json({ ok: true, menu: leerMenu() });
  return json({ ok: false, error: 'Acción no válida' });
}

function leerMenu() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_MENU);
  return hoja.getDataRange().getValues().slice(1)
    .filter(f => f[0] !== '' && f[2] !== '' && f[5] === true)   // solo productos marcados como disponibles
    .map(f => ({
      id: String(f[0]).trim(),
      categoria: String(f[1] || 'Otros').trim(),
      nombre: String(f[2]).trim(),
      descripcion: String(f[3] || '').trim(),
      precio: leerPrecio_(f[4]),
      imagen: String(f[6] || '').trim(),
      destacado: f[7] === true,
    }))
    .filter(p => p.precio > 0);   // un plato sin precio válido nunca se ofrece (evita vender en $0)
}

// Acepta 5.99, "5,99", "$ 5,99" o "1.234,50" (por si alguien escribe el precio como texto)
function leerPrecio_(v) {
  if (typeof v === 'number') return Math.round(v * 100) / 100;
  let t = String(v || '').replace(/[^\d.,]/g, '');
  if (t.indexOf(',') !== -1 && t.indexOf('.') !== -1) {
    t = t.lastIndexOf(',') > t.lastIndexOf('.') ? t.replace(/\./g, '').replace(',', '.') : t.replace(/,/g, '');
  } else {
    t = t.replace(',', '.');
  }
  const n = Number(t);
  return isFinite(n) ? Math.round(n * 100) / 100 : 0;
}


/* ---------------------- Web: recibir pedido ---------------------- */
// Error pensado para mostrárselo al cliente tal cual
function fallo_(mensaje) { const e = new Error(mensaje); e.paraCliente = true; throw e; }

function doPost(e) {
  // 1) Leer y validar el pedido FUERA del candado: así varios clientes se validan a la vez
  //    y el candado solo se usa para numerar y guardar (lo más rápido posible).
  let d, clave, pedido;
  try {
    try { d = JSON.parse((e && e.postData && e.postData.contents) || '{}'); }
    catch (x) { fallo_('El pedido llegó incompleto. Recarga la página e intenta de nuevo.'); }
    if (!d || typeof d !== 'object') fallo_('El pedido llegó incompleto. Recarga la página e intenta de nuevo.');
    if (d.website) return json({ ok: true, id: 0, ref: 0, total: 0 });   // trampa anti-bots
    clave = String(d.clave || '').replace(/[^\w-]/g, '').slice(0, 64);
    pedido = validarPedido_(d);
  } catch (err) {
    return respuestaError_(err);
  }

  // 2) Numerar y guardar DENTRO del candado (un pedido a la vez: nunca se repiten números)
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) {
    return json({ ok: false, ocupado: true, error: 'Estamos recibiendo muchos pedidos en este momento. Intenta de nuevo en unos segundos.' });
  }
  let respuesta, nums;
  try {
    // Reintento del mismo pedido (p. ej. se cortó internet justo al enviar): se devuelve el resultado
    // original en lugar de crear un pedido duplicado. La web manda la misma "clave" en cada reintento.
    const cache = CacheService.getScriptCache();
    if (clave) {
      const previo = cache.get('pedido_' + clave);
      if (previo) return json(JSON.parse(previo));
    }
    nums = siguienteNumero_();
    const p = pedido;
    const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_PEDIDOS);
    // setValues (en vez de appendRow) respeta el formato de cada columna (fecha sin segundos, etc.)
    hoja.getRange(hoja.getLastRow() + 1, 1, 1, COLUMNAS_PEDIDOS.length).setValues([[
      nums.nDia, nums.id, new Date(), ESTADOS[0], p.nombre, "'" + p.telefono, p.entrega, p.direccion,
      p.lineas.join('\n'), p.total, p.pago, p.notas, enlaceWhatsApp(p.telefono)]]);
    SpreadsheetApp.flush();
    respuesta = { ok: true, id: nums.nDia, ref: nums.id, total: p.total };
    if (clave) cache.put('pedido_' + clave, JSON.stringify(respuesta), 21600);   // se recuerda 6 horas
  } catch (err) {
    return respuestaError_(err);
  } finally {
    lock.releaseLock();
  }

  // 3) Avisar al equipo FUERA del candado (Telegram/correo no frenan a los demás pedidos).
  //    Si el aviso falla, el pedido ya quedó guardado.
  try { notificar(Object.assign({ id: nums.id, nDia: nums.nDia }, pedido)); }
  catch (err) { console.error('notificar: ' + err); }
  return json(respuesta);
}

// Valida los datos y calcula el total con los precios de la hoja (nunca los del navegador)
function validarPedido_(d) {
  const nombre = limpiar(d.nombre, 60);
  const telefono = String(d.telefono || '').replace(/[^\d+ ]/g, '').trim().slice(0, 20);   // tal cual (con + si lo puso)
  const entrega = d.entrega === 'retiro' ? 'Retiro en local' : 'Delivery';
  const direccion = limpiar(d.direccion, 200);
  const pago = limpiar(d.pago, 40);
  const notas = limpiar(d.notas, 300);

  if (!nombre) fallo_('Escribe tu nombre.');
  if (telefono.replace(/\D/g, '').length < 7) fallo_('Revisa el número de teléfono.');
  if (entrega === 'Delivery' && !direccion) fallo_('Falta la dirección de entrega.');
  if (!Array.isArray(d.items)) fallo_('El pedido llegó incompleto. Recarga la página e intenta de nuevo.');

  const menu = Object.create(null);   // sin prototipo: ids como "constructor" no engañan
  leerMenu().forEach(p => menu[p.id] = p);

  // Suma las cantidades por producto (así nadie salta el máximo repitiendo líneas)
  const cantidades = Object.create(null);
  d.items.slice(0, 100).forEach(it => {
    if (!it || typeof it !== 'object') return;
    const cant = Math.floor(Number(it.cantidad)) || 0;
    if (cant > 0) cantidades[String(it.id)] = (cantidades[String(it.id)] || 0) + cant;
  });
  let total = 0;
  const lineas = [];
  Object.keys(cantidades).forEach(pid => {
    const p = menu[pid];
    if (!p) fallo_('Un producto de tu carrito se agotó. Recarga la página.');
    const cant = Math.min(cantidades[pid], CONFIG.MAX_UNIDADES);
    total += p.precio * cant;
    lineas.push(cant + ' x ' + p.nombre);
  });
  if (!lineas.length) fallo_('El carrito está vacío.');
  total = Math.round(total * 100) / 100;
  return { nombre, telefono, entrega, direccion, pago, notas, lineas, total };
}

function respuestaError_(err) {
  if (err && err.paraCliente) return json({ ok: false, error: err.message });
  console.error('doPost: ' + (err && err.stack || err));   // queda en "Ejecuciones" de Apps Script
  return json({ ok: false, error: 'No pudimos registrar tu pedido. Intenta de nuevo o escríbenos por WhatsApp.' });
}


/* ---------------------- Archivo automático ---------------------- */

// Mueve a "Historial AAAA" los pedidos Entregados/Rechazados de días anteriores.
// Los pedidos de hoy y los que siguen abiertos se quedan en "Pedidos".
function archivarPedidos() {
  const lock = LockService.getScriptLock();
  // que no entre un pedido mientras se reescribe la hoja; si está ocupado, se reintenta mañana
  if (!lock.tryLock(60000)) { console.error('archivarPedidos: la hoja estaba ocupada, se omite hoy.'); return; }
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const p = ss.getSheetByName(HOJA_PEDIDOS);
    const n = p.getLastRow() - 1;
    if (n < 1) return;
    const ancho = COLUMNAS_PEDIDOS.length;
    const filas = p.getRange(2, 1, n, ancho).getValues();
    const hoy = Utilities.formatDate(new Date(), CONFIG.ZONA_HORARIA, 'yyyy-MM-dd');

    const quedan = [], porAnio = {};
    filas.forEach(f => {
      const fecha = f[COL.FECHA - 1];
      const dia = fecha instanceof Date ? Utilities.formatDate(fecha, CONFIG.ZONA_HORARIA, 'yyyy-MM-dd') : hoy;
      const cerrado = ESTADOS_CERRADOS.indexOf(f[COL.ESTADO - 1]) !== -1;
      if (cerrado && dia < hoy) {
        const anio = dia.slice(0, 4);
        (porAnio[anio] = porAnio[anio] || []).push(f);
      } else {
        quedan.push(f);
      }
    });
    const movidos = filas.length - quedan.length;
    if (!movidos) { console.log('Nada que archivar.'); return; }

    // 1) Copiar al historial del año correspondiente
    Object.keys(porAnio).forEach(anio => {
      const nombre = PREFIJO_HISTORIAL + anio;
      let h = ss.getSheetByName(nombre);
      if (!h) {
        h = ss.insertSheet(nombre);
        h.appendRow(COLUMNAS_PEDIDOS);
        darFormatoPedidos_(h);
      }
      const datos = porAnio[anio].map(textoTelefono_);
      h.getRange(h.getLastRow() + 1, 1, datos.length, ancho).setValues(datos);
    });
    SpreadsheetApp.flush();

    // 2) Reescribir "Pedidos" solo con lo que queda
    p.getRange(2, 1, n, ancho).clearContent();
    if (quedan.length) p.getRange(2, 1, quedan.length, ancho).setValues(quedan.map(textoTelefono_));
    console.log('Archivados: ' + movidos + ' · Quedan en Pedidos: ' + quedan.length);
  } finally {
    lock.releaseLock();
  }
}

function textoTelefono_(fila) {
  const f = fila.slice();
  const t = String(f[COL.TELEFONO - 1] || '');
  f[COL.TELEFONO - 1] = t && t[0] !== "'" ? "'" + t : t;
  return f;
}

// Ejecutar UNA vez: programa archivarPedidos todos los días a CONFIG.HORA_ARCHIVO
function instalarArchivoDiario() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'archivarPedidos')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('archivarPedidos').timeBased()
    .everyDays(1).atHour(CONFIG.HORA_ARCHIVO).inTimezone(CONFIG.ZONA_HORARIA).create();
  console.log('Archivo diario programado a las ' + CONFIG.HORA_ARCHIVO + ':00 (' + CONFIG.ZONA_HORARIA + ').');
}


/* ---------------------- Avisos al equipo ---------------------- */
function notificar(p) {
  const texto =
    '🔥 NUEVO PEDIDO #' + p.nDia + '  (ID ' + p.id + ')\n' +
    '👤 ' + p.nombre + ' · ' + p.telefono + '\n' +
    (p.entrega === 'Delivery' ? '🛵 Delivery: ' + p.direccion : '🏪 Retira en el local') + '\n\n' +
    p.lineas.map(l => '• ' + l).join('\n') + '\n\n' +
    '💰 Total: ' + CONFIG.MONEDA + p.total.toFixed(2) + ' · ' + p.pago +
    (p.notas ? '\n📝 ' + p.notas : '') +
    '\n\n💬 ' + enlaceWhatsApp(p.telefono);

  // Si un aviso falla, el pedido YA quedó guardado en la hoja: no se pierde.
  if (CONFIG.TELEGRAM_TOKEN && CONFIG.TELEGRAM_CHAT_ID) {
    try {
      UrlFetchApp.fetch('https://api.telegram.org/bot' + CONFIG.TELEGRAM_TOKEN + '/sendMessage', {
        method: 'post',
        payload: { chat_id: CONFIG.TELEGRAM_CHAT_ID, text: texto, disable_web_page_preview: 'true' },
        muteHttpExceptions: true,
      });
    } catch (e) { console.error('Telegram: ' + e); }
  }
  if (CONFIG.EMAIL_AVISO) {
    try { MailApp.sendEmail(CONFIG.EMAIL_AVISO, 'Nuevo pedido #' + p.nDia + ' (ID ' + p.id + ')', texto); }
    catch (e) { console.error('Correo: ' + e); }
  }
}


/* ---------------------- Utilidades de configuración ---------------------- */

// Paso Telegram: agrega el bot al grupo, escribe /start en el grupo y luego ejecuta esto.
// Mira el "Registro de ejecución": ahí aparece el chat_id (en grupos empieza con -).
function verChatId() {
  const r = UrlFetchApp.fetch('https://api.telegram.org/bot' + CONFIG.TELEGRAM_TOKEN + '/getUpdates',
                              { muteHttpExceptions: true });
  const datos = JSON.parse(r.getContentText());
  if (!datos.ok) { console.log('Token inválido: ' + r.getContentText()); return; }
  const chats = {};
  datos.result.forEach(u => {
    const c = (u.message || u.my_chat_member || u.channel_post || {}).chat;
    if (c) chats[c.id] = c.title || c.first_name;
  });
  console.log(Object.keys(chats).length ? chats : 'Nada aún: escribe /start en el grupo y vuelve a ejecutar.');
}

// Envía un pedido falso a Telegram/correo para comprobar que los avisos llegan.
function probarAviso() {
  notificar({ id: 0, nDia: 0, nombre: 'Prueba', telefono: '04240000000', entrega: 'Retiro en local', direccion: '',
              lineas: ['1 x Producto de prueba'], total: 1, pago: 'Pago Móvil', notas: 'Esto es una prueba' });
}

// Marca como "Rechazado" los pedidos de prueba (cliente que empieza con PRUEBA o ESTRES).
// No borra nada: el archivo de la madrugada los pasará al historial.
function marcarPruebasComoRechazadas() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_PEDIDOS);
  const n = hoja.getLastRow() - 1;
  if (n < 1) return;
  const rango = hoja.getRange(2, COL.ESTADO, n, 2);          // Estado y Cliente
  let marcados = 0;
  const valores = rango.getValues().map(([estado, cliente]) => {
    if (/^(PRUEBA|ESTRES)/i.test(String(cliente)) && estado !== 'Rechazado') { marcados++; return ['Rechazado', cliente]; }
    return [estado, cliente];
  });
  rango.setValues(valores);
  console.log('Pedidos de prueba marcados como Rechazado: ' + marcados);
}

// Muestra en el registro los contadores actuales (útil para revisar)
function verContadores() { console.log(leerContadores_()); }


/* ---------------------- Internas ---------------------- */
function limpiar(v, max) {
  // Quita =,+,-,@ iniciales para que nadie pueda meter fórmulas en la hoja
  return String(v || '').trim().replace(/^[=+\-@]+/, '').slice(0, max);
}

function enlaceWhatsApp(tel) {
  // Convierte lo que escribió el cliente en un enlace wa.me con código de país.
  // Reconoce números de Venezuela (58) y Colombia (57), y cualquier número escrito con +.
  const texto = String(tel || '').trim();
  const n = texto.replace(/\D/g, '');
  if (!n) return '';
  if (texto.startsWith('+')) return 'https://wa.me/' + n;           // ya viene con código de país
  if (n.startsWith('00')) return 'https://wa.me/' + n.slice(2);      // 0057..., 0058...
  if (/^(58|57)\d{10}$/.test(n)) return 'https://wa.me/' + n;        // 58424..., 57300...
  if (/^0\d{10}$/.test(n)) return 'https://wa.me/58' + n.slice(1);   // Venezuela: 0424 1234567
  if (/^4\d{9}$/.test(n)) return 'https://wa.me/58' + n;             // Venezuela sin el 0: 424 1234567
  if (/^3\d{9}$/.test(n)) return 'https://wa.me/57' + n;             // Colombia: 300 1234567
  if (n.length >= 11) return 'https://wa.me/' + n;                   // otro país, ya con su código
  return 'https://wa.me/' + CONFIG.CODIGO_PAIS + n;
}

// Recalcula la columna "WhatsApp cliente" de todos los pedidos (útil tras corregir esta función).
function repararEnlacesWhatsApp() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_PEDIDOS);
  const filas = hoja.getLastRow() - 1;
  if (filas < 1) return;
  const tels = hoja.getRange(2, COL.TELEFONO, filas, 1).getValues();
  hoja.getRange(2, COL.WHATSAPP, filas, 1).setValues(tels.map(t => [enlaceWhatsApp(t[0])]));
  console.log('Enlaces actualizados: ' + filas);
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
