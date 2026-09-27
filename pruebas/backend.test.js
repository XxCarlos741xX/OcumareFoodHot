const { crearEntorno } = require('./gas-mock');
const resultados = [];
function prueba(nombre, fn) {
  try { fn(); resultados.push({ nombre, ok: true }); }
  catch (e) { resultados.push({ nombre, ok: false, error: e.message }); }
}
const esperar = (cond, msg) => { if (!cond) throw new Error(msg); };
const igual = (a, b, msg) => { if (JSON.stringify(a) !== JSON.stringify(b)) throw new Error(`${msg}: esperado ${JSON.stringify(b)}, obtenido ${JSON.stringify(a)}`); };

const base = { nombre: 'Ana', telefono: '0424 1234567', entrega: 'delivery', direccion: 'Calle 1', pago: 'Pago Móvil', notas: '', website: '' };
function nuevo(ahora = '2026-09-27T15:00:00Z') { const g = crearEntorno({ ahora }); g.configurarHojas(); return g; }
const filas = g => g.hoja('Pedidos').getDataRange().getValues().slice(1);

// ---------- Instalación ----------
prueba('B01 configurarHojas crea Pedidos y Menu, borra "Hoja 1"', () => {
  const g = nuevo();
  igual(g.libro.getSheets().map(h => h.getName()).sort(), ['Menu', 'Pedidos'], 'hojas');
  igual(g.hoja('Pedidos').getRange(1, 1, 1, 13).getValues()[0][0], 'N° del día', 'encabezado');
});
prueba('B02 configurarHojas es idempotente (2 ejecuciones no duplican menú ni columnas)', () => {
  const g = nuevo(); g.configurarHojas();
  igual(g.leerMenu().length, 50, 'platos'); igual(g.hoja('Pedidos').getRange(1, 1, 1, 13).getValues()[0][1], 'ID', 'col B');
});

// ---------- Menú ----------
prueba('B03 GET menú devuelve 50 platos y 5 destacados', () => {
  const g = nuevo(); const r = g.get({ accion: 'menu' });
  igual(r.ok, true, 'ok'); igual(r.menu.length, 50, 'platos'); igual(r.menu.filter(p => p.destacado).length, 5, 'destacados');
});
prueba('B04 plato con "disponible" desmarcado no aparece', () => {
  const g = nuevo(); g.hoja('Menu').escribir(2, 6, false);
  esperar(!g.leerMenu().some(p => p.id === 'HB1'), 'HB1 sigue apareciendo');
});
prueba('B05 acción GET desconocida responde error controlado', () => { igual(nuevo().get({ accion: 'x' }).ok, false, 'ok'); });
prueba('B06 precio escrito con coma ("5,99") no queda en $0', () => {
  const g = nuevo(); g.hoja('Menu').escribir(2, 5, '5,99');
  const p = g.leerMenu().find(x => x.id === 'HB1');
  esperar(p && p.precio === 5.99, 'precio leído: ' + (p && p.precio));
});
prueba('B07 plato sin precio (0 o vacío) no se ofrece ni se puede pedir', () => {
  const g = nuevo(); g.hoja('Menu').escribir(2, 5, '');
  esperar(!g.leerMenu().some(p => p.id === 'HB1'), 'se ofrece un plato gratis');
});

// ---------- Pedidos: cálculo y validación ----------
prueba('B08 pedido válido: total calculado con precios de la hoja', () => {
  const g = nuevo(); const r = g.post({ ...base, items: [{ id: 'PE1', cantidad: 1 }, { id: 'PP5', cantidad: 2 }, { id: 'HB4', cantidad: 1 }] });
  igual(r, { ok: true, id: 1, ref: 1, total: 29.99 }, 'respuesta');
});
prueba('B09 manipular el precio desde el navegador no tiene efecto', () => {
  const g = nuevo(); const r = g.post({ ...base, total: 0.01, items: [{ id: 'PA3', cantidad: 1, precio: 0.01 }] });
  igual(r.total, 28, 'total');
});
prueba('B10 cantidades: 999→30, negativa ignorada, 2.7→2, "3"→3', () => {
  const g = nuevo(); const r = g.post({ ...base, items: [{ id: 'PP1', cantidad: 999 }, { id: 'PP2', cantidad: -5 }, { id: 'PP3', cantidad: 2.7 }, { id: 'PP4', cantidad: '3' }] });
  igual(r.total, Math.round((30 * 2.5 + 2 * 4.99 + 3 * 4.99) * 100) / 100, 'total');
});
prueba('B11 producto inexistente → error claro', () => { const r = nuevo().post({ ...base, items: [{ id: 'NOEXISTE', cantidad: 1 }] }); igual(r.ok, false, 'ok'); esperar(/agotó/.test(r.error), r.error); });
prueba('B11b ids engañosos (constructor, __proto__, toString) → error, nunca total NaN', () => { const g = nuevo(); ['constructor','__proto__','toString'].forEach(id => { const r = g.post({ ...base, items: [{ id, cantidad: 1 }] }); igual(r.ok, false, id); }); igual(filas(g).length, 0, 'filas'); });
prueba('B12 carrito vacío → error', () => { igual(nuevo().post({ ...base, items: [] }).ok, false, 'ok'); });
prueba('B13 items con formato inválido (texto) → error controlado en español', () => {
  const r = nuevo().post({ ...base, items: 'hackeo' }); igual(r.ok, false, 'ok'); esperar(!/is not a function/.test(r.error), 'mensaje técnico: ' + r.error);
});
prueba('B14 sin nombre → error', () => { igual(nuevo().post({ ...base, nombre: '  ', items: [{ id: 'PP1', cantidad: 1 }] }).ok, false, 'ok'); });
prueba('B15 teléfono corto → error', () => { igual(nuevo().post({ ...base, telefono: '123', items: [{ id: 'PP1', cantidad: 1 }] }).ok, false, 'ok'); });
prueba('B16 delivery sin dirección → error; retiro sin dirección → OK', () => {
  const g = nuevo();
  igual(g.post({ ...base, direccion: '', items: [{ id: 'PP1', cantidad: 1 }] }).ok, false, 'delivery');
  igual(g.post({ ...base, entrega: 'retiro', direccion: '', items: [{ id: 'PP1', cantidad: 1 }] }).ok, true, 'retiro');
});
prueba('B17 cuerpo que no es JSON → error controlado en español', () => { const r = nuevo().post('esto no es json'); igual(r.ok, false, 'ok'); esperar(/incompleto/.test(r.error), r.error); });
prueba('B17b items con elementos nulos o 5.000 líneas → no se cae', () => { const g = nuevo(); igual(g.post({ ...base, items: [null, 5, { id: 'PP1', cantidad: 1 }] }).total, 2.5, 'nulos'); igual(g.post({ ...base, items: Array(5000).fill({ id: 'PP1', cantidad: 1 }) }).total, 75, 'muchas líneas: tope de 30 por producto'); });
prueba('B06b precios "$ 1.234,50" y 7 (número) se leen bien', () => { const g = nuevo(); g.hoja('Menu').escribir(2, 5, '$ 1.234,50'); igual(g.leerMenu().find(x => x.id === 'HB1').precio, 1234.5, 'texto'); igual(g.leerMenu().find(x => x.id === 'PE1').precio, 7, 'número'); });
prueba('B18 trampa anti-bots: responde OK pero NO guarda', () => {
  const g = nuevo(); const r = g.post({ ...base, website: 'spam.com', items: [{ id: 'PP1', cantidad: 1 }] });
  igual(r.ok, true, 'ok'); igual(filas(g).length, 0, 'filas guardadas');
});
prueba('B19 inyección de fórmulas neutralizada (=, +, -, @)', () => {
  const g = nuevo(); g.post({ ...base, nombre: '=IMPORTXML("http://x")', direccion: '+cmd', notas: '@SUM(1)', items: [{ id: 'PP1', cantidad: 1 }] });
  const f = filas(g)[0];
  esperar(![f[4], f[7], f[11]].some(v => /^[=+\-@]/.test(String(v))), 'quedó texto peligroso: ' + [f[4], f[7], f[11]]);
});
prueba('B20 textos largos se recortan (nombre 60, dirección 200, notas 300)', () => {
  const g = nuevo(); g.post({ ...base, nombre: 'x'.repeat(500), direccion: 'y'.repeat(500), notas: 'z'.repeat(900), items: [{ id: 'PP1', cantidad: 1 }] });
  const f = filas(g)[0]; igual([f[4].length, f[7].length, f[11].length], [60, 200, 300], 'longitudes');
});
prueba('B21 el pedido guarda las 13 columnas en su lugar', () => {
  const g = nuevo(); g.post({ ...base, items: [{ id: 'PP1', cantidad: 2 }] });
  const f = filas(g)[0];
  igual([f[0], f[1], f[3], f[4], f[5], f[6], f[8], f[9], f[12]], [1, 1, 'Nuevo', 'Ana', '0424 1234567', 'Delivery', '2 x Perro normal', 5, 'https://wa.me/584241234567'], 'fila');
  esperar(f[2] instanceof Date || (f[2] && f[2].getTime), 'fecha no es fecha');
});
prueba('B22 se envía aviso por Telegram/correo solo si está configurado', () => {
  const g = nuevo(); g.post({ ...base, items: [{ id: 'PP1', cantidad: 1 }] });
  igual(g.registro.fetch.length + g.registro.correos.length, 0, 'avisos sin configurar');
  g.CONFIG.TELEGRAM_TOKEN = 'T'; g.CONFIG.TELEGRAM_CHAT_ID = '-1'; g.CONFIG.EMAIL_AVISO = 'a@b.c';
  g.post({ ...base, items: [{ id: 'PP1', cantidad: 1 }] });
  igual([g.registro.fetch.length, g.registro.correos.length], [1, 1], 'avisos configurados');
  esperar(/NUEVO PEDIDO #2 \(ID 2\)/.test(g.registro.fetch[0][1].payload.text.replace(/\s+/g, ' ')), 'texto del aviso');
});
prueba('B23 si Google tarda (lock ocupado) el cliente recibe mensaje en español', () => {
  const g = crearEntorno({ ahora: '2026-09-27T15:00:00Z', lockFalla: true }); g.configurarHojas();
  const r = g.post({ ...base, items: [{ id: 'PP1', cantidad: 1 }] });
  igual(r.ok, false, 'ok'); esperar(!/Lock timeout/i.test(r.error), 'mensaje técnico en inglés: ' + r.error);
});

prueba('B23b reintento con la misma clave NO duplica el pedido; clave distinta sí crea otro', () => {
  const g = nuevo(); const pedido = { ...base, clave: 'abc-123', items: [{ id: 'PP1', cantidad: 1 }] };
  const a = g.post(pedido), b = g.post(pedido), c = g.post({ ...pedido, clave: 'otra' });
  igual([a.ref, b.ref, c.ref], [1, 1, 2], 'refs'); igual(filas(g).length, 2, 'filas');
});

// ---------- Numeración ----------
prueba('B24 IDs y N° del día consecutivos', () => {
  const g = nuevo(); const ids = [];
  for (let i = 0; i < 5; i++) ids.push(g.post({ ...base, items: [{ id: 'PP1', cantidad: 1 }] }));
  igual(ids.map(r => [r.id, r.ref]), [[1, 1], [2, 2], [3, 3], [4, 4], [5, 5]], 'numeración');
});
prueba('B25 el N° del día se reinicia a medianoche de Venezuela; el ID sigue', () => {
  const g = nuevo('2026-09-27T03:30:00Z');        // 23:30 del 26 en Caracas
  g.post({ ...base, items: [{ id: 'PP1', cantidad: 1 }] }); g.post({ ...base, items: [{ id: 'PP1', cantidad: 1 }] });
  g.setAhora('2026-09-27T04:10:00Z');              // 00:10 del 27 en Caracas
  const r = g.post({ ...base, items: [{ id: 'PP1', cantidad: 1 }] });
  igual([r.id, r.ref], [1, 3], 'nuevo día');
});
prueba('B26 borrar filas no repite IDs', () => {
  const g = nuevo(); g.post({ ...base, items: [{ id: 'PP1', cantidad: 1 }] }); g.post({ ...base, items: [{ id: 'PP1', cantidad: 1 }] });
  g.hoja('Pedidos').getRange(2, 1, 2, 13).clearContent();
  igual(g.post({ ...base, items: [{ id: 'PP1', cantidad: 1 }] }).ref, 3, 'ID tras borrar');
});
prueba('B27 si se pierden los contadores, se reconstruyen desde el ID más alto guardado', () => {
  const g = nuevo(); for (let i = 0; i < 3; i++) g.post({ ...base, items: [{ id: 'PP1', cantidad: 1 }] });
  Object.keys(g.props).forEach(k => delete g.props[k]);
  igual(g.post({ ...base, items: [{ id: 'PP1', cantidad: 1 }] }).ref, 4, 'ID reconstruido');
});
prueba('B28 migración desde la versión anterior (columna "N°")', () => {
  const g = crearEntorno({ ahora: '2026-09-27T02:00:00Z' });
  const p = g.libro.insertSheet('Pedidos');
  p.appendRow(['N°', 'Fecha', 'Estado', 'Cliente', 'Teléfono', 'Entrega', 'Dirección', 'Productos', 'Total', 'Pago', 'Notas', 'WhatsApp cliente']);
  p.appendRow([1, g.fecha('2026-09-27T01:00:00Z'), 'Nuevo', 'A', "'0424", 'Delivery', 'x', '1 x P', 2.5, 'Pago', '', 'l']);
  p.appendRow([2, g.fecha('2026-09-27T01:30:00Z'), 'Nuevo', 'B', "'0424", 'Delivery', 'x', '1 x P', 2.5, 'Pago', '', 'l']);
  g.configurarHojas();
  igual(p.getRange(2, 1, 2, 4).getValues().map(f => [f[0], f[1], f[3]]), [[1, 1, 'Nuevo'], [2, 2, 'Nuevo']], 'filas migradas');
  const r = g.post({ ...base, items: [{ id: 'PP1', cantidad: 1 }] });
  igual([r.id, r.ref], [3, 3], 'sigue la numeración');
});

// ---------- Archivo diario ----------
function sembrar(g) {
  const p = g.hoja('Pedidos');
  const fila = (dia, id, fechaIso, estado) => [dia, id, g.fecha(fechaIso), estado, 'C' + id, "'0424 000000" + id, 'Delivery', 'x', '1 x P', 2.5, 'Pago', '', 'l'];
  [fila(1, 1, '2025-12-31T20:00:00Z', 'Entregado'),   // año anterior, cerrado
   fila(1, 2, '2026-09-26T15:00:00Z', 'Rechazado'),   // ayer, cerrado
   fila(2, 3, '2026-09-26T16:00:00Z', 'Preparando'),  // ayer, abierto (debe quedar)
   fila(1, 4, '2026-09-27T14:00:00Z', 'Entregado'),   // hoy, cerrado (debe quedar)
   fila(2, 5, '2026-09-27T14:30:00Z', 'Nuevo')]       // hoy, abierto
    .forEach(f => p.appendRow(f));
}
prueba('B29 archivo: mueve solo cerrados de días anteriores, separa por año', () => {
  const g = nuevo('2026-09-27T15:00:00Z'); sembrar(g); g.archivarPedidos();
  igual(filas(g).map(f => f[1]), [3, 4, 5], 'quedan en Pedidos');
  igual(g.hoja('Historial 2025').getDataRange().getValues().slice(1).map(f => f[1]), [1], 'Historial 2025');
  igual(g.hoja('Historial 2026').getDataRange().getValues().slice(1).map(f => f[1]), [2], 'Historial 2026');
});
prueba('B30 archivo: el teléfono sigue como texto (con cero inicial)', () => {
  const g = nuevo('2026-09-27T15:00:00Z'); sembrar(g); g.archivarPedidos();
  igual(g.hoja('Historial 2026').getDataRange().getValues()[1][5], '0424 0000002', 'teléfono');
  igual(filas(g)[0][5], '0424 0000003', 'teléfono que queda');
});
prueba('B31 archivo: ejecutarlo 2 veces no duplica nada', () => {
  const g = nuevo('2026-09-27T15:00:00Z'); sembrar(g); g.archivarPedidos(); g.archivarPedidos();
  igual(g.hoja('Historial 2026').getLastRow(), 2, 'filas en historial');
});
prueba('B32 archivo + nuevos pedidos: los IDs no se repiten tras archivar', () => {
  const g = nuevo('2026-09-27T15:00:00Z'); sembrar(g); g.props.ULTIMO_ID = '5'; g.archivarPedidos();
  igual(g.post({ ...base, items: [{ id: 'PP1', cantidad: 1 }] }).ref, 6, 'siguiente ID');
});
prueba('B33 instalarArchivoDiario deja exactamente 1 disparador aunque se ejecute 2 veces', () => {
  const g = nuevo(); g.instalarArchivoDiario(); g.instalarArchivoDiario();
  igual(g.registro.triggers.filter(t => t.fn === 'archivarPedidos').length, 1, 'disparadores');
});

prueba('B33b marcarPruebasComoRechazadas solo toca pedidos de prueba', () => {
  const g = nuevo();
  ['PRUEBA 1', 'Ana', 'ESTRES 7'].forEach(n => g.post({ ...base, nombre: n, items: [{ id: 'PP1', cantidad: 1 }] }));
  g.marcarPruebasComoRechazadas();
  igual(filas(g).map(f => f[3]), ['Rechazado', 'Nuevo', 'Rechazado'], 'estados');
});

// ---------- WhatsApp ----------
prueba('B34 enlaces de WhatsApp (VE, CO, internacionales)', () => {
  const g = nuevo();
  const casos = { '0412 1234567': '584121234567', '0414-123-4567': '584141234567', '0416 1234567': '584161234567',
    '0424 1234567': '584241234567', '0426 1234567': '584261234567', '424 1234567': '584241234567',
    '+58 412 1234567': '584121234567', '3001234567': '573001234567', '+57 300 123 4567': '573001234567',
    '0057 3001234567': '573001234567', '+1 305 555 1234': '13055551234', '': '' };
  Object.entries(casos).forEach(([t, esp]) => igual(g.enlaceWhatsApp(t), esp ? 'https://wa.me/' + esp : '', t));
});

const fallos = resultados.filter(r => !r.ok);
resultados.forEach(r => console.log((r.ok ? '✅' : '❌') + ' ' + r.nombre + (r.ok ? '' : '\n     → ' + r.error)));
console.log(`\n${resultados.length - fallos.length}/${resultados.length} pruebas del backend OK`);
process.exitCode = fallos.length ? 1 : 0;
