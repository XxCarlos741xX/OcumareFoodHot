// Pruebas de la web con Playwright. La API de Google se simula para controlar cada escenario.
const { chromium } = require('playwright');
const path = require('path');
const { crearEntorno } = require('./gas-mock');
const URL_PAGINA = 'file://' + path.resolve(__dirname, '../index.html');

const resultados = [];
const esperar = (c, m) => { if (!c) throw new Error(m); };

// Backend simulado real (mismo Code.gs) para que las respuestas sean las de verdad
function backend() { const g = crearEntorno({ ahora: '2026-09-27T15:00:00Z' }); g.configurarHojas(); return g; }

async function abrir(browser, { vp = { width: 390, height: 844 }, api = {}, sinStorage = false, g = backend() } = {}) {
  const ctx = await browser.newContext({ viewport: vp });
  if (sinStorage) await ctx.addInitScript(() => { Object.defineProperty(window, 'localStorage', { get() { throw new Error('bloqueado'); } }); });
  const p = await ctx.newPage();
  p.errores = []; p.dialogos = 0; p.posts = [];
  p.on('pageerror', e => p.errores.push(e.message));
  p.on('console', m => { if (m.type() === 'error' && !/ERR_|Failed to load resource/.test(m.text())) p.errores.push(m.text()); });
  p.on('dialog', d => { p.dialogos++; d.dismiss(); });
  await ctx.route('**/fonts.googleapis.com/**', r => r.abort());
  await ctx.route('**/fonts.gstatic.com/**', r => r.abort());
  await ctx.route('**/script.google.com/**', async r => {
    const req = r.request();
    if (req.method() === 'GET') {
      if (api.menuFalla) return r.abort();
      if (api.menuLento) await new Promise(s => setTimeout(s, api.menuLento));
      const cuerpo = api.menu ? { ok: true, menu: api.menu } : g.get({ accion: 'menu' });
      return r.fulfill({ status: 200, contentType: 'application/json', body: JSON.stringify(cuerpo) });
    }
    p.posts.push(req.postData());
    if (api.postFalla) return r.abort();
    if (api.postLento) await new Promise(s => setTimeout(s, api.postLento));
    const resp = api.postRespuesta || g.post(req.postData());
    return r.fulfill({ status: 200, contentType: 'application/json', body: JSON.stringify(resp) });
  });
  await p.goto(URL_PAGINA);
  await p.waitForTimeout(api.menuLento ? 200 : 600);
  return p;
}
const agregar = (p, id, veces = 1) => p.evaluate(([id, n]) => { for (let i = 0; i < n; i++) document.querySelector(`.grilla [data-ctr="${id}"] button[data-d="1"]:last-of-type`).click(); }, [id, veces]);
async function llenar(p, o = {}) {
  await p.click('#btn-barra');
  await p.fill('input[name=nombre]', o.nombre ?? 'Ana');
  await p.fill('input[name=telefono]', o.tel ?? '0424 1234567');
  if (o.retiro) await p.check('input[value=retiro]'); else await p.fill('textarea[name=direccion]', o.dir ?? 'Calle 1');
}

async function prueba(browser, nombre, fn) {
  try { await fn(); resultados.push({ nombre, ok: true }); }
  catch (e) { resultados.push({ nombre, ok: false, error: e.message.split('\n')[0] }); }
}

(async () => {
  const b = await chromium.launch();

  await prueba(b, 'F01 carga el menú desde la API: 50 platos, 9 categorías, 5 destacados', async () => {
    const p = await abrir(b);
    const r = await p.evaluate(() => ({ platos: document.querySelectorAll('.grilla .producto').length, cats: document.querySelectorAll('nav.categorias a').length, top: document.querySelectorAll('#destacados .top').length }));
    esperar(r.platos === 50 && r.cats === 9 && r.top === 5, JSON.stringify(r));
    esperar(!p.errores.length, p.errores.join(' | '));
  });

  await prueba(b, 'F02 si la API no responde: mensaje claro con enlace a WhatsApp, sin errores JS', async () => {
    const p = await abrir(b, { api: { menuFalla: true } });
    const t = await p.textContent('#menu');
    esperar(/No pudimos cargar el menú/.test(t), t);
    esperar(await p.$('#menu a[href*="wa.me"]'), 'sin enlace WhatsApp');
    esperar(!p.errores.length, p.errores.join(' | '));
  });

  await prueba(b, 'F03 mientras carga (API lenta) se ve "Cargando el menú…"', async () => {
    const p = await abrir(b, { api: { menuLento: 3000 } });
    esperar(/Cargando/.test(await p.textContent('#menu')), 'no muestra cargando');
  });

  await prueba(b, 'F04 carrito: sumar, restar, total y barra correctos', async () => {
    const p = await abrir(b);
    await agregar(p, 'PP1', 3); await agregar(p, 'HB4', 1);
    await p.click('.grilla [data-ctr="PP1"] button.menos');
    const r = await p.evaluate(() => [document.getElementById('barra-cant').textContent, document.getElementById('barra-total').textContent]);
    esperar(r[0].includes('(3)') && r[1] === '$13.99', r.join(' '));
  });

  await prueba(b, 'F05 tope de 30 unidades por plato aunque se pulse 50 veces', async () => {
    const p = await abrir(b); await agregar(p, 'PP1', 50);
    esperar((await p.textContent('.grilla [data-ctr="PP1"] span')) === '30', 'no respeta el tope');
  });

  await prueba(b, 'F06 el carrito sobrevive a recargar la página', async () => {
    const p = await abrir(b); await agregar(p, 'PE1', 2);
    await p.reload(); await p.waitForTimeout(600);
    esperar((await p.textContent('#barra-total')) === '$14', await p.textContent('#barra-total'));
  });

  await prueba(b, 'F07 plato que se agota: sale del carrito al recargar', async () => {
    const g = backend();
    const p = await abrir(b, { g }); await agregar(p, 'HB1', 1); await agregar(p, 'PE1', 1);
    g.hoja('Menu').escribir(2, 6, false);   // HB1 agotado
    await p.reload(); await p.waitForTimeout(600);
    esperar((await p.textContent('#barra-total')) === '$7', await p.textContent('#barra-total'));
  });

  await prueba(b, 'F08 pedido completo: envía solo id+cantidad, muestra #N y referencia, vacía el carrito', async () => {
    const p = await abrir(b); await agregar(p, 'PE1', 1); await agregar(p, 'PP5', 2);
    await llenar(p); await p.click('#btn-enviar'); await p.waitForSelector('.exito');
    const enviado = JSON.parse(p.posts[0]);
    esperar(enviado.items.every(i => Object.keys(i).sort().join() === 'cantidad,id') && enviado.clave, 'payload: ' + p.posts[0]);
    const t = await p.textContent('.exito');
    esperar(/#1/.test(t) && /Referencia: 1/.test(t) && /\$21/.test(t), t.replace(/\s+/g, ' '));
    esperar(!(await p.evaluate(() => document.getElementById('barra').classList.contains('visible'))), 'carrito no se vació');
  });

  await prueba(b, 'F09 doble clic en "Enviar" no crea dos pedidos', async () => {
    const p = await abrir(b, { api: { postLento: 1500 } }); await agregar(p, 'PE1', 1);
    await llenar(p);
    await p.evaluate(() => { const b = document.getElementById('btn-enviar'); b.click(); b.click(); b.click(); document.getElementById('form').requestSubmit?.(); });
    await p.waitForTimeout(2500);
    esperar(p.posts.length === 1, 'pedidos enviados: ' + p.posts.length);
  });

  await prueba(b, 'F10 error del servidor: se muestra, el botón se reactiva y el carrito se conserva', async () => {
    const p = await abrir(b, { api: { postRespuesta: { ok: false, error: 'Un producto de tu carrito se agotó. Recarga la página.' } } });
    await agregar(p, 'PE1', 1); await llenar(p); await p.click('#btn-enviar'); await p.waitForTimeout(500);
    esperar(/agotó/.test(await p.textContent('#error')), 'sin mensaje');
    esperar(!(await p.isDisabled('#btn-enviar')), 'botón sigue bloqueado');
    esperar((await p.textContent('#barra-total')) === '$7', 'carrito perdido');
  });

  await prueba(b, 'F11 sin internet al enviar: mensaje amigable (no técnico)', async () => {
    const p = await abrir(b, { api: { postFalla: true } });
    await agregar(p, 'PE1', 1); await llenar(p); await p.click('#btn-enviar'); await p.waitForTimeout(500);
    const t = await p.textContent('#error');
    esperar(t && !/Failed to fetch|TypeError|NetworkError/i.test(t), 'mensaje: ' + t);
  });

  await prueba(b, 'F11b reintento tras corte de red usa la misma clave (no duplica)', async () => {
    const g = backend(); let primero = true;
    const p = await abrir(b, { g, api: {} });
    await p.context().unroute('**/script.google.com/**');
    await p.context().route('**/script.google.com/**', async r => {
      if (r.request().method() === 'GET') return r.fulfill({ status: 200, contentType: 'application/json', body: JSON.stringify(g.get({ accion: 'menu' })) });
      p.posts.push(r.request().postData());
      const resp = g.post(r.request().postData());          // el servidor SÍ guarda…
      if (primero) { primero = false; return r.abort(); }    // …pero la respuesta se pierde
      return r.fulfill({ status: 200, contentType: 'application/json', body: JSON.stringify(resp) });
    });
    await agregar(p, 'PE1', 1); await llenar(p); await p.click('#btn-enviar'); await p.waitForTimeout(500);
    await p.click('#btn-enviar'); await p.waitForSelector('.exito');
    const claves = p.posts.map(x => JSON.parse(x).clave);
    esperar(claves.length === 2 && claves[0] === claves[1], 'claves: ' + claves);
    esperar(g.hoja('Pedidos').getLastRow() === 2, 'filas en la hoja: ' + (g.hoja('Pedidos').getLastRow() - 1));
  });

  await prueba(b, 'F11c si el local está saturado, reintenta solo y el pedido entra', async () => {
    const g = backend(); let llamadas = 0;
    const p = await abrir(b, { g });
    await p.context().unroute('**/script.google.com/**');
    await p.context().route('**/script.google.com/**', async r => {
      if (r.request().method() === 'GET') return r.fulfill({ status: 200, contentType: 'application/json', body: JSON.stringify(g.get({ accion: 'menu' })) });
      llamadas++;
      const resp = llamadas === 1 ? { ok: false, ocupado: true, error: 'Estamos recibiendo muchos pedidos' } : g.post(r.request().postData());
      return r.fulfill({ status: 200, contentType: 'application/json', body: JSON.stringify(resp) });
    });
    await agregar(p, 'PE1', 1); await llenar(p); await p.click('#btn-enviar');
    await p.waitForSelector('.exito', { timeout: 10000 });
    esperar(llamadas === 2, 'llamadas: ' + llamadas);
  });

  await prueba(b, 'F12 formulario: sin nombre no envía; "Retiro" oculta y no exige dirección', async () => {
    const p = await abrir(b); await agregar(p, 'PE1', 1);
    await llenar(p, { nombre: '' }); await p.click('#btn-enviar'); await p.waitForTimeout(300);
    esperar(p.posts.length === 0, 'envió sin nombre');
    await p.fill('input[name=nombre]', 'Ana'); await p.fill('textarea[name=direccion]', ''); await p.check('input[value=retiro]');
    esperar(!(await p.isVisible('#campo-direccion')), 'dirección visible');
    await p.click('#btn-enviar'); await p.waitForSelector('.exito');
  });

  await prueba(b, 'F13 seguridad: HTML/JS dentro del menú se muestra como texto (sin XSS)', async () => {
    const malo = [{ id: 'X1', categoria: '<b>Cat</b>', nombre: '<img src=x onerror=alert(1)>', descripcion: '<script>alert(2)</script>', precio: 1, imagen: 'x" onerror="alert(3)', destacado: true }];
    const p = await abrir(b, { api: { menu: malo } });
    await agregar(p, 'X1', 1); await p.click('#btn-barra'); await p.waitForTimeout(300);
    esperar(p.dialogos === 0, 'se ejecutó código: ' + p.dialogos + ' alertas');
    esperar((await p.textContent('.grilla h3')).includes('<img'), 'no se muestra como texto');
  });

  await prueba(b, 'F14 id de producto con comillas no rompe los botones', async () => {
    const p = await abrir(b, { api: { menu: [{ id: "A'1\"x", categoria: 'C', nombre: 'Raro', descripcion: '', precio: 2, imagen: '', destacado: false }] } });
    await p.click('.grilla .contador button'); await p.waitForTimeout(100);
    esperar((await p.textContent('#barra-total')) === '$2', 'no agregó');
    esperar(!p.errores.length, p.errores.join(' | '));
  });

  await prueba(b, 'F15 con el almacenamiento del navegador bloqueado la web sigue funcionando', async () => {
    const p = await abrir(b, { sinStorage: true }); await agregar(p, 'PE1', 1);
    await llenar(p); await p.click('#btn-enviar'); await p.waitForSelector('.exito');
    esperar(!p.errores.length, p.errores.join(' | '));
  });

  await prueba(b, 'F16 navegación por categorías llega a cada sección', async () => {
    const p = await abrir(b);
    for (const c of ['cat-almuerzos', 'cat-pollo', 'cat-hamburguesas']) {
      await p.click(`nav.categorias a[href="#${c}"]`); await p.waitForTimeout(1300);
      const top = await p.evaluate(c => document.getElementById(c).getBoundingClientRect().top, c);
      esperar(top >= 40 && top <= 90, c + ' quedó en ' + top);
    }
  });

  await prueba(b, 'F17 carrusel: flechas avanzan y retroceden (computador)', async () => {
    const p = await abrir(b, { vp: { width: 1366, height: 800 } });
    await p.click('#flecha-der'); await p.waitForTimeout(700);
    const x = await p.evaluate(() => document.getElementById('destacados').scrollLeft);
    await p.click('#flecha-izq'); await p.waitForTimeout(700);
    const y = await p.evaluate(() => document.getElementById('destacados').scrollLeft);
    esperar(x > 100 && y < 5, `avanzó ${x}, volvió ${y}`);
  });

  for (const w of [320, 390, 768, 1366]) {
    await prueba(b, `F18 sin desbordes horizontales a ${w}px`, async () => {
      const p = await abrir(b, { vp: { width: w, height: 800 } }); await agregar(p, 'HB4', 1); await p.click('#btn-barra'); await p.waitForTimeout(200);
      const r = await p.evaluate(() => [document.documentElement.scrollWidth, innerWidth]);
      esperar(r[0] <= r[1], `ancho ${r[0]} > ${r[1]}`);
    });
  }

  await prueba(b, 'F19 accesibilidad básica: lang, alt en imágenes, botones con nombre, labels', async () => {
    const p = await abrir(b); await agregar(p, 'PE1', 1); await p.click('#btn-barra');
    const r = await p.evaluate(() => ({
      lang: document.documentElement.lang,
      sinAlt: [...document.images].filter(i => !i.hasAttribute('alt')).length,
      botonesSinNombre: [...document.querySelectorAll('button')].filter(b => !(b.textContent.trim() || b.getAttribute('aria-label'))).length,
      camposSinLabel: [...document.querySelectorAll('input:not(.oculto),select,textarea')].filter(i => !i.closest('label')).length,
    }));
    esperar(r.lang === 'es' && !r.sinAlt && !r.botonesSinNombre && !r.camposSinLabel, JSON.stringify(r));
  });

  await prueba(b, 'F20 tecla Escape cierra el panel del pedido', async () => {
    const p = await abrir(b); await agregar(p, 'PE1', 1); await p.click('#btn-barra');
    await p.keyboard.press('Escape'); await p.waitForTimeout(150);
    esperar(!(await p.evaluate(() => document.getElementById('panel-fondo').classList.contains('abierto'))), 'no se cierra');
  });

  await prueba(b, 'F21 peso de la carga inicial en celular < 1,5 MB', async () => {
    const ctx = await b.newContext({ viewport: { width: 390, height: 844 } });
    const p = await ctx.newPage(); let bytes = 0;
    await ctx.route('**/fonts.*/**', r => r.abort());
    await ctx.route('**/script.google.com/**', r => r.fulfill({ status: 200, contentType: 'application/json', body: JSON.stringify(backend().get({ accion: 'menu' })) }));
    p.on('response', async r => { try { bytes += (await r.body()).length; } catch {} });
    await p.goto(URL_PAGINA); await p.waitForLoadState('networkidle'); await p.waitForTimeout(500);
    esperar(bytes < 1.5e6, (bytes / 1e6).toFixed(2) + ' MB');
    resultados.push({ nombre: `   (peso inicial medido: ${(bytes / 1e6).toFixed(2)} MB)`, ok: true });
  });

  await prueba(b, 'F22 el mensaje de WhatsApp de la pantalla final lleva pedido, total y pago', async () => {
    const p = await abrir(b); await agregar(p, 'PE1', 1);
    await llenar(p); await p.click('#btn-enviar'); await p.waitForSelector('.exito');
    const href = decodeURIComponent(await p.getAttribute('.btn-wa', 'href'));
    esperar(/pedido #1 \(ref\. 1\)/.test(href) && /1 x Pargo rosado/.test(href) && /Total: \$7/.test(href) && /Pago Móvil/.test(href), href);
  });

  await b.close();
  const fallos = resultados.filter(r => !r.ok);
  resultados.forEach(r => console.log((r.ok ? '✅' : '❌') + ' ' + r.nombre + (r.ok ? '' : '\n     → ' + r.error)));
  console.log(`\n${resultados.filter(r => r.ok && !r.nombre.startsWith('   ')).length}/${resultados.filter(r => !r.nombre.startsWith('   ')).length} pruebas de la web OK`);
  process.exitCode = fallos.length ? 1 : 0;
})();
