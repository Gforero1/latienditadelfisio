/**
 * @OnlyCurrentDoc
 * La Tiendita del Fisio: catálogo, tasa BCV y pedidos.
 * Hoja de cálculo = panel de la tienda. La web lee de aquí y registra aquí los pedidos.
 *
 * Instalación (una sola vez): ejecutar instalar() y luego Implementar > Nueva implementación > Aplicación web
 * (Ejecutar como: yo. Acceso: cualquier usuario).
 */
var H = { PROD: 'Productos', TASA: 'Tasa', PED: 'Pedidos', AJ: 'Ajustes' };
var ZONA = 'America/Caracas';
var ESTADOS = ['Pendiente de pago', 'Pagado', 'Enviado', 'Entregado', 'Cancelado'];
var COLOR_ESTADO = { 'Pendiente de pago': '#FFE9A8', 'Pagado': '#BFE8C6', 'Enviado': '#BFD9FF', 'Entregado': '#E3D0F5', 'Cancelado': '#E6E6E6' };
var FILA_HIST = 15; // primera fila de datos del historial de tasas

// clave, etiqueta visible, valor inicial
var AJUSTES = [
  ['nombre', 'Nombre de la tienda', 'La Tiendita del Fisio'],
  ['whatsapp', 'WhatsApp de pedidos (con 58, sin + ni espacios)', '584121854095'],
  ['moneda_tasa', 'Tasa para calcular Bs (EUR o USD)', 'EUR'],
  ['mostrar_agotados', 'Mostrar productos agotados (Sí o No)', 'Sí'],
  ['aviso', 'Aviso arriba de la tienda (opcional)', ''],
  ['ventajas', 'Ventajas que se muestran arriba (separadas por coma)', 'Envíos a toda Venezuela, Paga en Bs o en divisas, Atención directa por WhatsApp, Retiro en Lechería'],
  ['instagram', 'Instagram', 'latienditadelfisio.ve'],
  ['ubicacion', 'Ubicación', 'C.E. Nueva Prensa, Av. Intercomunal Jorge Rodríguez, al lado del Hospital Las Garzas. Lechería, Anzoátegui.'],
  ['horario', 'Horario de atención', ''],
  ['paso1', 'Cómo comprar: paso 1', 'Elige tus productos y toca "Agregar".'],
  ['paso2', 'Cómo comprar: paso 2', 'Toca "Ver pedido" y luego "Enviar pedido por WhatsApp".'],
  ['paso3', 'Cómo comprar: paso 3', 'Te confirmamos disponibilidad por el chat, pagas y coordinamos la entrega.'],
  ['pagos', 'Métodos de pago (separados por coma)', 'Pago Móvil (Venezuela, Mercantil, Banesco), Transferencia, Zelle, Binance, Efectivo en divisas'],
  ['entrega', 'Opciones de entrega (separadas por coma)', 'Retiro en tienda, Delivery en Lechería, Barcelona y Puerto La Cruz (costo adicional), Envíos nacionales por Zoom, MRW y Tealca (cobro a destino)'],
  ['nota_precios', 'Nota sobre los precios', 'El monto en bolívares se calcula con la tasa oficial del BCV del día del pago.'],
  ['msg_Pendiente de pago', 'Mensaje al cliente: Pendiente de pago', 'Hola {nombre}, recibimos tu pedido {pedido} por {total}. Te paso los datos para el pago. Cuando lo hagas, envíanos el comprobante por aquí.'],
  ['msg_Pagado', 'Mensaje al cliente: Pagado', 'Hola {nombre}, confirmamos el pago de tu pedido {pedido}. ¡Gracias por tu compra! Ya lo estamos preparando.'],
  ['msg_Enviado', 'Mensaje al cliente: Enviado', 'Hola {nombre}, tu pedido {pedido} ya va en camino. Cualquier duda, escríbenos por aquí.'],
  ['msg_Entregado', 'Mensaje al cliente: Entregado', 'Hola {nombre}, tu pedido {pedido} fue entregado. ¡Gracias por confiar en La Tiendita del Fisio!'],
  ['msg_Cancelado', 'Mensaje al cliente: Cancelado', 'Hola {nombre}, tu pedido {pedido} fue cancelado. Si quieres retomarlo, escríbenos por aquí.']
];

/* ============================ INSTALACIÓN ============================ */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🛒 Tienda')
    .addItem('Actualizar tasa BCV ahora', 'actualizarTasa')
    .addItem('Refrescar la web ahora', 'refrescarWeb')
    .addItem('Crear o cambiar la clave del panel', 'cambiarClavePanel')
    .addSeparator()
    .addItem('Instalar o reparar', 'instalar')
    .addToUi();
}

function instalar() {
  var ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetTimeZone(ZONA);
  if (ss.getName().indexOf('Hoja de cálculo sin título') === 0 || ss.getName().indexOf('Untitled') === 0) ss.rename('La Tiendita del Fisio: tienda');
  crearProductos_(ss);
  crearTasa_(ss);
  crearPedidos_(ss);
  crearAjustes_(ss);
  var sobra = ss.getSheetByName('Hoja 1') || ss.getSheetByName('Sheet1');
  if (sobra && sobra.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(sobra);
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'actualizarTasa') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('actualizarTasa').timeBased().everyHours(1).create();
  actualizarTasa();
  ss.setActiveSheet(ss.getSheetByName(H.PROD));
}

function hoja_(ss, nombre, idx) {
  return ss.getSheetByName(nombre) || ss.insertSheet(nombre, idx);
}

function crearProductos_(ss) {
  var sh = hoja_(ss, H.PROD, 0);
  var cab = ['Nombre', 'Precio $', 'Categoría', 'Disponible', 'Foto', 'Descripción corta', 'Opciones (color, talla...)', 'Precio antes $ (oferta)', 'Destacado'];
  sh.getRange(1, 1, 1, cab.length).setValues([cab]).setFontWeight('bold').setBackground('#3B0A57').setFontColor('#FFFFFF');
  sh.setFrozenRows(1); sh.setFrozenColumns(1);
  if (sh.getLastRow() < 2) {
    var filas = SEMILLA.map(function (p) { return [p[0], p[1] === null ? '' : p[1], p[2], true, p[3], p[4], p[5]]; });
    sh.getRange(2, 1, filas.length, cab.length).setValues(filas);
  }
  var n = Math.max(sh.getMaxRows() - 1, 1);
  sh.getRange(2, 4, n, 1).insertCheckboxes();
  sh.getRange(2, 9, n, 1).insertCheckboxes();
  sh.getRange(2, 8, n, 1).setNumberFormat('0.##');
  sh.getRange(1, 8).setNote('Opcional. Si el producto está en oferta, escribe aquí el precio de antes. En la web sale tachado junto al precio actual y el producto aparece en la pestaña Ofertas.');
  sh.getRange(1, 9).setNote('Marca los productos que más se venden. Salen primero y con la etiqueta Popular.');
  var cats = {};
  SEMILLA.forEach(function (p) { cats[p[2]] = 1; });
  sh.getRange(2, 3, n, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(Object.keys(cats), true).setAllowInvalid(true).build());
  sh.getRange(2, 2, n, 1).setNumberFormat('0.##');
  [260, 80, 170, 90, 200, 360, 240, 120, 90].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  sh.getRange(1, 2).setNote('Solo el número, en dólares. Si lo dejas vacío, la web muestra "Precio a consultar".');
  sh.getRange(1, 4).setNote('Desmarca la casilla cuando se agote. Vuelve a marcarla cuando llegue.');
  sh.getRange(1, 5).setNote('Foto del producto. Para una foto nueva: súbela a Google Drive, compártela como "Cualquier persona con el enlace", copia el enlace y pégalo aquí. Si la dejas vacía sale el logo.');
  sh.getRange(1, 7).setNote('Opcional. Separadas por coma, por ejemplo: Azul, Negro, Rosado. El cliente elige una al agregar.');
}

function crearTasa_(ss) {
  var sh = hoja_(ss, H.TASA, 1);
  if (sh.getRange('A1').getValue() !== 'TASA BCV') {
    sh.getRange('A1').setValue('TASA BCV').setFontSize(14).setFontWeight('bold');
    sh.getRange('A3:A8').setValues([['Dólar (USD)'], ['Euro (EUR)'], ['Fecha de la tasa'], ['Fuente'], ['Última revisión'], ['Resultado']]);
    sh.getRange('A10').setValue('TASA MANUAL: si escribes un número aquí, manda sobre la automática. Bórralo para volver a la automática.').setFontWeight('bold').setWrap(true);
    sh.getRange('A11:A12').setValues([['Dólar manual'], ['Euro manual']]);
    sh.getRange('B11:B12').setBackground('#FFF4C2');
    sh.getRange(FILA_HIST - 1, 1, 1, 5).setValues([['Fecha valor', 'USD', 'EUR', 'Fuente', 'Guardado']]).setFontWeight('bold').setBackground('#3B0A57').setFontColor('#FFFFFF');
    sh.getRange('A3:A8').setFontWeight('bold');
    sh.setColumnWidth(1, 330); sh.setColumnWidth(2, 170); sh.setColumnWidth(3, 130); sh.setColumnWidth(4, 150); sh.setColumnWidth(5, 150);
    sh.getRange('A' + FILA_HIST + ':A').setNumberFormat('@');
    sh.getRange('B5').setNumberFormat('@');
  }
}

function crearPedidos_(ss) {
  var sh = hoja_(ss, H.PED, 2);
  var cab = ['Fecha', 'N° pedido', 'Cliente', 'Productos', 'Total $', 'Total Bs', 'Tasa', 'Estado', 'Teléfono del cliente', 'Avisar al cliente', 'Notas'];
  sh.getRange(1, 1, 1, cab.length).setValues([cab]).setFontWeight('bold').setBackground('#3B0A57').setFontColor('#FFFFFF');
  sh.setFrozenRows(1);
  var n = sh.getMaxRows() - 1;
  var rEstado = sh.getRange(2, 8, n, 1);
  rEstado.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(ESTADOS, true).setAllowInvalid(false).build());
  var reglas = ESTADOS.map(function (e) {
    return SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(e).setBackground(COLOR_ESTADO[e]).setRanges([rEstado]).build();
  });
  sh.setConditionalFormatRules(reglas);
  sh.getRange(2, 1, n, 1).setNumberFormat('dd/MM HH:mm');
  sh.getRange(2, 5, n, 1).setNumberFormat('"$"#,##0.00');
  sh.getRange(2, 6, n, 1).setNumberFormat('"Bs "#,##0.00');
  sh.getRange(2, 9, n, 1).setNumberFormat('@');
  sh.getRange(2, 4, n, 1).setWrap(true);
  [95, 100, 140, 360, 85, 120, 80, 150, 150, 150, 200].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  sh.getRange(1, 9).setNote('Pega aquí el número desde el que te escribió el cliente (ej. 04121234567). Así el enlace "Avisar" abre su chat directo.');
  sh.getRange(1, 10).setNote('Toca el enlace para abrir WhatsApp con el mensaje ya escrito según el estado del pedido.');
}

function crearAjustes_(ss) {
  var sh = hoja_(ss, H.AJ, 3);
  var actuales = {};
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function (f) { actuales[f[0]] = f[1]; });
  sh.getRange(1, 1, 1, 2).setValues([['Ajuste', 'Valor (edita solo esta columna)']]).setFontWeight('bold').setBackground('#3B0A57').setFontColor('#FFFFFF');
  var filas = AJUSTES.map(function (a) { return [a[1], actuales.hasOwnProperty(a[1]) ? actuales[a[1]] : a[2]]; });
  sh.getRange(2, 1, filas.length, 2).setValues(filas);
  sh.getRange(2, 2, filas.length, 1).setNumberFormat('@').setWrap(true).setBackground('#FFF9E0');
  sh.getRange(2, 1, filas.length, 1).setFontWeight('bold').setWrap(true);
  sh.setColumnWidth(1, 330); sh.setColumnWidth(2, 520); sh.setFrozenRows(1);
}

/* ============================ TASA BCV ============================ */

function hoyCaracas_() { return Utilities.formatDate(new Date(), ZONA, 'yyyy-MM-dd'); }

function numVe_(txt) { // "876,79760000" o "1.234,56" -> número
  var n = parseFloat(String(txt).replace(/\s/g, '').replace(/\./g, '').replace(',', '.'));
  return isFinite(n) ? n : null;
}

function tasaValida_(usd, eur) {
  return usd && eur && usd > 1 && usd < 1e7 && eur > usd * 0.7 && eur < usd * 1.6;
}

function leerBCV_() {
  var r = UrlFetchApp.fetch('https://www.bcv.org.ve/', { muteHttpExceptions: true, validateHttpsCertificates: false, followRedirects: true, headers: { 'User-Agent': 'Mozilla/5.0', 'Accept-Language': 'es' } });
  if (r.getResponseCode() !== 200) throw new Error('BCV respondió ' + r.getResponseCode());
  var html = r.getContentText();
  var mu = html.match(/id="dolar"[\s\S]{0,700}?<strong[^>]*>\s*([\d.,]+)\s*<\/strong>/);
  var me = html.match(/id="euro"[\s\S]{0,700}?<strong[^>]*>\s*([\d.,]+)\s*<\/strong>/);
  var mf = html.match(/Fecha Valor:[\s\S]{0,300}?content="(\d{4}-\d{2}-\d{2})/) || html.match(/date-display-single[^>]*content="(\d{4}-\d{2}-\d{2})/);
  if (!mu || !me || !mf) throw new Error('No encontré la tasa en la página del BCV (cambió el formato)');
  var usd = numVe_(mu[1]), eur = numVe_(me[1]);
  if (!tasaValida_(usd, eur)) throw new Error('Tasa del BCV fuera de rango: ' + mu[1] + ' / ' + me[1]);
  return { fecha: mf[1], usd: usd, eur: eur, fuente: 'BCV' };
}

function leerRespaldo_() {
  function una(url) {
    var r = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (r.getResponseCode() !== 200) throw new Error('Respaldo respondió ' + r.getResponseCode());
    return JSON.parse(r.getContentText());
  }
  var d = una('https://ve.dolarapi.com/v1/dolares/oficial'), e = una('https://ve.dolarapi.com/v1/euros/oficial');
  var usd = Number(d.promedio), eur = Number(e.promedio);
  if (!tasaValida_(usd, eur)) throw new Error('Tasa de respaldo fuera de rango');
  return { fecha: String(d.fechaActualizacion).slice(0, 10), usd: usd, eur: eur, fuente: 'Respaldo (dolarapi)' };
}

function historial_(sh) {
  var ult = sh.getLastRow();
  if (ult < FILA_HIST) return [];
  return sh.getRange(FILA_HIST, 1, ult - FILA_HIST + 1, 4).getDisplayValues()
    .map(function (f, i) { return { fila: FILA_HIST + i, fecha: f[0], usd: numCelda_(f[1]), eur: numCelda_(f[2]), fuente: f[3] }; })
    .filter(function (h) { return /^\d{4}-\d{2}-\d{2}$/.test(h.fecha) && h.usd && h.eur; });
}

function numCelda_(txt) { // valor mostrado según la configuración regional de la hoja
  var s = String(txt).trim();
  if (!s) return null;
  if (s.indexOf(',') > -1) s = s.replace(/\./g, '').replace(',', '.');
  var n = parseFloat(s);
  return isFinite(n) ? n : null;
}

function guardarTasa_(sh, t) {
  var hist = historial_(sh), igual = null;
  hist.forEach(function (h) { if (h.fecha === t.fecha) igual = h; });
  var ahora = Utilities.formatDate(new Date(), ZONA, 'dd/MM/yyyy HH:mm');
  if (igual) {
    if (igual.fuente === 'BCV' && t.fuente !== 'BCV') return; // el respaldo nunca pisa al BCV
    sh.getRange(igual.fila, 1, 1, 5).setValues([[t.fecha, t.usd, t.eur, t.fuente, ahora]]);
  } else {
    sh.insertRowBefore(FILA_HIST);
    sh.getRange(FILA_HIST, 1, 1, 5).setValues([[t.fecha, t.usd, t.eur, t.fuente, ahora]]).setFontWeight('normal').setBackground(null).setFontColor('#000000');
    sh.getRange(FILA_HIST, 1).setNumberFormat('@');
    sh.getRange(FILA_HIST, 2, 1, 2).setNumberFormat('0.0000');
    if (sh.getLastRow() > FILA_HIST + 400) sh.deleteRows(FILA_HIST + 400, sh.getLastRow() - FILA_HIST - 399);
  }
}

/** Tasa vigente = la última del historial cuya fecha valor ya llegó (hora de Caracas). */
function vigente_(sh) {
  var hoy = hoyCaracas_(), hist = historial_(sh).sort(function (a, b) { return a.fecha < b.fecha ? 1 : -1; });
  var v = null;
  hist.forEach(function (h) { if (!v && h.fecha <= hoy) v = h; });
  if (!v && hist.length) v = hist[hist.length - 1];
  return v;
}

/** Se ejecuta sola cada hora. BCV primero, respaldo si falla, y si todo falla se queda la anterior. */
function actualizarTasa() {
  var sh = SpreadsheetApp.getActive().getSheetByName(H.TASA), notas = [];
  try { guardarTasa_(sh, leerBCV_()); notas.push('BCV: ok'); }
  catch (e) { notas.push('BCV: falló (' + e.message + ')'); }
  var v = vigente_(sh);
  if (!v || notas[0] !== 'BCV: ok' || v.fecha > hoyCaracas_()) {
    try { guardarTasa_(sh, leerRespaldo_()); notas.push('Respaldo: ok'); }
    catch (e2) { notas.push('Respaldo: falló (' + e2.message + ')'); }
    v = vigente_(sh);
  }
  if (v) {
    sh.getRange('B3:B6').setValues([[v.usd], [v.eur], [v.fecha.split('-').reverse().join('/')], [v.fuente]]);
    sh.getRange('B3:B4').setNumberFormat('0.0000');
  }
  if (notas.join(' ').indexOf('ok') === -1) notas.push('Se mantiene la tasa anterior.');
  sh.getRange('B7:B8').setValues([[Utilities.formatDate(new Date(), ZONA, 'dd/MM/yyyy HH:mm')], [notas.join(' | ')]]);
  refrescarWeb();
}

function tasaParaWeb_(ss) {
  var sh = ss.getSheetByName(H.TASA), v = vigente_(sh) || { usd: null, eur: null, fecha: '', fuente: '' };
  var man = sh.getRange('B11:B12').getValues(), mu = Number(man[0][0]) || null, me = Number(man[1][0]) || null;
  return { usd: mu || v.usd, eur: me || v.eur, fecha: v.fecha, fuente: v.fuente, manual_usd: !!mu, manual_eur: !!me, hoy: hoyCaracas_() };
}

/* ============================ WEB: LEER CATÁLOGO ============================ */

function ajustes_(ss) {
  var sh = ss.getSheetByName(H.AJ), porEtiqueta = {}, out = {};
  sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 2).getDisplayValues().forEach(function (f) { porEtiqueta[f[0]] = f[1]; });
  AJUSTES.forEach(function (a) { out[a[0]] = porEtiqueta.hasOwnProperty(a[1]) ? String(porEtiqueta[a[1]]).trim() : a[2]; });
  return out;
}

function fotoUrl_(x) { // enlace de Drive compartido -> imagen directa
  var m = x.match(/drive\.google\.com\/(?:file\/d\/|open\?id=|uc\?[^#]*id=)([\w-]{20,})/);
  return m ? 'https://drive.google.com/thumbnail?id=' + m[1] + '&sz=w640' : x;
}

function datos_() {
  var ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName(H.PROD);
  var filas = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 9).getValues() : [];
  var productos = [];
  filas.forEach(function (f) {
    var nombre = String(f[0]).trim();
    if (!nombre) return;
    var precio = f[1] === '' || f[1] === null ? null : Number(String(f[1]).replace(',', '.'));
    var fotos = String(f[4]).split(/[\n;]+|,\s+/).map(function (x) { return x.trim(); }).filter(String).map(fotoUrl_);
    productos.push({
      n: nombre, p: isFinite(precio) && precio > 0 ? precio : null, c: String(f[2]).trim() || 'Otros',
      d: f[3] !== false && String(f[3]).toLowerCase() !== 'no', f: fotos, t: String(f[5]).trim(),
      o: String(f[6]).split(',').map(function (x) { return x.trim(); }).filter(String),
      pa: Number(f[7]) > 0 ? Number(f[7]) : null, x: f[8] === true
    });
  });
  var aj = ajustes_(ss), pub = {};
  Object.keys(aj).forEach(function (k) { if (k.indexOf('msg_') !== 0) pub[k] = aj[k]; });
  return { ok: true, productos: productos, tasa: tasaParaWeb_(ss), ajustes: pub, generado: new Date().toISOString() };
}

function refrescarWeb() { CacheService.getScriptCache().remove('datos'); }

function doGet() {
  var cache = CacheService.getScriptCache(), txt = cache.get('datos');
  if (!txt) {
    try { txt = JSON.stringify(datos_()); if (txt.length < 95000) cache.put('datos', txt, 120); }
    catch (e) { txt = JSON.stringify({ ok: false, error: String(e) }); }
  }
  return ContentService.createTextOutput(txt).setMimeType(ContentService.MimeType.JSON);
}

/* ============================ WEB: REGISTRAR PEDIDO ============================ */

function limpio_(v, max) {
  var s = String(v === undefined || v === null ? '' : v).replace(/[\u0000-\u0008\u000B-\u001F]/g, '').slice(0, max);
  return /^[=+\-@]/.test(s) ? "'" + s : s; // nunca fórmulas escritas desde fuera
}

function doPost(e) {
  var res = { ok: false };
  try {
    var d = JSON.parse(e.postData.contents);
    if (d && typeof d.accion === 'string' && d.accion.indexOf('admin_') === 0) return salida_(admin_(d));
    if (!d || !/^TDF-[A-Z0-9]{4,8}$/.test(d.pedido) || !Array.isArray(d.items) || !d.items.length || d.items.length > 80) throw new Error('Pedido inválido');
    var lineas = d.items.map(function (i) {
      var q = Math.max(1, Math.min(999, parseInt(i.q, 10) || 1));
      return q + ' x ' + limpio_(i.n, 120).replace(/^'/, '') + (i.o ? ' (' + limpio_(i.o, 60).replace(/^'/, '') + ')' : '') + (i.p ? ' = $' + (Number(i.p) * q).toFixed(2) : ' (precio a consultar)');
    });
    var lock = LockService.getScriptLock(); lock.waitLock(15000);
    try {
      var ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName(H.PED);
      var ids = sh.getLastRow() > 1 ? sh.getRange(2, 2, sh.getLastRow() - 1, 1).getValues() : [];
      for (var k = 0; k < ids.length; k++) if (ids[k][0] === d.pedido) { res.ok = true; res.repetido = true; break; }
      if (!res.repetido) {
        sh.insertRowBefore(2); // el pedido más nuevo siempre arriba
        sh.getRange(2, 1, 1, 9).setValues([[new Date(), d.pedido, limpio_(d.nombre, 60), limpio_(lineas.join('\n'), 4000),
          Number(d.total_usd) || 0, Number(d.total_bs) || 0, Number(d.tasa) || '', ESTADOS[0], '']]);
        enlaceAviso_(ss, sh, 2);
        res.ok = true;
      }
    } finally { lock.releaseLock(); }
  } catch (err) { res.error = String(err); }
  return ContentService.createTextOutput(JSON.stringify(res)).setMimeType(ContentService.MimeType.JSON);
}

/* ============================ PEDIDOS: AVISAR AL CLIENTE ============================ */

function telefono_(t) {
  var d = String(t).replace(/\D/g, '');
  if (!d) return '';
  if (d.length === 11 && d.charAt(0) === '0') return '58' + d.slice(1);
  if (d.length === 10 && d.charAt(0) === '4') return '58' + d;
  return d;
}

function avisoUrl_(aj, f) { // f = fila de Pedidos (9 columnas)
  var pedido = f[1], estado = f[7];
  if (!pedido || !estado) return '';
  var plantilla = aj['msg_' + estado] || '';
  var total = '$' + (Number(f[4]) || 0).toFixed(2) + (Number(f[5]) ? ' (Bs ' + Number(f[5]).toLocaleString('es-VE', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ')' : '');
  var nombre = String(f[2]).replace(/^'/, '').trim();
  var msg = plantilla.replace(/\{nombre\}/g, nombre).replace(/\{pedido\}/g, pedido).replace(/\{total\}/g, total).replace(/Hola\s+,/, 'Hola,');
  return 'https://wa.me/' + telefono_(f[8]) + '?text=' + encodeURIComponent(msg);
}

function enlaceAviso_(ss, sh, fila) {
  var f = sh.getRange(fila, 1, 1, 9).getValues()[0], celda = sh.getRange(fila, 10), url = avisoUrl_(ajustes_(ss), f);
  if (!url) { celda.clearContent(); return; }
  celda.setRichTextValue(SpreadsheetApp.newRichTextValue().setText('Avisar: ' + f[7]).setLinkUrl(url).build());
}

/** Al editar a mano: rehace el enlace "Avisar" del pedido y refresca la web si cambió el catálogo. */
function onEdit(e) {
  try {
    var sh = e.range.getSheet(), nombre = sh.getName();
    if (nombre === H.PED) {
      var c1 = e.range.getColumn(), c2 = e.range.getLastColumn();
      if (c2 < 2 || c1 > 9) return;
      var ss = e.source;
      for (var r = Math.max(2, e.range.getRow()); r <= e.range.getLastRow(); r++) enlaceAviso_(ss, sh, r);
    } else if (nombre === H.PROD || nombre === H.AJ || nombre === H.TASA) {
      refrescarWeb();
    }
  } catch (err) { }
}

/* ============================ PANEL PRIVADO (admin.html) ============================ */

function salida_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function hash_(txt, sal) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, sal + '|' + txt, Utilities.Charset.UTF_8)
    .map(function (b) { return ('0' + (b & 255).toString(16)).slice(-2); }).join('');
}

/** Menú Tienda > Crear o cambiar la clave del panel. La clave no se guarda en claro. */
function cambiarClavePanel() {
  var ui = SpreadsheetApp.getUi();
  var r = ui.prompt('Clave del panel', 'Escribe la clave nueva (mínimo 8 caracteres). Es la que se pide al entrar al panel de productos.', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  var clave = r.getResponseText().trim();
  if (clave.length < 8) { ui.alert('La clave debe tener al menos 8 caracteres. No se cambió nada.'); return; }
  var sal = Utilities.getUuid();
  PropertiesService.getScriptProperties().setProperties({ CLAVE_SAL: sal, CLAVE_HASH: hash_(clave, sal) });
  ui.alert('Listo. La clave del panel quedó guardada.');
}

function claveOk_(clave) {
  var cache = CacheService.getScriptCache(), fallos = Number(cache.get('fallos_clave')) || 0;
  if (fallos >= 8) throw new Error('Demasiados intentos. Espera 10 minutos.');
  var p = PropertiesService.getScriptProperties(), sal = p.getProperty('CLAVE_SAL'), h = p.getProperty('CLAVE_HASH');
  if (!sal || !h) throw new Error('El panel aún no tiene clave. Créala en la hoja: menú Tienda > Crear o cambiar la clave del panel.');
  if (typeof clave !== 'string' || hash_(clave.trim(), sal) !== h) {
    cache.put('fallos_clave', String(fallos + 1), 600);
    throw new Error('Clave incorrecta.');
  }
}

function carpetaFotos_() {
  var p = PropertiesService.getScriptProperties(), id = p.getProperty('CARPETA_FOTOS');
  if (id) return id;
  var c = Drive.Files.create({ name: 'Fotos La Tiendita del Fisio', mimeType: 'application/vnd.google-apps.folder' }, null, { fields: 'id' });
  Drive.Permissions.create({ role: 'reader', type: 'anyone' }, c.id);
  p.setProperty('CARPETA_FOTOS', c.id);
  return c.id;
}

/** Crea la carpeta de fotos en Drive (y pide el permiso si falta). Basta ejecutarla una vez. */
function prepararFotos() { Logger.log('Carpeta de fotos: ' + carpetaFotos_()); }

function subirFoto_(foto, nombre) {
  if (!foto || !/^image\/(webp|jpeg|png)$/.test(foto.mime) || typeof foto.b64 !== 'string' || foto.b64.length > 900000) throw new Error('Foto no válida o demasiado grande.');
  var ext = foto.mime.split('/')[1].replace('jpeg', 'jpg');
  var base = String(nombre).toLowerCase().replace(/[^a-z0-9]+/g, '-').slice(0, 40) || 'producto';
  var blob = Utilities.newBlob(Utilities.base64Decode(foto.b64), foto.mime, base + '-' + Date.now() + '.' + ext);
  var f = Drive.Files.create({ name: blob.getName(), parents: [carpetaFotos_()] }, blob, { fields: 'id' });
  return 'https://drive.google.com/thumbnail?id=' + f.id + '&sz=w640';
}

function filasAdmin_(sh) {
  var n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, 9).getValues().map(function (f, i) {
    var foto = String(f[4]).split(/[\n;]+|,\s+/)[0].trim();
    return { fila: i + 2, n: String(f[0]).trim(), p: f[1] === '' ? null : Number(f[1]), c: String(f[2]).trim(), d: f[3] !== false, f: foto ? fotoUrl_(foto) : '', t: String(f[5]), o: String(f[6]), pa: Number(f[7]) > 0 ? Number(f[7]) : null, x: f[8] === true };
  }).filter(function (x) { return x.n; });
}

function pedidosAdmin_(ss, pe) {
  var n = Math.min(pe.getLastRow() - 1, 200);
  if (n < 1) return [];
  var aj = ajustes_(ss);
  return pe.getRange(2, 1, n, 11).getValues().filter(function (f) { return f[1]; }).map(function (f) {
    return { id: String(f[1]), fecha: f[0] instanceof Date ? f[0].toISOString() : String(f[0]), cliente: String(f[2]), productos: String(f[3]),
      usd: Number(f[4]) || 0, bs: Number(f[5]) || 0, estado: String(f[7]), tel: String(f[8]), notas: String(f[10]), aviso: avisoUrl_(aj, f) };
  });
}

function admin_(d) {
  var res = { ok: false };
  try {
    claveOk_(d.clave);
    var lock = LockService.getScriptLock(); lock.waitLock(20000);
    try {
      var sh = SpreadsheetApp.getActive().getSheetByName(H.PROD);
      if (d.accion === 'admin_guardar' || d.accion === 'admin_eliminar') {
        var fila = d.fila ? parseInt(d.fila, 10) : 0;
        if (fila) { // la fila debe seguir siendo el mismo producto
          if (fila < 2 || fila > sh.getLastRow() || String(sh.getRange(fila, 1).getValue()).trim() !== String(d.original || '').trim())
            throw new Error('La lista cambió mientras editabas. Vuelve a abrir el producto.');
        }
        if (d.accion === 'admin_eliminar') {
          if (!fila) throw new Error('Falta el producto a eliminar.');
          sh.deleteRow(fila);
        } else {
          var x = d.producto || {}, nombre = limpio_(x.n, 120).trim();
          if (!nombre) throw new Error('El producto necesita un nombre.');
          var precio = x.p === '' || x.p === null || x.p === undefined ? '' : Number(String(x.p).replace(',', '.'));
          if (precio !== '' && (!isFinite(precio) || precio < 0 || precio > 100000)) throw new Error('El precio no es válido.');
          var antes = x.pa === '' || x.pa === null || x.pa === undefined ? '' : Number(String(x.pa).replace(',', '.'));
          if (antes !== '' && (!isFinite(antes) || antes < 0 || antes > 100000)) throw new Error('El precio de antes no es válido.');
          var foto = d.foto ? subirFoto_(d.foto, nombre) : null;
          if (!fila) {
            fila = Math.max(sh.getLastRow(), 1) + 1;
            if (fila > sh.getMaxRows()) sh.insertRowAfter(sh.getMaxRows());
            sh.getRange(fila, 4).insertCheckboxes();
            sh.getRange(fila, 9).insertCheckboxes();
            sh.getRange(fila, 5).setValue(foto || '');
          } else if (foto) sh.getRange(fila, 5).setValue(foto);
          sh.getRange(fila, 1, 1, 4).setValues([[nombre, precio, limpio_(x.c, 60).trim() || 'Otros', x.d !== false]]);
          sh.getRange(fila, 6, 1, 4).setValues([[limpio_(x.t, 400), limpio_(x.o, 300), antes, x.x === true]]);
        }
        SpreadsheetApp.flush();
        refrescarWeb();
      } else if (d.accion === 'admin_pedidos' || d.accion === 'admin_pedido') {
        var ss = SpreadsheetApp.getActive(), pe = ss.getSheetByName(H.PED);
        if (d.accion === 'admin_pedido') {
          if (ESTADOS.indexOf(d.estado) < 0) throw new Error('Estado no válido.');
          var ids = pe.getLastRow() > 1 ? pe.getRange(2, 2, pe.getLastRow() - 1, 1).getValues() : [], fp = 0;
          for (var k = 0; k < ids.length; k++) if (ids[k][0] === d.pedido) { fp = k + 2; break; }
          if (!fp) throw new Error('No encontré ese pedido. Actualiza la lista.');
          pe.getRange(fp, 8, 1, 2).setValues([[d.estado, limpio_(String(d.tel || '').replace(/[^\d+ ()-]/g, ''), 25)]]);
          pe.getRange(fp, 11).setValue(limpio_(d.notas, 300));
          enlaceAviso_(ss, pe, fp);
          SpreadsheetApp.flush();
        }
        res.pedidos = pedidosAdmin_(ss, pe);
        res.estados = ESTADOS;
        res.ok = true;
        return res;
      } else if (d.accion !== 'admin_listar') throw new Error('Acción desconocida.');
      res.productos = filasAdmin_(sh);
      res.ok = true;
    } finally { lock.releaseLock(); }
  } catch (err) { res.error = String(err && err.message || err); }
  return res;
}

/* ============================ CATÁLOGO INICIAL ============================ */
// nombre, precio $, categoría, foto, descripción, opciones
var SEMILLA = [
  ["TENS 7000", 110, "Electroterapia", "tens-7000.webp", "5 modos. Intensidad 0-100 mA. 2 canales. Batería 9V incluida. 4 electrodos 5x5cm, 2 cables, manual y maletín.", ""],
  ["TENS-EMS Kingpavonini", 80, "Electroterapia", "tens-ems-kingpavonini.webp", "24 modos y programa personalizado. 40 niveles de intensidad. 2 canales. Recargable. 12 electrodos, 2 cables y estuche.", ""],
  ["TENS-EMS 7000", 150, "Electroterapia", "tens-ems-7000.webp", "5 modos TENS y 3 EMS. Intensidad 0-100 mA. 2 canales. Recargable. 4 electrodos 5x5cm, 2 cables, manual y maletín.", ""],
  ["TENS-EMS 7000 (4 canales)", 180, "Electroterapia", "tens-ems-7000-4-canales.webp", "5 modos TENS y 3 EMS. 4 canales. Recargable. 8 electrodos 5x5cm, 4 cables, manual y maletín.", ""],
  ["Compex Sport Elite 3.0", 665, "Electroterapia", "compex-sport-elite-3-0.webp", "10 programas. Intensidad 0-120 mA. 4 canales. Recargable. 4 electrodos 5x5cm y 2 de 5x10cm, 4 cables, manual y maletín.", ""],
  ["Compex Performance 3.0", 510, "Electroterapia", "compex-performance-3-0.webp", "6 programas. Intensidad 0-120 mA. 4 canales. Recargable, hasta 8 horas. 4 electrodos 5x5cm y 2 de 5x10cm, 4 cables y maletín.", ""],
  ["TENS-EMS portátil", 30, "Electroterapia", "tens-ems-portatil.webp", "3 modos TENS y 3 EMS. 2 canales. Recargable. 19 niveles de intensidad. 8 electrodos y 2 cables.", ""],
  ["Electrodos 5x10cm conector de broche", 7, "Electroterapia", "electrodos-5x10cm-conector-de-broche.webp", "Paquete de 4 electrodos.", ""],
  ["Electrodos 5x5cm conector de broche", 7, "Electroterapia", "electrodos-5x5cm-conector-de-broche.webp", "Paquete de 4 electrodos.", ""],
  ["Electrodos 5x10cm conector tipo pin", 7, "Electroterapia", "electrodos-5x10cm-conector-tipo-pin.webp", "Paquete de 4 electrodos.", ""],
  ["Electrodos TDF conector tipo pin", 6, "Electroterapia", "electrodos-tdf-conector-tipo-pin.webp", "Paquete de 4 electrodos.", "5x5cm, 4x4cm"],
  ["Electrodos faciales", 6, "Electroterapia", "electrodos-faciales.webp", "Paquete de 4 electrodos. Conector de pin.", ""],
  ["Cables de TENS (4 puntas)", 8, "Electroterapia", "cables-de-tens-4-puntas.webp", "Se venden por par. 1.2 metros. 4 puntas cada cable.", ""],
  ["Cables de TENS 7000", 7, "Electroterapia", "cables-de-tens-7000.webp", "Se venden por par. 1.2 metros.", ""],
  ["Cables de TENS", 7, "Electroterapia", "cables-de-tens.webp", "Se venden por par. 1.2 metros.", "Punta gruesa, Punta fina"],
  ["Ultrasonido portátil", 185, "Electroterapia", "ultrasonido-portatil.webp", "3 niveles de intensidad. Tiempo de 5, 10 y 20 min. Modalidad continua 1MHz. Corriente de enchufe.", ""],
  ["Lentes de protección para láser", 13, "Electroterapia", "lentes-de-proteccion-para-laser.webp", "Protección ocular al usar equipos que emiten láser.", ""],
  ["Lápiz de electroacupuntura", 80, "Electroterapia", "lapiz-de-electroacupuntura.webp", "2 cabezales. Recargable (cable USB). 20 niveles de intensidad. 3 modos. Barra conductora.", ""],
  ["Ondas de choque V100 portátil", 750, "Electroterapia", "ondas-de-choque-v100-portatil.webp", "Potencia 100W. Energía de 5 a 250mJ. Frecuencia 1 a 16Hz. 14 cabezales. Corriente de enchufe.", ""],
  ["Láser terapéutico 20 diodos", 260, "Electroterapia", "laser-terapeutico-20-diodos.webp", "4 niveles de intensidad. Modo continuo y pulsátil. 16 diodos 650nm y 4 de 808nm. Tiempo de 1 a 60 min. Incluye lentes y cargador.", ""],
  ["Láser terapéutico 17 diodos", 180, "Electroterapia", "laser-terapeutico-17-diodos.webp", "4 niveles de intensidad. 14 diodos 650nm y 3 de 808nm. Tiempo de 5 a 20 min. Batería 2600mAh. Incluye lentes y cargador.", ""],
  ["Batería 9V recargable", 12, "Electroterapia", "bateria-9v-recargable.webp", "Recargable. Se vende por unidad.", ""],
  ["Bombillo de lámpara infrarrojo 100W", 18, "Calor y frío", "bombillo-de-lampara-infrarrojo.webp", "Pequeño, 100W. 110V.", ""],
  ["Bombillo de lámpara infrarrojo 150W", 20, "Calor y frío", "bombillo-de-lampara-infrarrojo.webp", "Mediano, 150W. 110V.", ""],
  ["Lámpara infrarrojo de pedestal", 120, "Calor y frío", "lampara-infrarrojo-de-pedestal.webp", "Ajustable 360 grados. Bombillo de 275W. 110V. Altura ajustable.", ""],
  ["Fomentera eléctrica pequeña 30x60cm", 55, "Calor y frío", "fomentera-electrica.webp", "10 niveles de calor. Apagado automático. 3 tiempos (45, 90 y 120 min).", ""],
  ["Fomentera eléctrica mediana 40x76cm", 65, "Calor y frío", "fomentera-electrica.webp", "10 niveles de calor. Apagado automático. 3 tiempos (45, 90 y 120 min).", ""],
  ["Fomentera eléctrica grande 50x100cm", 75, "Calor y frío", "fomentera-electrica.webp", "10 niveles de calor. Apagado automático. 3 tiempos (45, 90 y 120 min).", ""],
  ["Medias de compresión con calor", 64, "Calor y frío", "medias-de-compresion-con-calor.webp", "Temperatura de 39° a 65°C. 78% algodón, 20% poliéster, 2% spandex. Incluye 2 power bank de 4000mAh y cable USB.", ""],
  ["Manta de luz LED con vibración", 55, "Calor y frío", "manta-de-luz-led-con-vibracion.webp", "Neopreno, 30x28cm. Batería 5000mAh. Luz roja 660nm y 850nm, 4 modos. 3 modos de vibración. Tiempos de 15, 30 y 45 min.", ""],
  ["Compresa frío/calor", 12, "Calor y frío", "compresa-frio-calor.webp", "Calor en microondas, frío en refrigerador. 23x13cm.", ""],
  ["Rodillera térmica", 40, "Calor y frío", "rodillera-termica.webp", "3 niveles de calor. Corriente de enchufe. Cintas ajustables.", ""],
  ["Rodillera multifuncional con calor y masaje", 55, "Calor y frío", "rodillera-multifuncional-con-calor-y-mas.webp", "Para codo, hombro o rodilla. Calor graduable 65/75 grados. Masaje vibratorio de 3 modos. Recargable, batería 5000mAh.", ""],
  ["Hombrera térmica", 40, "Calor y frío", "hombrera-termica.webp", "Calor de 65/75 grados. Corriente de enchufe. 3 niveles con indicador de luz.", ""],
  ["Rodillera multifuncional Konbest frío/calor", 345, "Calor y frío", "rodillera-multifuncional-konbest-frio-ca.webp", "Enfriamiento rápido (16, 12 y 8°C) y calentamiento (35, 38 y 42°C). Luz LED roja. 3 niveles de vibración. Recargable.", ""],
  ["Lámpara de luz LED", 50, "Calor y frío", "lampara-de-luz-led.webp", "Luz roja fría 660nm + 850nm. Pedestal ajustable. Temporizador. Incluye lentes. Corriente de enchufe.", ""],
  ["Cinturón de tracción lumbar", 50, "Masaje y recuperación", "cinturon-de-traccion-lumbar.webp", "Incluye bomba para inflar. Medidas 73x124cm. Descompresión, relajación y alivio del dolor lumbar.", ""],
  ["Cinturón de tracción lumbar con calor", 70, "Masaje y recuperación", "cinturon-de-traccion-lumbar-con-calor.webp", "Incluye bomba para inflar. Recargable, batería 5000mAh. 3 modos de calor y 3 de vibración.", ""],
  ["Mini pistola de masaje", 45, "Masaje y recuperación", "mini-pistola-de-masaje.webp", "Batería 1500mAh. 4 cabezales. 30 niveles de velocidad. Cable USB.", ""],
  ["Pistola de masaje 6 cabezales", 92, "Masaje y recuperación", "pistola-de-masaje-6-cabezales.webp", "Batería 2500mAh, 6 a 8 horas de trabajo. 30 niveles de velocidad. 3600rpm. Incluye cargador y maletín.", ""],
  ["Pistola de masaje 9 cabezales", null, "Masaje y recuperación", "pistola-de-masaje-9-cabezales.webp", "Batería 2500mAh, 6 a 8 horas de trabajo. 30 niveles de velocidad. Cable USB.", ""],
  ["Pistola de masaje AERLANG con calor", 120, "Masaje y recuperación", "pistola-de-masaje-aerlang-con-calor.webp", "6 cabezales y 1 cabezal de calor. Batería de 3 a 6 horas. 20 niveles de velocidad. 3200rpm. Cable USB.", ""],
  ["Presoterapia Konbest 8 cámaras", 720, "Masaje y recuperación", "presoterapia-konbest-8-camaras.webp", "Modelo 1. 80 a 260 mmHg. 6 modos de masaje. Recargable, batería 5200mAh. Tiempo ajustable de 10 a 90 min.", ""],
  ["Presoterapia Konbest 6 cámaras", null, "Masaje y recuperación", "presoterapia-konbest-6-camaras.webp", "Modelo 2. 80 a 260 mmHg. 4 modos de masaje. Recargable, batería 5200mAh. Tiempo ajustable de 10 a 90 min.", ""],
  ["Pistola de masaje Triton Vibes", 50, "Masaje y recuperación", "pistola-de-masaje-triton-vibes.webp", "8 cabezales. 99 niveles de velocidad. 8 horas de trabajo. 3200rpm. Maletín y cable USB.", ""],
  ["Masajeador de cuello", 40, "Masaje y recuperación", "masajeador-de-cuello.webp", "3 niveles de calor (42/47/52°C). 4 modos de masaje de vibración. Batería 550mAh.", ""],
  ["Terapia manual (combo de 2)", null, "Masaje y recuperación", "terapia-manual-combo-de-2.webp", "Herramientas para técnicas miofasciales. Se vende en combo de 2.", ""],
  ["Cupping 24 piezas TDF", 50, "Masaje y recuperación", "cupping-24-piezas-tdf.webp", "Set de 24 copas. Pistola de succión y manguera de extensión.", ""],
  ["Set de cupping 32 piezas", 55, "Masaje y recuperación", "set-de-cupping-32-piezas.webp", "Pistola de succión. Incluye herramienta guasha.", ""],
  ["Pistola de succión para cupping", 10, "Masaje y recuperación", "pistola-de-succion-para-cupping.webp", "Material ABS resistente y de larga duración.", ""],
  ["Cupping eléctrico", 40, "Masaje y recuperación", "cupping-electrico.webp", "12 niveles de calor (38 a 50°C) y 12 de succión. 2 modos de succión. 20 min de trabajo.", ""],
  ["Foam roller", 20, "Masaje y recuperación", "foam-roller.webp", "Texturizado, 33cm de largo. Se vende por unidad.", ""],
  ["Arco de tracción de columna", 15, "Masaje y recuperación", "arco-de-traccion-de-columna.webp", "Arco de fácil ajuste para tracción de columna.", ""],
  ["Kit de manos 8 piezas", 35, "Mano y rehabilitación", "kit-de-manos-8-piezas.webp", "3 pelotas de goma, 2 aros de goma y 3 gomas de extensión de dedos, con resistencias diferentes. Bolsa de traslado.", ""],
  ["Kit de manos 6 piezas", 35, "Mano y rehabilitación", "kit-de-manos-6-piezas.webp", "Prensa de mano graduable, aro texturizado, goma de extensión, pelota, muñequera extensora y prensa de dedos. Bolsa de traslado.", ""],
  ["Kit de resistencia de manos", 25, "Mano y rehabilitación", "kit-de-resistencia-de-manos.webp", "8 piezas: 4 prensas de dedos y 4 gomas de extensión de dedos. Se vende por set o por separado.", ""],
  ["Plastilina terapéutica", 8, "Mano y rehabilitación", "plastilina-terapeutica.webp", "Plastilina de resistencia para rehabilitación de mano.", "Extra suave, Muy suave, Suave, Duro, Muy duro"],
  ["Aro de goma para mano", 3, "Mano y rehabilitación", "aro-de-goma-para-mano.webp", "Se vende por unidad.", "Verde ligero 30 lb, Azul medio 40 lb, Naranja firme 50 lb"],
  ["Pelota de goma para mano", 3, "Mano y rehabilitación", "pelota-de-goma-para-mano.webp", "Se vende por unidad.", "Amarillo extra ligero, Naranja ligero, Rojo medio, Azul firme, Verde extra firme"],
  ["Mancuerna de mano", 13, "Mano y rehabilitación", "mancuerna-de-mano.webp", "Resistencia graduable. Cuenta las repeticiones.", ""],
  ["Guantes eléctricos de rehabilitación", 230, "Mano y rehabilitación", "guantes-electricos-de-rehabilitacion.webp", "Guante robótico y guante simulador de movimiento. Programa velocidad e intensidad de apertura y cierre de la mano. Entrada de 5V.", ""],
  ["Pedalera", 70, "Mano y rehabilitación", "pedalera.webp", "Cuenta revoluciones y calorías. Indica el tiempo.", ""],
  ["Ejercitador de dedos", 10, "Mano y rehabilitación", "ejercitador-de-dedos.webp", "Resistencia ajustable hasta 11 libras.", ""],
  ["Vendaje coban", 3, "Vendajes y soportes", "vendaje-coban.webp", "5cm de ancho y 5m de largo.", ""],
  ["Tijera Lister", 8, "Vendajes y soportes", "tijera-lister.webp", "Tijera para vendajes.", "Azul, Blanco, Negro, Amarillo, Verde, Rojo, Naranja"],
  ["Vendaje neuromuscular Med Pride", 12, "Vendajes y soportes", "vendaje-neuromuscular-med-pride.webp", "Kinesiotape marca Med Pride.", "Piel, Negro, Azul, Rosado"],
  ["Vendaje neuromuscular", 8, "Vendajes y soportes", "vendaje-neuromuscular.webp", "Kinesiotape.", "Piel, Negro, Azul"],
  ["Plantillas de látex", 15, "Vendajes y soportes", "plantillas-de-latex.webp", "Látex, 1.8cm de grosor. Tallas de la 35 a la 46.", "35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46"],
  ["Rodilleras camufladas de neopreno", 25, "Vendajes y soportes", "rodilleras-camufladas-de-neopreno.webp", "Neopreno. Se vende por par.", "XS, S, M, L"],
  ["Medias de compresión", 12, "Vendajes y soportes", "medias-de-compresion.webp", "Compresión 20-30 mmHg. A nivel de la rodilla. Se vende por par.", ""],
  ["Tobilleras", 14, "Vendajes y soportes", "tobilleras.webp", "68% nylon, 32% spandex. Se vende por par. Talla XL (23-25cm).", ""],
  ["Rodilleras ajustables talla única", 15, "Vendajes y soportes", "rodilleras-ajustables-talla-unica.webp", "Tela suave. Talla única. Se vende por par.", ""],
  ["Rodilleras Vilico negras", 17, "Vendajes y soportes", "rodilleras-vilico-negras.webp", "Tela elástica 3D.", "M, L, XL"],
  ["Rodilleras Vilico negro y rojo", 17, "Vendajes y soportes", "rodilleras-vilico-negro-y-rojo.webp", "Tela elástica 3D. Talla M.", ""],
  ["Rodilleras elásticas", 15, "Vendajes y soportes", "rodilleras-elasticas.webp", "Rodilleras de compresión.", "S, M, L"],
  ["Rotulera", 13, "Vendajes y soportes", "rotulera.webp", "Tela suave acolchada. Talla única. Se vende por par.", ""],
  ["Ortesis de juanete", 10, "Vendajes y soportes", "ortesis-de-juanete.webp", "Se venden por par.", "Tela, Silicón"],
  ["Férula para pulgar", 11, "Vendajes y soportes", "ferula-para-pulgar.webp", "Tela suave. Cinchas para ajustar la presión. Se vende por unidad.", ""],
  ["Férula para dedos", 8, "Vendajes y soportes", "ferula-para-dedos.webp", "Plástico. Cinchas para ajustar la presión. Disponible en 5 tamaños. Se vende por unidad.", ""],
  ["Conos grandes (set de 6)", null, "Ejercicio y deporte", "conos.webp", "Set de 6 conos de 18cm.", ""],
  ["Set de minas (10 unidades)", null, "Ejercicio y deporte", "set-de-minas-10-unidades.webp", "Set de 10 minas, 19cm de circunferencia.", ""],
  ["Straps SKDK", 15, "Ejercicio y deporte", "straps-skdk.webp", "Marca SKDK.", "Negro, Gris"],
  ["Banda de resistencia con tobilleras", 25, "Ejercicio y deporte", "banda-de-resistencia-con-tobilleras.webp", "2 tobilleras de anclaje y 3 ligas: amarilla 30 lb, roja 20 lb y negra 10 lb.", ""],
  ["Tabla de flexiones", 30, "Ejercicio y deporte", "tabla-de-flexiones.webp", "Ajustable para trabajar diferentes grupos musculares.", ""],
  ["Escalera de agilidad", 17, "Ejercicio y deporte", "escalera-de-agilidad.webp", "7 peldaños. Bolso para transportarla.", ""],
  ["Kit de agilidad", null, "Ejercicio y deporte", "kit-de-agilidad.webp", "Escalera de 12 peldaños, paracaídas con ajuste en cintura, set de 12 conos y bolso.", ""],
  ["Rodilleras de neopreno Aolikes", null, "Ejercicio y deporte", "rodilleras-de-neopreno-aolikes.webp", "Neopreno. Talla XL. Se vende por par.", ""],
  ["Rodilleras de neopreno negras", 25, "Ejercicio y deporte", "rodilleras-de-neopreno-negras.webp", "Neopreno. Se vende por par.", "M, L, XL"],
  ["Guantes de gimnasio fucsia", 13, "Ejercicio y deporte", "guantes-de-gimnasio-fucsia.webp", "Poliéster. Se vende por par.", "M, L"],
  ["Guantes de gimnasio negros", null, "Ejercicio y deporte", "guantes-de-gimnasio-negros.webp", "Poliéster. Se vende por par.", "M, L"],
  ["Guantes de gimnasio rosado claro", 13, "Ejercicio y deporte", "guantes-de-gimnasio-rosado-claro.webp", "Poliéster. Se vende por par.", "S, M, L"],
  ["Guantes de gimnasio negro y gris", 13, "Ejercicio y deporte", "guantes-de-gimnasio-negro-y-gris.webp", "Poliéster. Se vende por par.", "S, M, L"],
  ["Guantes de gimnasio negro con rosado", 13, "Ejercicio y deporte", "guantes-de-gimnasio-negro-con-rosado.webp", "Nailon. Se vende por par.", "M, L"],
  ["Guantes de gimnasio negros de nailon", 13, "Ejercicio y deporte", "guantes-de-gimnasio-negros-de-nailon.webp", "Nailon. Se vende por par.", "M, L"],
  ["Bandas de resistencia (set de 5)", 12, "Ejercicio y deporte", "bandas-de-resistencia-set-de-5.webp", "5 bandas de diferentes resistencias.", ""],
  ["Taco de yoga", 10, "Ejercicio y deporte", "taco-de-yoga.webp", "Bloque de yoga.", ""],
  ["Banda de pull up negra 65 lb", 15, "Ejercicio y deporte", "banda-de-pull-up-negra-65-lb.webp", "Liga de resistencia de 65 libras.", ""],
  ["Banda de pull up roja o rosada 35 lb", 10, "Ejercicio y deporte", "banda-de-pull-up-negra-65-lb.webp", "Liga de resistencia de 35 libras.", "Roja, Rosada"],
  ["Banda de pull up amarilla 15 lb", 8, "Ejercicio y deporte", "banda-de-pull-up-negra-65-lb.webp", "Liga de resistencia de 15 libras.", ""],
  ["Banda de pull up naranja 85-230 lb", 25, "Ejercicio y deporte", "banda-de-pull-up-naranja-85-230-lb.webp", "Liga de resistencia de 85 a 230 libras.", ""],
  ["Banda de pull up azul 65-175 lb", 20, "Ejercicio y deporte", "banda-de-pull-up-naranja-85-230-lb.webp", "Liga de resistencia de 65 a 175 libras.", ""],
  ["Banda de pull up verde 55-125 lb", 17, "Ejercicio y deporte", "banda-de-pull-up-verde-55-125-lb.webp", "Liga de resistencia de 55 a 125 libras.", ""],
  ["Banda de pull up morada 35-85 lb", 15, "Ejercicio y deporte", "banda-de-pull-up-verde-55-125-lb.webp", "Liga de resistencia de 35 a 85 libras.", ""],
  ["Banda de resistencia larga", 11, "Ejercicio y deporte", "banda-de-resistencia-larga.webp", "Resistencia de 25 libras. Se vende por unidad.", ""],
  ["Set de ligas tubulares", 20, "Ejercicio y deporte", "set-de-ligas-tubulares.webp", "5 ligas: amarilla 10 lb, verde 15 lb, roja 20 lb, azul 25 lb y negra 30 lb. Con agarres, tobilleras y anclaje de puerta.", ""],
  ["Barra de fuerza de torsión", 40, "Ejercicio y deporte", "barra-de-fuerza-de-torsion.webp", "50 kilos. Acero cromado.", ""],
  ["Pelota de yoga 55cm", 30, "Ejercicio y deporte", "pelota-de-yoga-55cm.webp", "Pelota de ejercicio de 55cm.", "Azul, Gris, Morado"],
  ["Pelota de yoga 25cm", 7, "Ejercicio y deporte", "pelota-de-yoga-25cm.webp", "Pelota de ejercicio de 25cm.", "Azul, Rosado, Morado"],
  ["Pelota de yoga 65cm", 32, "Ejercicio y deporte", "pelota-de-yoga-65cm.webp", "Pelota de ejercicio de 65cm.", "Azul, Gris, Morado"],
  ["Pelota de yoga 75cm", 34, "Ejercicio y deporte", "pelota-de-yoga-75cm.webp", "Pelota de ejercicio de 75cm.", "Azul, Gris, Morado"],
  ["Muñequera camuflaje", 15, "Ejercicio y deporte", "munequera-camuflaje.webp", "Tela dura elástica rugosa. Con gancho. Se venden por par.", ""],
  ["Muñequeras estampadas", 15, "Ejercicio y deporte", "munequeras-estampadas.webp", "Tela dura elástica rugosa. Con gancho. Se venden por par.", ""],
  ["Cayeras negras", 20, "Ejercicio y deporte", "cayeras-negras.webp", "Semi cuero. Se venden por par.", ""],
  ["Cayeras de semi cuero", 20, "Ejercicio y deporte", "cayeras-de-semi-cuero.webp", "Semi cuero. Se venden por par.", "Cuadrado, Redondo"],
  ["Cayeras de fibra de carbono", 17, "Ejercicio y deporte", "cayeras-de-fibra-de-carbono.webp", "Micro fibra de carbono. Se venden por par. Disponibles en gris y morado.", "S, M, L, XL, XXL"],
  ["Cuerda de saltar TDF", 13, "Ejercicio y deporte", "cuerda-de-saltar-tdf.webp", "Marca TDF. Material de metal.", ""],
  ["Cayeras MKAS modelo 1", 25, "Ejercicio y deporte", "cayeras-mkas-modelo-1.webp", "Máxima protección contra el roce. Negro con amarillo. Micro fibra de carbono reforzada con algodón. Se venden por par.", "S-M, L-XL"],
  ["Cayeras MKAS modelo 2", 20, "Ejercicio y deporte", "cayeras-mkas-modelo-2.webp", "Negro con blanco. Micro fibra de carbono reforzada con algodón. Se venden por par.", "S-M, L-XL"],
  ["Cinturón camuflajeado", 20, "Ejercicio y deporte", "cinturon-camuflajeado.webp", "Gris en S, M, L y XL. Rosado en S y M.", "Gris S, Gris M, Gris L, Gris XL, Rosado S, Rosado M"],
  ["Cayeras MKAS modelo 3", 20, "Ejercicio y deporte", "cayeras-mkas-modelo-3.webp", "Negro con amarillo. Micro fibra de carbono reforzada con algodón. Se venden por par.", "S-M, L-XL"],
  ["Cinturón MKAS", 20, "Ejercicio y deporte", "cinturon-mkas.webp", "Color negro.", "S, M, L, XL"],
  ["Cinturón SKDK", 20, "Ejercicio y deporte", "cinturon-skdk.webp", "Color negro.", "S, M, L, XL"],
  ["Magnesio líquido", 7, "Ejercicio y deporte", "magnesio-liquido.webp", "Magnesio líquido para agarre.", ""],
  ["Magnesio en bloque", 6, "Ejercicio y deporte", "magnesio-en-bloque.webp", "Bloque de magnesio para agarre.", ""],
  ["Cinta métrica 150cm", null, "Consultorio", "cinta-metrica-150cm.webp", "150 cm.", "Rojo, Verde, Azul, Negro, Rosado"],
  ["Adaptadores de acupuntura", 7, "Consultorio", "adaptadores-de-acupuntura.webp", "Se venden por par.", ""],
  ["Kit de glucómetro Sinocare", 27, "Consultorio", "kit-de-glucometro-sinocare.webp", "Incluye 2 pilas AAA, 50 tiras reactivas, 50 lancetas y dispositivo de punción. No sirven tiras de otras marcas.", ""],
  ["Kit de tensiómetro", 25, "Consultorio", "kit-de-tensiometro.webp", "Tensiómetro manual y estetoscopio.", "Morado, Rosado, Negro, Azul"],
  ["Lápiz de acupuntura 4 cabezales", 25, "Consultorio", "lapiz-de-acupuntura-4-cabezales.webp", "Recargable (cable USB). 9 niveles de intensidad. Incluye 2 envases pequeños de gel.", ""],
  ["Lápiz de acupuntura 2 cabezales", 20, "Consultorio", "lapiz-de-acupuntura-4-cabezales.webp", "Recargable (cable USB). 9 niveles de intensidad. Incluye 2 envases pequeños de gel.", ""],
  ["Martillo Buck neurológico 3 en 1", 12, "Consultorio", "martillo-buck-neurologico-3-en-1.webp", "Con brocha y aguja.", ""],
  ["Camilla portátil", 300, "Consultorio", "camilla-portatil.webp", "Pesa 14kg y soporta 250kg. Largo 2.13m, ancho 60cm. Con posa brazos, cabecera y reposa manos.", "Azul, Negro"],
  ["Goniómetro", 8, "Consultorio", "goniometro.webp", "Plástico grueso y resistente.", ""],
  ["Agujas de acupuntura 0.25x13mm (100 unidades)", 12, "Consultorio", "agujas-de-acupuntura-0-100-unidades.webp", "Set de 100 agujas.", ""],
  ["Agujas de acupuntura 0.25x25mm (100 unidades)", 12, "Consultorio", "agujas-de-acupuntura-0-100-unidades.webp", "Set de 100 agujas.", ""],
  ["Agujas de acupuntura 0.25x30mm (100 unidades)", 13, "Consultorio", "agujas-de-acupuntura-0-100-unidades.webp", "Set de 100 agujas.", ""],
  ["Agujas de acupuntura 0.25x40mm (100 unidades)", 13, "Consultorio", "agujas-de-acupuntura-0-100-unidades.webp", "Set de 100 agujas.", ""],
  ["Agujas de acupuntura 0.30x40mm (100 unidades)", 14, "Consultorio", "agujas-de-acupuntura-0-100-unidades.webp", "Set de 100 agujas.", ""],
  ["Agujas de acupuntura 0.30x60mm (100 unidades)", 14, "Consultorio", "agujas-de-acupuntura-0-100-unidades.webp", "Set de 100 agujas.", ""],
  ["Agujas de acupuntura 0.30x75mm (100 unidades)", 14, "Consultorio", "agujas-de-acupuntura-0-100-unidades.webp", "Set de 100 agujas.", ""],
  ["Agujas de acupuntura 0.40x75mm (100 unidades)", 15, "Consultorio", "agujas-de-acupuntura-0-100-unidades.webp", "Set de 100 agujas.", ""]
];
