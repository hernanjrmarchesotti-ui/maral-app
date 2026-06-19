// ============================================================
// MARAL — Google Apps Script Backend
// Pegar este código en el editor de Apps Script del Spreadsheet
// Menú: Extensiones → Apps Script
// Luego ejecutar inicializarHojas() una sola vez, y redesplegar.
// ============================================================

const SS = SpreadsheetApp.getActiveSpreadsheet();

// ─── GET: devuelve todos los datos para la app ───────────────
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || 'getData';
  if (action === 'getData') {
    const data = getAppData();
    return ContentService
      .createTextOutput(JSON.stringify(data))
      .setMimeType(ContentService.MimeType.JSON);
  }
  return ContentService
    .createTextOutput(JSON.stringify({ error: 'Acción no reconocida' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── POST: guarda un registro en la hoja correspondiente ─────
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    processRecord(payload);
    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ─── Lee los datos completos para la app ─────────────────────
function getAppData() {
  const cfg = getConfig();
  return {
    u: cfg.usuario || 'Hernán',
    cfg: cfg,
    lotes: getLotes(),
    kpi: getKPI(),
    sal: getSocios()
  };
}

function getLotes() {
  const sh = getOrCreateSheet('Lotes', ['id','codigo','etapa','animales','peso_entrada','costo_compra','fecha_ingreso','machos','hembras','peso_actual','ultima_pesada']);
  const rows = sh.getDataRange().getValues();
  if (rows.length <= 1) return [];
  return rows.slice(1).map(function(r) {
    return {
      id: r[0], c: r[1], e: r[2], n: Number(r[3]), we: Number(r[4]), cc: Number(r[5]),
      fi: r[6] ? Utilities.formatDate(new Date(r[6]), 'America/Argentina/Buenos_Aires', 'yyyy-MM-dd') : null,
      m: Number(r[7]), h: Number(r[8]), wa: Number(r[9]),
      lp: r[10] ? Utilities.formatDate(new Date(r[10]), 'America/Argentina/Buenos_Aires', 'yyyy-MM-dd') : null
    };
  });
}

function getKPI() {
  const sh = getOrCreateSheet('KPI', ['clave','valor']);
  const rows = sh.getDataRange().getValues();
  const map = {};
  rows.slice(1).forEach(function(r) { map[r[0]] = r[1]; });
  return {
    v: Number(map['ventas_acum']) || 0,
    g: Number(map['gastos_acum']) || 0,
    ks: Number(map['stock_kg']) || 0
  };
}

function getSocios() {
  const sh = getOrCreateSheet('Socios', ['nombre','aportado','retirado','color']);
  const rows = sh.getDataRange().getValues();
  if (rows.length <= 1) return [];
  return rows.slice(1).map(function(r) {
    return { n: r[0], a: Number(r[1]), r: Number(r[2]), col: r[3] };
  });
}

function getConfig() {
  const sh = getOrCreateSheet('Config', ['clave','valor']);
  const rows = sh.getDataRange().getValues();
  const map = {};
  rows.slice(1).forEach(function(r) { map[r[0]] = r[1]; });
  return {
    pm:      Number(map['pm'])      || 5000,
    obj:     Number(map['obj'])     || 340,
    adft:    Number(map['adft'])    || 1.1,
    adrc:    Number(map['adrc'])    || 0.7,
    ckft:    Number(map['ckft'])    || 2720,
    ckrc:    Number(map['ckrc'])    || 1750,
    da:      Number(map['da'])      || 20,
    usuario: map['usuario']         || 'Hernán'
  };
}

// ─── Procesa y guarda un registro ────────────────────────────
function processRecord(p) {
  const ts = new Date();
  const fecha = p.fecha ? new Date(p.fecha) : ts;

  if (p.tipo === 'pesada') {
    appendRow('Pesadas',
      ['fecha','lote_id','lote_codigo','peso_promedio','cantidad','es_estimado','observacion','usuario','timestamp'],
      [fecha, p.lote_id, p.lote_codigo, p.peso_promedio, p.cantidad, p.es_estimado, p.observacion, p.usuario, ts]);
    updateLotePeso(p.lote_id, p.peso_promedio, fecha);
  }

  if (p.tipo === 'alimento') {
    appendRow('Alimentos',
      ['fecha','lote_id','lote_codigo','tipo_alimento','kg_cargados','kg_sobrante','proveedor','precio_total','usuario','timestamp'],
      [fecha, p.lote_id, p.lote_codigo, p.tipo_alimento, p.kg_cargados, p.kg_sobrante, p.proveedor, p.precio_total, p.usuario, ts]);
    if (p.precio_total) sumarKPI('gastos_acum', parseFloat(p.precio_total));
  }

  if (p.tipo === 'gasto') {
    appendRow('Gastos',
      ['fecha','lote_id','lote_codigo','tipo_gasto','descripcion','monto','pagado_por','usuario','timestamp'],
      [fecha, p.lote_id, p.lote_codigo, p.tipo_gasto, p.descripcion, p.monto, p.pagado_por, p.usuario, ts]);
    if (p.monto) sumarKPI('gastos_acum', parseFloat(p.monto));
  }

  if (p.tipo === 'movimiento') {
    appendRow('Movimientos',
      ['fecha','lote_id','lote_codigo','tipo_movimiento','desde','hacia','observacion','usuario','timestamp'],
      [fecha, p.lote_id, p.lote_codigo, p.tipo_movimiento, p.desde, p.hacia, p.observacion, p.usuario, ts]);
  }

  if (p.tipo === 'venta') {
    const kgBrutos = parseFloat(p.kg_brutos) || 0;
    const precioPorKg = parseFloat(p.precio_kg) || 0;
    const desbaste = parseFloat(p.desbaste) || 8;
    const ingresoNeto = kgBrutos * (1 - desbaste / 100) * precioPorKg;
    appendRow('Ventas',
      ['fecha','lote_id','lote_codigo','kg_brutos','precio_kg','desbaste','cantidad','comprador','ingreso_neto','usuario','timestamp'],
      [fecha, p.lote_id, p.lote_codigo, p.kg_brutos, p.precio_kg, p.desbaste, p.cantidad, p.comprador, ingresoNeto, p.usuario, ts]);
    if (ingresoNeto) sumarKPI('ventas_acum', ingresoNeto);
  }

  SpreadsheetApp.flush();
}

// ─── Helpers internos ─────────────────────────────────────────
function updateLotePeso(loteId, nuevoPeso, fecha) {
  var sh = SS.getSheetByName('Lotes');
  if (!sh) return;
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(loteId)) {
      sh.getRange(i + 1, 10).setValue(parseFloat(nuevoPeso)); // columna peso_actual
      sh.getRange(i + 1, 11).setValue(fecha);                 // columna ultima_pesada
      break;
    }
  }
}

function sumarKPI(clave, delta) {
  var sh = getOrCreateSheet('KPI', ['clave','valor']);
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === clave) {
      var actual = parseFloat(data[i][1]) || 0;
      sh.getRange(i + 1, 2).setValue(actual + delta);
      return;
    }
  }
  sh.appendRow([clave, delta]);
}

function getOrCreateSheet(name, headers) {
  var sh = SS.getSheetByName(name);
  if (!sh) {
    sh = SS.insertSheet(name);
    sh.appendRow(headers);
  }
  return sh;
}

function appendRow(sheetName, headers, values) {
  var sh = getOrCreateSheet(sheetName, headers);
  sh.appendRow(values);
}

// ─── INICIALIZAR HOJAS ────────────────────────────────────────
// Ejecutar UNA SOLA VEZ desde el editor de Apps Script:
// seleccioná esta función y presioná ▶ Ejecutar
function inicializarHojas() {
  // Config
  var config = getOrCreateSheet('Config', ['clave','valor']);
  if (config.getLastRow() <= 1) {
    config.getRange(2, 1, 8, 2).setValues([
      ['pm',      5000],
      ['obj',     340],
      ['adft',    1.1],
      ['adrc',    0.7],
      ['ckft',    2720],
      ['ckrc',    1750],
      ['da',      20],
      ['usuario', 'Hernán']
    ]);
  }

  // Socios
  var socios = getOrCreateSheet('Socios', ['nombre','aportado','retirado','color']);
  if (socios.getLastRow() <= 1) {
    socios.getRange(2, 1, 2, 4).setValues([
      ['Paco',   14200000, 0, '#64b5f6'],
      ['Hernán', 15800000, 0, '#81c784']
    ]);
  }

  // KPI
  var kpi = getOrCreateSheet('KPI', ['clave','valor']);
  if (kpi.getLastRow() <= 1) {
    kpi.getRange(2, 1, 3, 2).setValues([
      ['ventas_acum', 50960500],
      ['gastos_acum', 11983777],
      ['stock_kg',    7080]
    ]);
  }

  // Lotes con datos iniciales
  var lotes = getOrCreateSheet('Lotes', ['id','codigo','etapa','animales','peso_entrada','costo_compra','fecha_ingreso','machos','hembras','peso_actual','ultima_pesada']);
  if (lotes.getLastRow() <= 1) {
    lotes.getRange(2, 1, 9, 11).setValues([
      [1,'RC-2605-01',   'Recría',  11, 160.5, 10592400, new Date('2026-05-11'), 5, 6, 160,   null],
      [2,'RC-2604-01',   'Recría',   5, 184,    3896200, new Date('2026-04-29'), 2, 3, 190,   null],
      [3,'RC-2602-02',   'Recría',   3, 239.3,  2871200, new Date('2026-02-06'), 0, 3, 239.3, null],
      [4,'RC-2603-01-RC','Recría',   1, 206.7,  1033333, new Date('2026-03-18'), 0, 1, 206.7, null],
      [5,'RC-2602-01',   'Recría',   1, 200,     700000, new Date('2026-02-06'), 0, 1, 200,   null],
      [6,'RC-2601-02',   'Recría',   1, 140,     490000, new Date('2026-01-16'), 1, 0, 140,   null],
      [7,'RC-2601-01',   'Recría',   1, 140,     500000, new Date('2026-01-30'), 1, 0, 140,   null],
      [8,'FT-ACTUAL-01', 'Feedlot',  8, 215.8,  7769700, new Date('2026-03-20'), 2, 6, 215.8, null],
      [9,'RC-2603-01',   'Feedlot',  2, 206.7,  2066667, new Date('2026-03-18'), 0, 2, 206.7, null]
    ]);
  }

  // Crear hojas de registro vacías
  getOrCreateSheet('Pesadas',    ['fecha','lote_id','lote_codigo','peso_promedio','cantidad','es_estimado','observacion','usuario','timestamp']);
  getOrCreateSheet('Alimentos',  ['fecha','lote_id','lote_codigo','tipo_alimento','kg_cargados','kg_sobrante','proveedor','precio_total','usuario','timestamp']);
  getOrCreateSheet('Gastos',     ['fecha','lote_id','lote_codigo','tipo_gasto','descripcion','monto','pagado_por','usuario','timestamp']);
  getOrCreateSheet('Movimientos',['fecha','lote_id','lote_codigo','tipo_movimiento','desde','hacia','observacion','usuario','timestamp']);
  getOrCreateSheet('Ventas',     ['fecha','lote_id','lote_codigo','kg_brutos','precio_kg','desbaste','cantidad','comprador','ingreso_neto','usuario','timestamp']);

  SpreadsheetApp.flush();
  Logger.log('✅ Hojas inicializadas correctamente. Ahora redesplegar el Web App.');
}
