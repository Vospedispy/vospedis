function doGet(e) {
  return handleRequest(e);
}
function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  var output;
  try {
    var action = '';
    var body = {};
    
    if (e.postData) {
      body = JSON.parse(e.postData.contents);
      action = body.action || '';
    } else if (e.parameter) {
      action = e.parameter.action || '';
      body = e.parameter;
    }

    var result = {};
    if (action === 'getVentas')   result = getVentas();
    if (action === 'addVenta')    result = addVenta(body.venta);
    if (action === 'updateVenta') result = updateVenta(body.venta);
    if (action === 'deleteVenta') result = deleteVenta(body.id);
    if (action === 'getCostos')   result = getCostos();
    if (action === 'saveCosto')   result = saveCosto(body.id, body.costo);
    if (action === 'getConfig')   result = getConfig();
    if (action === 'saveConfig')  result = saveConfig(body.key, body.value);

    output = ContentService
      .createTextOutput(JSON.stringify({ ok: true, data: result }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    output = ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  return output;
}

function getSheet(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (name === 'Ventas') {
      sh.appendRow(['id','fecha','cliente','ci','tel','ciudad','marca','categoria','producto','imei','color','precio','pago','estado','notas','ts']);
      sh.setFrozenRows(1);
    }
    if (name === 'Costos') {
      sh.appendRow(['id','costo','ts']);
      sh.setFrozenRows(1);
    }
    if (name === 'Config') {
      sh.appendRow(['key','value']);
      sh.setFrozenRows(1);
    }
  }
  return sh;
}

function getVentas() {
  var sh = getSheet('Ventas');
  var data = sh.getDataRange().getValues();
  if (data.length <= 1) return [];
  var headers = data[0];
  return data.slice(1).map(function(row) {
    var obj = {};
    headers.forEach(function(h, i) { obj[h] = row[i]; });
    return obj;
  }).reverse();
}

function addVenta(v) {
  var sh = getSheet('Ventas');
  sh.appendRow([v.id,v.fecha,v.cliente,v.ci,v.tel,v.ciudad,v.marca,v.categoria,v.producto,v.imei,v.color,v.precio,v.pago,v.estado,v.notas,new Date().toISOString()]);
  return { id: v.id };
}

function updateVenta(v) {
  var sh = getSheet('Ventas');
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(v.id)) {
      sh.getRange(i+1,1,1,16).setValues([[v.id,v.fecha,v.cliente,v.ci,v.tel,v.ciudad,v.marca,v.categoria,v.producto,v.imei,v.color,v.precio,v.pago,v.estado,v.notas,new Date().toISOString()]]);
      return { updated: true };
    }
  }
  return { updated: false };
}

function deleteVenta(id) {
  var sh = getSheet('Ventas');
  var data = sh.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === String(id)) {
      sh.deleteRow(i + 1);
      return { deleted: true };
    }
  }
  return { deleted: false };
}

function getCostos() {
  var sh = getSheet('Costos');
  var data = sh.getDataRange().getValues();
  var result = {};
  for (var i = 1; i < data.length; i++) {
    result[String(data[i][0])] = data[i][1];
  }
  return result;
}

function saveCosto(id, costo) {
  var sh = getSheet('Costos');
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sh.getRange(i+1, 2).setValue(costo);
      sh.getRange(i+1, 3).setValue(new Date().toISOString());
      return { saved: true };
    }
  }
  sh.appendRow([id, costo, new Date().toISOString()]);
  return { saved: true };
}

function getConfig() {
  var sh = getSheet('Config');
  var data = sh.getDataRange().getValues();
  var result = {};
  for (var i = 1; i < data.length; i++) {
    result[data[i][0]] = data[i][1];
  }
  return result;
}

function saveConfig(key, value) {
  var sh = getSheet('Config');
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      sh.getRange(i+1, 2).setValue(value);
      return { saved: true };
    }
  }
  sh.appendRow([key, value]);
  return { saved: true };
}
