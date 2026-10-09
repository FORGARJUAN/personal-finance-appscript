/**
 * Code.gs - Contabilidad 2026 (Acordeones, Filtros Historial y Tab Educación)
 */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('Contabilidad 2026')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// --------------------------------------------------------
// LECTURA Y CREACIÓN DE CONFIGURACIÓN
// --------------------------------------------------------
function obtenerConfiguracion() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetConfig = ss.getSheetByName("Config");

  if (!sheetConfig) sheetConfig = crearHojaConfiguracion(ss);

  var sheetGastos = ss.getSheetByName("Gastos");
  if (!sheetGastos) {
    sheetGastos = ss.insertSheet("Gastos");
    sheetGastos.appendRow(["Fecha", "Mes", "Tipo", "Método", "Categoría", "Subcategoría", "Concepto", "Monto", "Evidencia", "Usuario Registra"]);
    sheetGastos.getRange("A1:J1").setFontWeight("bold").setBackground("#0d6efd").setFontColor("white");
    sheetGastos.setFrozenRows(1);
    sheetGastos.autoResizeColumns(1, 10);
  }

  var sheetProy = ss.getSheetByName("Proyecciones");
  if (!sheetProy) {
    sheetProy = ss.insertSheet("Proyecciones");
    sheetProy.appendRow(["Subcategoría", "Día Límite", "Meta Mensual", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]);
    sheetProy.getRange("A1:O1").setFontWeight("bold").setBackground("#198754").setFontColor("white");
    sheetProy.setFrozenRows(1);
  }

  var pin = sheetConfig.getRange("B1").getValue().toString();
  var nombre = sheetConfig.getRange("B2").getValue().toString() || "Usuario";
  var ultimaFila = sheetConfig.getLastRow();
  
  var opciones = {}; var tipos = []; var metodos = []; var categoriasMaestras = [];
  
  if (ultimaFila >= 4) {
    var data = sheetConfig.getRange(4, 1, ultimaFila - 3, 5).getValues();
    data.forEach(function(row) {
      if (row[0]) {
        categoriasMaestras.push(row[0].toString().trim());
        opciones[row[0]] = row[1] ? row[1].toString().split(",").map(s => s.trim()) : [];
      }
      if (row[3]) tipos.push(row[3].toString().trim());
      if (row[4]) metodos.push(row[4].toString().trim());
    });
  }

  if(tipos.length === 0) tipos = ["Hogar", "Personal"];
  if(metodos.length === 0) metodos = ["Efectivo", "Nequi/Débito", "Tarjeta Crédito"];

  return { pin: pin, nombre: nombre, opciones: opciones, tipos: tipos, metodos: metodos, categoriasMaestras: categoriasMaestras };
}

function crearHojaConfiguracion(ss) {
  var sheet = ss.insertSheet("Config");
  sheet.getRange("A1").setValue("PIN de Seguridad").setFontWeight("bold").setBackground("#ffc107");
  sheet.getRange("B1").setValue("0000").setHorizontalAlignment("center"); 
  sheet.getRange("A2").setValue("Nombre Usuario").setFontWeight("bold").setBackground("#0dcaf0");
  sheet.getRange("B2").setValue("Tu Nombre").setHorizontalAlignment("center");
  sheet.getRange("A3:E3").setValues([["Categoría", "Subcategorías (separadas por coma)", "", "Tipos (Perfiles)", "Métodos de Pago"]])
       .setFontWeight("bold").setBackground("#343a40").setFontColor("white");

  var configInicial = [
    ["Gastos Variables", "Alimentación, Transporte, Ocio, Imprevistos, Salud, Otros", "", "Hogar", "Efectivo"],
    ["Gastos Fijos", "Arriendo, Periódico, Internet, Plan Datos, Servicios Públicos", "", "Personal", "Nequi/Débito"],
    ["Deuda", "Tarjetas de Crédito, Préstamos Personales", "", "Negocio", "Tarjeta Crédito"],
    ["Ingresos", "Salario, Ingreso Extra, Ventas", "", "", "Transferencia"],
    ["Ahorro", "Fondo Emergencia, Inversiones", "", "", ""]
  ];

  sheet.getRange(4, 1, configInicial.length, 5).setValues(configInicial);
  sheet.autoResizeColumns(1, 5); sheet.setColumnWidth(3, 20); sheet.setFrozenRows(3);
  return sheet;
}

// --------------------------------------------------------
// MODIFICACIÓN DE CONFIGURACIÓN (Gestor Central)
// --------------------------------------------------------
function cambiarNombreUsuario(nuevoNombre, pinIngresado) {
  var config = obtenerConfiguracion();
  if (pinIngresado !== config.pin) throw new Error("⛔ PIN incorrecto.");
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Config").getRange("B2").setValue(nuevoNombre);
  return "Nombre actualizado a " + nuevoNombre;
}

function agregarElementoConfig(tipoElemento, valor1, valor2, pinIngresado) {
  var config = obtenerConfiguracion();
  if (pinIngresado !== config.pin) throw new Error("⛔ PIN incorrecto.");
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Config");
  var data = sheet.getDataRange().getValues();

  if (tipoElemento === 'categoria') {
    let row = 4; while(row <= data.length && sheet.getRange(row, 1).getValue() !== "") row++;
    sheet.getRange(row, 1).setValue(valor1);
    return "✅ Categoría general agregada.";
  } 
  else if (tipoElemento === 'subcategoria') {
    for (var i = 3; i < data.length; i++) {
      if (data[i][0] === valor1) {
        var actual = data[i][1] ? data[i][1].toString().trim() : "";
        var nuevoVal = actual ? actual + ", " + valor2 : valor2;
        sheet.getRange(i + 1, 2).setValue(nuevoVal);
        return "✅ Subcategoría agregada a " + valor1;
      }
    }
    throw new Error("Categoría maestra no encontrada.");
  }
  else if (tipoElemento === 'perfil') {
    let row = 4; while(row <= data.length && sheet.getRange(row, 4).getValue() !== "") row++;
    sheet.getRange(row, 4).setValue(valor1);
    return "✅ Perfil (De quién es) agregado.";
  }
  else if (tipoElemento === 'metodo') {
    let row = 4; while(row <= data.length && sheet.getRange(row, 5).getValue() !== "") row++;
    sheet.getRange(row, 5).setValue(valor1);
    return "✅ Método de pago agregado.";
  }
}

// --------------------------------------------------------
// PROYECCIONES Y PRESUPUESTO
// --------------------------------------------------------
function autoConstruirProyecciones(pinIngresado) {
  var config = obtenerConfiguracion(); 
  if (pinIngresado !== config.pin) throw new Error("⛔ PIN incorrecto.");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetProy = ss.getSheetByName("Proyecciones");
  var sheetConfig = ss.getSheetByName("Config");
  
  sheetProy.clear();
  var meses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];
  var encabezados = ["Subcategoría", "Día Límite", "Meta Mensual"].concat(meses);
  sheetProy.appendRow(encabezados);
  sheetProy.getRange("A1:O1").setFontWeight("bold").setBackground("#198754").setFontColor("white");
  sheetProy.setFrozenRows(1);
  
  var dataConfig = sheetConfig.getDataRange().getValues();
  var items = [];
  for (var i = 3; i < dataConfig.length; i++) {
    var subcatString = dataConfig[i][1] ? dataConfig[i][1].toString() : "";
    var subs = subcatString.split(",");
    subs.forEach(function(s) { if (s.trim() !== "") items.push(s.trim()); });
  }
  items = [...new Set(items)]; 
  
  var matrix = [];
  for (var i = 0; i < items.length; i++) {
    var filaActual = i + 2;
    var fila = [items[i], "", 0]; 
    for (var m = 0; m < 12; m++) {
       var letraMes = String.fromCharCode(68 + m); 
       var formula = "=SUMIFS(Gastos!$H:$H, Gastos!$B:$B, " + letraMes + "$1, Gastos!$F:$F, $A" + filaActual + ")";
       fila.push(formula);
    }
    matrix.push(fila);
  }
  
  if (matrix.length > 0) {
    sheetProy.getRange(2, 1, matrix.length, 15).setValues(matrix);
    sheetProy.getRange(2, 3, matrix.length, 13).setNumberFormat("$#,##0");
  }
  sheetProy.autoResizeColumns(1, 15);
  return "✅ ¡La hoja de Proyecciones ha sido programada con éxito!";
}

function obtenerListaProyecciones() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Proyecciones");
  if (!sheet) return [];
  var data = sheet.getRange(2, 1, Math.max(1, sheet.getLastRow() - 1), 3).getValues();
  var lista = [];
  data.forEach(function(row) { if (row[0]) lista.push({ nombre: row[0], meta: row[2] }); });
  return lista;
}

function actualizarMetaProyeccion(subcat, monto, pinIngresado) {
  var config = obtenerConfiguracion(); 
  if (pinIngresado !== config.pin) throw new Error("⛔ PIN incorrecto.");
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Proyecciones");
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === subcat) {
      sheet.getRange(i + 1, 3).setValue(monto);
      return "Meta actualizada para " + subcat;
    }
  }
  throw new Error("❌ No se encontró la subcategoría en Proyecciones. Genera la tabla primero.");
}

// --------------------------------------------------------
// GUARDAR Y GESTIONAR GASTOS
// --------------------------------------------------------
function guardarGasto(datos) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) throw new Error("⏳ Sistema ocupado.");
  try {
    var config = obtenerConfiguracion(); 
    if (datos.pin !== config.pin) throw new Error("⛔ PIN incorrecto.");

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Gastos");
    var usuario = Session.getActiveUser().getEmail() || "Usuario";
    var now = new Date(); 
    var meses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];
    var mesNombre = meses[now.getMonth()]; 

    var fileUrl = "";
    if (datos.archivo && datos.archivo.data) {
      var dataArchivo = Utilities.base64Decode(datos.archivo.data);
      var blob = Utilities.newBlob(dataArchivo, datos.archivo.mimeType, datos.archivo.nombre);
      var rootFolder = getOrCreateFolder(DriveApp.getRootFolder(), "Evidencias Contables");
      var monthFolder = getOrCreateFolder(rootFolder, mesNombre);
      fileUrl = monthFolder.createFile(blob).getUrl();
    }
    sheet.appendRow([now, mesNombre, datos.tipo, datos.metodo, datos.categoria, datos.subcategoria, datos.concepto, datos.monto, fileUrl, usuario]);
    return "✅ Guardado correctamente.";
  } catch (error) { throw new Error(error.toString()); } finally { lock.releaseLock(); }
}

function getOrCreateFolder(parent, name) {
  var folders = parent.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : parent.createFolder(name);
}

function obtenerHistorial() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Gastos");
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues(); // Usar getValues para extraer Fechas Reales (Date objects)
  var historial = [];
  
  // Ampliado a 60 para alimentar bien el gráfico
  for (var i = data.length - 1; i > 0 && historial.length < 60; i--) {
    var f = new Date(data[i][0]);
    var fechaISO = isNaN(f) ? "" : f.toISOString(); // Formato seguro para el gráfico
    
    historial.push({ fila: i + 1, fecha: fechaISO, mes: data[i][1], tipo: data[i][2], metodo: data[i][3], categoria: data[i][4], subcategoria: data[i][5], concepto: data[i][6], monto: data[i][7] });
  }
  return historial;
}

function editarRegistro(fila, datosEditados) {
  var config = obtenerConfiguracion(); 
  if (datosEditados.pin !== config.pin) throw new Error("⛔ PIN incorrecto.");
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Gastos").getRange(fila, 7, 1, 2).setValues([[datosEditados.concepto, datosEditados.monto]]);
  return "✅ Registro actualizado.";
}

function eliminarRegistro(fila, pinIngresado) {
  var config = obtenerConfiguracion(); 
  if (pinIngresado !== config.pin) throw new Error("⛔ PIN incorrecto.");
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Gastos").deleteRow(fila);
  return "🗑️ Registro eliminado con éxito.";
}

// --------------------------------------------------------
// IMPORTAR / EXPORTAR DATOS
// --------------------------------------------------------
function obtenerUrlExportacionXLSX() { return "https://docs.google.com/spreadsheets/d/" + SpreadsheetApp.getActiveSpreadsheet().getId() + "/export?format=xlsx"; }
function exportarDatosCSV() {
  var data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Gastos").getDataRange().getDisplayValues();
  var csv = ""; for (var i = 0; i < data.length; i++) csv += data[i].map(c => '"' + c.replace(/"/g, '""') + '"').join(",") + "
"; return csv;
}
function importarDatosCSV(csvString, pinIngresado) {
  var config = obtenerConfiguracion(); if (pinIngresado !== config.pin) throw new Error("⛔ PIN incorrecto.");
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Gastos");
  var rows = csvString.split("\n");
  for (var i = 1; i < rows.length; i++) { 
    if(rows[i].trim() === "") continue;
    var cols = rows[i].split(",").map(c => c.replace(/(^"|"$)/g, '').replace(/""/g, '"'));
    if (cols.length >= 8) sheet.appendRow(cols);
  }
  return "✅ Base importada.";
}

// --------------------------------------------------------
// DASHBOARD AVANZADO
// --------------------------------------------------------
function obtenerDatosDashboardAvanzado(filtros) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var data = ss.getSheetByName("Gastos").getDataRange().getValues();
  var configData = ss.getSheetByName("Config").getDataRange().getValues();
  
  var checklist = { "Gastos Fijos": [], "Deuda": [], "Ahorro": [] };
  
  for (var i = 3; i < configData.length; i++) {
    var cat = configData[i][0];
    if (checklist[cat] !== undefined) {
       var subs = configData[i][1] ? configData[i][1].toString().split(",") : [];
       subs.forEach(function(s) {
         if(s.trim() !== "") checklist[cat].push({ subcategoria: s.trim(), pagado: false, montoPagado: 0 });
       });
    }
  }

  var ingresos = 0; var gastos = 0; var ahorro = 0; var desgloseGastos = {};

  for (var i = 1; i < data.length; i++) {
    var fechaObj = new Date(data[i][0]);
    if (isNaN(fechaObj)) continue;
    var mesStr = data[i][1]; var anioStr = fechaObj.getFullYear().toString();
    var cat = data[i][4]; var subcat = data[i][5]; var monto = parseFloat(data[i][7]) || 0;

    if (filtros.anio !== "Todos" && anioStr !== filtros.anio) continue;
    if (filtros.mes !== "Todos" && mesStr !== filtros.mes) continue;

    if (checklist[cat]) {
      checklist[cat].forEach(function(obl) {
        if (obl.subcategoria === subcat) { obl.pagado = true; obl.montoPagado += monto; }
      });
    }

    if (cat === "Ingresos") { ingresos += monto; } 
    else if (cat === "Ahorro") { ahorro += monto; } 
    else {
      gastos += monto;
      if (!desgloseGastos[cat]) desgloseGastos[cat] = 0;
      desgloseGastos[cat] += monto;
    }
  }
  return { ingresos: ingresos, gastos: gastos, balance: ingresos - gastos - ahorro, labelsGrafico: Object.keys(desgloseGastos), datosGrafico: Object.values(desgloseGastos), checklistJerarquico: checklist };
}
