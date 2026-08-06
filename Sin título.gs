function diagnosticarPorEncabezados() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var celdaActiva = sheet.getActiveCell();

  if (!celdaActiva) {
    Logger.log("❌ No hay ninguna celda seleccionada.");
    return;
  }

  var fila = celdaActiva.getRow();

  Logger.log("==================================================");
  Logger.log("🔍 DIAGNÓSTICO EN PROPUESTAS COMERCIALES | FILA: " + fila);
  Logger.log("==================================================");

  // 1. Escanear encabezados en las filas 23 y 24 para encontrar las columnas exactas
  var colRef = -1;
  var colPlantilla = -1;
  var colFicha = -1;

  var maxCols = sheet.getLastColumn();
  
  for (var c = 1; c <= maxCols; c++) {
    var val23 = sheet.getRange(23, c).getValue().toString().trim().toUpperCase();
    var val24 = sheet.getRange(24, c).getValue().toString().trim().toUpperCase();
    
    if (val23 === "REF" || val24 === "REF") colRef = c;
    if (val23 === "PLANTILLA" || val24 === "PLANTILLA") colPlantilla = c;
    if (val23 === "FICHA" || val24 === "FICHA") colFicha = c;
  }

  Logger.log("📌 Ubi. Encabezado 'REF': Columna " + colRef);
  Logger.log("📌 Ubi. Encabezado 'Plantilla': Columna " + colPlantilla);
  Logger.log("📌 Ubi. Encabezado 'Ficha': Columna " + colFicha);

  // 2. Extraer valores de la fila seleccionada
  var valRef = (colRef !== -1) ? sheet.getRange(fila, colRef).getValue().toString().trim() : "";
  var valPlantilla = (colPlantilla !== -1) ? sheet.getRange(fila, colPlantilla).getValue().toString().trim() : "";
  var valFicha = (colFicha !== -1) ? sheet.getRange(fila, colFicha).getValue().toString().trim() : "";

  Logger.log("\n📄 Valor leído en REF: " + valRef);
  Logger.log("📄 Valor leído en Plantilla: " + valPlantilla);
  Logger.log("📄 Valor leído en Ficha: " + valFicha);

  // 3. Extraer el código NXXXX
  var textoUnificado = valRef + " " + valPlantilla + " " + valFicha;
  var coincidenciaRef = textoUnificado.match(/N\d+/i);

  if (!coincidenciaRef) {
    Logger.log("\n❌ No se encontró ninguna etiqueta con formato NXXXX en esta fila.");
    return;
  }

  var refFinal = coincidenciaRef[0].toUpperCase();
  Logger.log("\n🎯 Referencia detectada para buscar en Drive: " + refFinal);

  // 4. Búsqueda directa en Google Drive
  Logger.log("🔎 Buscando archivo Google Slides en Drive...");
  
  var query = "mimeType = 'application/vnd.google-apps.presentation' and trashed = false and title contains '" + refFinal + "'";
  var archivos = DriveApp.searchFiles(query);
  
  var total = 0;
  while (archivos.hasNext()) {
    var archivo = archivos.next();
    total++;
    Logger.log("   [" + total + "] " + archivo.getName() + " | ID: " + archivo.getId());
  }

  if (total === 0) {
    Logger.log("❌ No se encontró la plantilla para " + refFinal + " en Drive.");
  } else {
    Logger.log("✅ Se encontraron " + total + " coincidencia(s) en Drive.");
  }
  Logger.log("==================================================");
}