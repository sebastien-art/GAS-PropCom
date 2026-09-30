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

  var maxCols = sheet.getLastColumn();
  // Escaneamos las primeras 30 filas (o hasta la última fila con datos)
  var maxFilasBusqueda = Math.min(30, sheet.getLastRow()); 
  
  if (maxCols === 0 || maxFilasBusqueda === 0) {
    Logger.log("❌ La hoja está vacía.");
    return;
  }

  // Cargar matriz completa en memoria para rápido escaneo
  var datosHeader = sheet.getRange(1, 1, maxFilasBusqueda, maxCols).getValues();

  var colRef = -1;
  var colPlantilla = -1;
  var colFicha = -1;
  var filaEncabezadoEncontrada = -1;

  // Recorrer filas y columnas buscando las etiquetas
  for (var r = 0; r < datosHeader.length; r++) {
    for (var c = 0; c < maxCols; c++) {
      var val = datosHeader[r][c].toString().trim().toUpperCase();
      
      if (val === "REF" || val.indexOf("REF") !== -1) {
        if (colRef === -1) { colRef = c + 1; filaEncabezadoEncontrada = r + 1; }
      }
      if (val === "PLANTILLA" || val.indexOf("PLANTILLA") !== -1) {
        if (colPlantilla === -1) colPlantilla = c + 1;
      }
      if (val === "FICHA" || val.indexOf("FICHA") !== -1) {
        if (colFicha === -1) colFicha = c + 1;
      }
    }
  }

  Logger.log("📌 Encabezados detectados en la Fila: " + (filaEncabezadoEncontrada !== -1 ? filaEncabezadoEncontrada : "No detectada"));
  Logger.log("📌 Ubi. Encabezado 'REF': Columna " + colRef);
  Logger.log("📌 Ubi. Encabezado 'Plantilla': Columna " + colPlantilla);
  Logger.log("📌 Ubi. Encabezado 'Ficha': Columna " + colFicha);

  // Extraer valores de la fila seleccionada
  var valRef = (colRef !== -1) ? sheet.getRange(fila, colRef).getValue().toString().trim() : "";
  var valPlantilla = (colPlantilla !== -1) ? sheet.getRange(fila, colPlantilla).getValue().toString().trim() : "";
  var valFicha = (colFicha !== -1) ? sheet.getRange(fila, colFicha).getValue().toString().trim() : "";

  Logger.log("📝 Valor 'REF': " + valRef);
  Logger.log("📝 Valor 'Plantilla': " + valPlantilla);
  Logger.log("📝 Valor 'Ficha': " + valFicha);
  Logger.log("==================================================");
}