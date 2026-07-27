/*************************************************************
 * CREACIÓN AUTOMÁTICA DE REFERENCIAS NUEVAS DESDE LA TABLA
 *************************************************************/

function procesarReferenciasNuevas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  const INVENTARIO_WEB_APP_URL = "https://script.google.com/macros/s/AKfycbzpUroVXu87JyY05EZ9MvF9gc1vI5ljsQ-gPgDgANIMkVwMoVxe88L7EghjFTdrn3pUxA/exec";

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  // 1. LOCALIZAR LA TABLA DE PROPUESTAS (Encabezado 'Partida')
  let headerRowIdx = -1;
  let partidaColIdx = -1;

  for (let r = 1; r <= lastRow; r++) {
    let rowVals = sheet.getRange(r, 1, 1, lastCol).getValues()[0];
    let idx = rowVals.findIndex(v => String(v).trim().toLowerCase() === "partida");
    if (idx !== -1) {
      headerRowIdx = r;
      partidaColIdx = idx + 1; // Base 1
      break;
    }
  }

  if (headerRowIdx === -1) {
    SpreadsheetApp.getUi().alert("Error", "No se encontró la tabla de propuestas (falta encabezado 'Partida').", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  // 2. MAPEAR COLUMNAS DINÁMICAMENTE POR ENCABEZADO EN LA TABLA
  const headers = sheet.getRange(headerRowIdx, partidaColIdx, 1, lastCol - partidaColIdx + 1).getValues()[0]
                         .map(h => String(h).trim().toLowerCase());

  const refOffset = headers.findIndex(h => h === "ref");
  const sugM2Offset = headers.findIndex(h => h.includes("superficie sugerida"));
  const parqueOffset = headers.findIndex(h => h === "parque");
  const opOffset = headers.findIndex(h => h.includes("operaci") || h.includes("renta") || h.includes("venta") || h.includes("tipo"));
  const fichaOffset = headers.findIndex(h => h === "ficha");
  const naveOffset = headers.findIndex(h => h === "nave");

  if (refOffset === -1 || sugM2Offset === -1 || parqueOffset === -1) {
    SpreadsheetApp.getUi().alert("Error", "No se encontraron las columnas necesarias ('REF', 'Superficie sugerida (m2)', 'Parque') en la tabla.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const colRef = partidaColIdx + refOffset;
  const colSugM2 = partidaColIdx + sugM2Offset;
  const colParque = partidaColIdx + parqueOffset;
  const colOp = (opOffset !== -1) ? (partidaColIdx + opOffset) : -1;
  const colFicha = (fichaOffset !== -1) ? (partidaColIdx + fichaOffset) : -1;
  const colNave = (naveOffset !== -1) ? (partidaColIdx + naveOffset) : -1;

  let creadasCount = 0;

  // 3. RECORRER FILAS DE LA TABLA
  for (let r = headerRowIdx + 1; r <= lastRow; r++) {
    let valPartida = sheet.getRange(r, partidaColIdx).getValue();
    if (!valPartida) break; // Fin de la tabla

    let valRef = String(sheet.getRange(r, colRef).getValue() || "").trim();
    let rawSugM2 = sheet.getRange(r, colSugM2).getValue();
    let numSugM2 = parseFloat(String(rawSugM2).replace(/[^\d.]/g, "")) || 0;
    let valParque = String(sheet.getRange(r, colParque).getValue() || "").trim();
    
    // Leer la Operación de ESA fila en particular (si existe la columna)
    let valOp = (colOp !== -1) ? String(sheet.getRange(r, colOp).getValue() || "").trim() : "";

    // REGLA 1: Si NO tiene Superficie Sugerida, SE OMITE.
    if (!numSugM2) continue;

    // REGLA 2: SOLO procesar si REF está VACÍA o dice estrictamente "CREAR"
    let esElegibleParaCrear = (valRef === "" || valRef.toUpperCase() === "CREAR");

    if (esElegibleParaCrear && valParque) {
      // Enviamos Parque, m2 y la Operación individual de la fila a la Web App del Inventario
      let url = INVENTARIO_WEB_APP_URL 
        + "?action=crearNuevaRef"
        + "&parque=" + encodeURIComponent(valParque)
        + "&m2=" + encodeURIComponent(numSugM2)
        + "&operacion=" + encodeURIComponent(valOp);

      try {
        let response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
        let json = JSON.parse(response.getContentText());

        if (json.status === "success") {
          let cellRef = sheet.getRange(r, colRef);
          cellRef.clearDataValidations(); 
          cellRef.setValue(json.nuevaRef);
          cellRef.setFontLine("none").setFontColor("#000000");

          if (colFicha !== -1) {
            sheet.getRange(r, colFicha).setValue("");
          }
          if (colNave !== -1) {
            sheet.getRange(r, colNave).setValue("");
          }

          creadasCount++;
        } else {
          Logger.log(`Error en fila ${r} (Parque: ${valParque}): ${json.message}`);
        }
      } catch (err) {
        Logger.log(`Excepción al conectar en fila ${r}: ${err.toString()}`);
      }
    }
  }

  if (creadasCount > 0) {
    SpreadsheetApp.getUi().alert("Proceso Finalizado", `Se crearon e insertaron ${creadasCount} nueva(s) referencia(s) en el Inventario y en esta tabla.`, SpreadsheetApp.getUi().ButtonSet.OK);
  } else {
    SpreadsheetApp.getUi().alert("Aviso", "No se encontraron renglones con Superficie Sugerida que requieran crear una nueva REF.", SpreadsheetApp.getUi().ButtonSet.OK);
  }
}