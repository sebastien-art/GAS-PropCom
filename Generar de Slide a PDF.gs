function GenerarPDF() {
  const props = PropertiesService.getDocumentProperties();
  const presId = props.getProperty("LAST_SLIDES_ID");
  
  // 1. Validar que exista una presentación previa
  if (!presId) throw new Error('No hay "último Slides" detectado. Ejecuta CrearEnSlides() primero.');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  
  // 2. Usar SIEMPRE el nombre de la pestaña actual
  const sheetName = sheet.getName();
  
  // 3. Obtener o crear la carpeta basada estrictamente en la pestaña vigente
  const folder = getOrCreateSheetFolderInRoot_(sheetName);
  
  const fechaPdf = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd.MM.yy");
  const pdfName = `Propuestas de Naves IEM ${sheetName} ${fechaPdf}.pdf`.replace(/[\/\\:*?"<>|]/g, " ");

  // 4. Generar el PDF
  const pdfBlob = DriveApp.getFileById(presId).getAs(MimeType.PDF).setName(pdfName);
  
  // 5. Limpiar versiones anteriores del PDF en esa carpeta
  trashFilesByName_(folder, pdfName);
  const pdfFile = folder.createFile(pdfBlob);

  // 6. Localizar dónde escribir el link (Basado en tu columna D / outCol 4)
  const outCol = 4;
  let outRow = null;

  const finder = sheet.createTextFinder("Slides:").matchEntireCell(true);
  const matches = finder.findAll();
  
  if (matches.length) {
    const inCol = matches.filter(m => m.getColumn() === outCol);
    if (inCol.length) {
      // Escribir 2 filas debajo de "Slides:"
      outRow = inCol[inCol.length - 1].getRow() + 2;
    }
  }

  // Fallback si no encontró la celda "Slides:"
  if (!outRow) outRow = sheet.getLastRow() + 2;

  // 7. Asegurar espacio en la hoja y escribir el link
  if (outRow > sheet.getMaxRows()) {
    sheet.insertRowsAfter(sheet.getMaxRows(), outRow - sheet.getMaxRows());
  }

  sheet.getRange(outRow, outCol, 1, 2).clearContent();
  sheet.getRange(outRow, outCol).setValue("PDF:").setFontWeight("bold");
  sheet.getRange(outRow, outCol + 1).setFormula(`=HYPERLINK("${pdfFile.getUrl()}","Abrir PDF")`);

  ss.toast(`PDF generado en carpeta "${sheetName}".`, "Éxito", 5);
}