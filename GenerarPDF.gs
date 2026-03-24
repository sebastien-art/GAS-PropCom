/***************************************************************
 * 2) GENERAR PDF
 ***************************************************************/
function GenerarPDF() {
  const props = PropertiesService.getDocumentProperties();
  const presId = props.getProperty(PROP_LAST_SLIDES_ID);
  const anchor = props.getProperty(PROP_LAST_LINK_ANCHOR);
  const sheetNameStored = props.getProperty(PROP_LAST_SHEET_NAME);
  const folderId = props.getProperty(PROP_LAST_FOLDER_ID);

  if (!presId) throw new Error('No hay "último Slides". Primero ejecuta CrearEnSlides().');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  const sheetName = sheetNameStored || sheet.getName();
  const folder = folderId ? DriveApp.getFolderById(folderId) : getOrCreateSheetFolderInRoot_(sheetName);

  const fechaPdf = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd.MM.yy");
  const pdfName = sanitizeDriveName_(`Propuestas de Naves IEM ${sheetName} ${fechaPdf}.pdf`);

  const pdfBlob = DriveApp.getFileById(presId).getAs(MimeType.PDF).setName(pdfName);

  trashFilesByName_(folder, pdfName);
  const pdfFile = folder.createFile(pdfBlob);

  // Escribir link en la hoja
  let outRow = sheet.getLastRow() + 2;
  let outCol = 1;
  if (anchor) {
    const [r, c] = anchor.split(",").map(x => parseInt(x, 10));
    if (isFinite(r) && isFinite(c)) {
      outRow = r + 2;
      outCol = c;
    }
  }

  sheet.getRange(outRow, outCol, 1, 2).clearContent();
  sheet.getRange(outRow, outCol).setValue("PDF:").setFontWeight("bold");
  sheet.getRange(outRow, outCol + 1).setFormula(`=HYPERLINK("${pdfFile.getUrl()}","Abrir PDF")`);

  ss.toast(`PDF generado en carpeta "${sheetName}".`, "Generar PDF", 5);
}
