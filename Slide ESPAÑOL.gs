/***************************************************************
 * 1) CONFIGURACIÓN GLOBAL
 ***************************************************************/
const TEMPLATE_PRESENTATION_ID = "1Hgyml-zwmQLsicHcNW8yg38Bf6ecN9AM7ts8kYzmrNE";

const SLIDE_NUM_COVER = 1;

/***************************************************************
 * 2) FUNCIÓN PRINCIPAL
 ***************************************************************/
function CrearEnSlides() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const sheetName = sheet.getName();

  try {
    const empresa = promptOrThrow_(ui, "Crear en Slides", "Pregunta 1: Empresa (CLIENTE)?");
    const estadoZona = promptOrThrow_(ui, "Crear en Slides", "Pregunta 2: Estado y Zona?");
    const rangoSup = promptOrThrow_(ui, "Crear en Slides", "Pregunta 3: Rango de superficie?");
    const fecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");

    // 1. Crear o buscar carpeta para guardar la presentación
    const sheetFolder = getOrCreateSheetFolderInRoot_(sheetName);

    // 2. Definir nombre y limpiar archivos anteriores con el mismo nombre
    const presName = buildPresentationName_(empresa, sheetName, fecha);
    trashFilesByName_(sheetFolder, presName);

    // 3. Copiar la plantilla e instanciar Slides
    const copyFile = DriveApp.getFileById(TEMPLATE_PRESENTATION_ID).makeCopy(presName, sheetFolder);
    const presId = copyFile.getId();
    const pres = SlidesApp.openById(presId);
    const slides = pres.getSlides();

    // 4. Llenar solo los datos de la Portada
    ES_insertCoverInfo_FINAL(slides[SLIDE_NUM_COVER - 1], pres, { empresa, estadoZona, rangoSup, fecha });

    pres.saveAndClose();

    // 5. Guardar ID y escribir links en la hoja
    PropertiesService.getDocumentProperties().setProperty("LAST_SLIDES_ID", presId);
    ES_writeSlidesLink_(sheet, sheet.getLastRow() + 2, 4, copyFile.getUrl(), sheetFolder.getUrl());

    ss.toast("¡Presentación generada! ✅", "Éxito");

  } catch (e) {
    ui.alert("Error Crítico: " + e.toString());
  }
}

/***************************************************************
 * 3) LÓGICA DE LA PORTADA
 ***************************************************************/
function ES_insertCoverInfo_FINAL(slide, pres, info) {
  ["AUTO_COVER_CLIENT", "AUTO_COVER_SUB"].forEach(d => {
    slide.getPageElements().forEach(pe => {
      try { if(pe.getDescription() === d) pe.remove(); } catch(e) {}
    });
  });

  const x = pres.getPageWidth() * 0.43;
  const w = pres.getPageWidth() * 0.52;
  const y = pres.getPageHeight() * 0.30;

  ES_setTxt(
    slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x, y + 255, w, 105).setDescription("AUTO_COVER_CLIENT"),
    String(info.empresa).toUpperCase(), 
    62, 
    true, 
    "#1a73e8"
  );

  ES_setTxt(
    slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x, y + 345, w, 120).setDescription("AUTO_COVER_SUB"),
    `${info.estadoZona}\n${info.rangoSup}\n${info.fecha}`, 
    36, 
    false, 
    "#111111"
  );
}

function ES_setTxt(s, t, sz, b, c) {
  const tr = s.getText();
  tr.setText(t);
  tr.getTextStyle().setFontFamily("Arial").setFontSize(sz).setBold(b).setForegroundColor(c);
  tr.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
}

/***************************************************************
 * 4) HELPERS
 ***************************************************************/
function getOrCreateSheetFolderInRoot_(name) {
  const n = name.replace(/[\/\\:*?"<>|]/g, " ");
  const it = DriveApp.getFoldersByName(n);
  return it.hasNext() ? it.next() : DriveApp.createFolder(n);
}

function trashFilesByName_(folder, name) { 
  const it = folder.getFilesByName(name); 
  while (it.hasNext()) it.next().setTrashed(true); 
}

function buildPresentationName_(e, s, f) { 
  return `Propuesta - ${e} - ${s} - ${f}`; 
}

function promptOrThrow_(ui, t, m) { 
  const r = ui.prompt(t, m, ui.ButtonSet.OK_CANCEL); 
  if (r.getSelectedButton() !== ui.Button.OK) throw new Error("Cancelado"); 
  return r.getResponseText().trim(); 
}

function ES_writeSlidesLink_(sheet, row, col, url, folder) {
  sheet.getRange(row, col, 2, 1).setValues([["Slides:"],["Carpeta:"]]).setFontWeight("bold");
  sheet.getRange(row, col + 1).setFormula(`=HYPERLINK("${url}","Abrir Slides")`);
  sheet.getRange(row + 1, col + 1).setFormula(`=HYPERLINK("${folder}","Abrir carpeta")`);
}