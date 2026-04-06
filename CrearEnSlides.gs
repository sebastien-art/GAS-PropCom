/***************************************************************
 * 1) CONFIGURACIÓN GLOBAL
 ***************************************************************/
const TEMPLATE_PRESENTATION_ID = "1Hgyml-zwmQLsicHcNW8yg38Bf6ecN9AM7ts8kYzmrNE";

const SLIDE_NUM_COVER = 1;
const SLIDE_NUM_TABLE = 6;
const SLIDE_NUM_MAP = 7;

// Estilos de Portada
const COVER_FONT = "Arial";
const COVER_COLOR_DARK = "#2b2b2b";
const COVER_COLOR_BLUE = "#1a73e8";
const COVER_SUB_COLOR = "#111111";

// Layout y Márgenes
const FALLBACK_SIDE_MARGIN = 70;
const FALLBACK_TOP_MARGIN = 165;
const FALLBACK_BOTTOM_MARGIN = 95;

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

    const sheetFolder = getOrCreateSheetFolderInRoot_(sheetName);
    const tableMeta = findFinalTableMeta_(sheet);
    const tableData = readFinalTableForSlides_(sheet, tableMeta);

    if (!tableData.values || tableData.values.length === 0) throw new Error('No hay datos en la tabla.');

    const mapFile = getMapFileForSheetStrict_(sheetName);
    const mapBlob = mapFile.getBlob().setContentType("image/png");

    // Generar Imagen de Tabla con Medidas Espejo
    const tableImgBlob = renderTablePngViaTempSlides_(tableData, tableMeta);

    const presName = buildPresentationName_(empresa, sheetName, fecha);
    trashFilesByName_(sheetFolder, presName);
    const copyFile = DriveApp.getFileById(TEMPLATE_PRESENTATION_ID).makeCopy(presName, sheetFolder);
    const presId = copyFile.getId();
    
    const pres = SlidesApp.openById(presId);
    const slides = pres.getSlides();

    ES_insertCoverInfoANTA_(slides[SLIDE_NUM_COVER - 1], pres, { empresa, estadoZona, rangoSup, fecha });
    esp_insertTableImageOnSlide_(slides[SLIDE_NUM_TABLE - 1], pres, tableImgBlob, "");
    esp_insertMapOnSlide_(slides[SLIDE_NUM_MAP - 1], pres, mapBlob);

    pres.saveAndClose();
    ES_writeSlidesLink_(sheet, sheet.getLastRow() + 2, 4, copyFile.getUrl(), sheetFolder.getUrl());
    ss.toast("¡Presentación generada con éxito! ✅", "Finalizado");

  } catch (e) {
    ui.alert("Error: " + e.toString());
  }
}

/***************************************************************
 * 3) MOTOR DE RENDERIZADO ESPEJO (ANCHO Y ALTO FIEL)
 ***************************************************************/
/***************************************************************
 * MOTOR DE RENDERIZADO: ESTRATEGIA DE GRÁFICO (CAPTURA RÁPIDA)
 ***************************************************************/
/***************************************************************
 * MOTOR DE RENDERIZADO: CAPTURA POR VISUALIZACIÓN (STRICT)
 ***************************************************************/
/***************************************************************
 * MOTOR DE RENDERIZADO: CAPTURA DE RANGO (PDF SNAPSHOT)
 ***************************************************************/
function renderTablePngViaTempSlides_(tableData, tableMeta) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const ssId = ss.getId();
  const sheetId = sheet.getSheetId();

  // 1. Definir coordenadas del rango (Header + Datos)
  const r1 = tableMeta.headerRow;
  const c1 = tableMeta.colPartida;
  const r2 = tableMeta.headerRow + tableData.values.length;
  const c2 = tableMeta.colPartida + tableData.headers.length - 1;

  // 2. Construir URL de exportación "Snapshot"
  // Esta URL configura el PDF para que no tenga márgenes y se ajuste al rango
  const url = "https://docs.google.com/spreadsheets/d/" + ssId + "/export?" +
    "format=pdf&" +
    "size=letter&" +
    "portrait=false&" +
    "fitw=true&" +              // Ajustar ancho
    "source=labnol&" +
    "sheetid=" + sheetId + "&" +
    "range=R" + r1 + "C" + c1 + ":R" + r2 + "C" + c2 + "&" + // Coordenadas R1C1
    "gridlines=false&" +        // Sin cuadrícula de fondo
    "printtitle=false&" +
    "scale=4";                  // Alta resolución

  // 3. Obtener el Blob usando el Token de autorización
  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, {
    headers: { 'Authorization': 'Bearer ' + token },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    throw new Error("Error al capturar imagen del rango: " + response.getContentText());
  }

  // 4. Convertir el PDF resultante en una imagen PNG 
  // (Google Slides acepta Blobs de PDF directamente para insertarlos como imagen)
  const pdfBlob = response.getBlob().setName("Tabla_Fiel_" + new Date().getTime() + ".pdf");

  return pdfBlob;
}

/***************************************************************
 * 4) FUNCIONES DE EDICIÓN
 ***************************************************************/
function ES_insertCoverInfoANTA_(slide, pres, info) {
  esp_removePageElementsByDescription_(slide, "AUTO_COVER_TITLE");
  esp_removePageElementsByDescription_(slide, "AUTO_COVER_CLIENT");
  esp_removePageElementsByDescription_(slide, "AUTO_COVER_SUB");
  esp_cleanupCoverRightSideTextSafe_(slide, pres);

  const pageW = pres.getPageWidth();
  const x = pageW * 0.43, w = pageW * 0.52, titleY = pres.getPageHeight() * 0.30;

  esp_setTextBoxStyle_(slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x, titleY, w, 170).setDescription("AUTO_COVER_TITLE"), 
    { text: "Propuestas de Naves\nIndustriales para", font: COVER_FONT, size: 52, bold: true, color: COVER_COLOR_DARK });
  
  esp_setTextBoxStyle_(slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x, titleY + 205, w, 105).setDescription("AUTO_COVER_CLIENT"), 
    { text: String(info.empresa).toUpperCase(), font: COVER_FONT, size: 62, bold: true, color: COVER_COLOR_BLUE });

  esp_setTextBoxStyle_(slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x, titleY + 345, w, 170).setDescription("AUTO_COVER_SUB"), 
    { text: `${info.estadoZona}\n${info.rangoSup}\n${info.fecha}`, font: COVER_FONT, size: 36, bold: false, color: COVER_SUB_COLOR });
}

function esp_insertMapOnSlide_(slide, pres, mapBlob) {
  esp_removePageElementsByDescription_(slide, "AUTO_MAP");
  const box = esp_computeContentBox_(slide, pres);
  const img = slide.insertImage(mapBlob).setDescription("AUTO_MAP");
  const ratio = Math.min(box.width / img.getWidth(), box.height / img.getHeight());
  img.setWidth(img.getWidth() * ratio).setHeight(img.getHeight() * ratio);
  img.setLeft(box.left + (box.width - img.getWidth()) / 2).setTop(box.top + (box.height - img.getHeight()) / 2);
}

function esp_insertTableImageOnSlide_(slide, pres, imgBlob, footnote) {
  esp_removePageElementsByDescription_(slide, "AUTO_TABLE_IMG");
  const box = esp_computeContentBox_(slide, pres);
  const img = slide.insertImage(imgBlob).setDescription("AUTO_TABLE_IMG");
  const ratio = Math.min(box.width / img.getWidth(), box.height / img.getHeight());
  img.setWidth(img.getWidth() * ratio).setHeight(img.getHeight() * ratio);
  img.setLeft(box.left + (box.width - img.getWidth()) / 2).setTop(box.top + (box.height - img.getHeight()) / 2);
}

/***************************************************************
 * 5) UTILIDADES
 ***************************************************************/
function esp_computeContentBox_(slide, pres) {
  const pageW = pres.getPageWidth(), pageH = pres.getPageHeight();
  let top = FALLBACK_TOP_MARGIN, bottom = pageH - FALLBACK_BOTTOM_MARGIN;
  let titleB = null, footerT = null;

  slide.getPageElements().forEach(pe => {
    if (pe.getPageElementType() !== SlidesApp.PageElementType.SHAPE) return;
    try {
      const txt = pe.asShape().getText().asString().toLowerCase();
      if (txt.includes("propuestas") || txt.includes("naves")) titleB = Math.max(titleB || 0, pe.getTop() + pe.getHeight());
      if (txt.includes("+52") || txt.includes("@industrial")) footerT = Math.min(footerT || pageH, pe.getTop());
    } catch (e) {}
  });
  if (titleB) top = titleB + 20;
  if (footerT) bottom = footerT - 15;
  return { left: FALLBACK_SIDE_MARGIN, top: top, width: pageW - (FALLBACK_SIDE_MARGIN * 2), height: Math.max(50, bottom - top) };
}

function findFinalTableMeta_(sheet) {
  const matches = sheet.createTextFinder("Partida").matchCase(false).findAll();
  if (!matches.length) throw new Error('No encontré "Partida".');
  let best = matches[matches.length - 1]; 
  return { headerRow: best.getRow(), colPartida: best.getColumn() };
}

function readFinalTableForSlides_(sheet, meta) {
  const headers = sheet.getRange(meta.headerRow, meta.colPartida, 1, 15).getDisplayValues()[0];
  let end = headers.findIndex(h => h.trim() === "");
  const finalHeaders = headers.slice(0, end > 0 ? end : headers.length);
  const values = [];
  for (let r = meta.headerRow + 1; r <= sheet.getLastRow(); r++) {
    const row = sheet.getRange(r, meta.colPartida, 1, finalHeaders.length).getDisplayValues()[0];
    if (!/^\d+$/.test(row[0]) || !row[1]) break;
    values.push(row);
  }
  return { headers: finalHeaders, values };
}

function getOrCreateFolderInRoot_(name) {
  const folders = DriveApp.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(name);
}
function getOrCreateSheetFolderInRoot_(name) { return getOrCreateFolderInRoot_(name.replace(/[\/\\:*?"<>|]/g, " ")); }
function trashFilesByName_(folder, name) { const it = folder.getFilesByName(name); while (it.hasNext()) it.next().setTrashed(true); }
function buildPresentationName_(e, s, f) { return `Propuesta - ${e} - ${s} - ${f}`; }
function promptOrThrow_(ui, t, m) {
  const r = ui.prompt(t, m, ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) throw new Error("Cancelado");
  return r.getResponseText().trim();
}
function esp_removePageElementsByDescription_(s, d) { s.getPageElements().forEach(pe => { try { if (pe.getDescription() === d) pe.remove(); } catch(e){} }); }
function esp_setTextBoxStyle_(s, c) {
  const tr = s.getText(); tr.setText(c.text);
  tr.getTextStyle().setFontFamily(c.font).setFontSize(c.size).setBold(!!c.bold).setForegroundColor(c.color);
  try { tr.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER); } catch(e){}
}
function esp_cleanupCoverRightSideTextSafe_(s, p) {
  s.getPageElements().forEach(pe => {
    if (pe.getLeft() > p.getPageWidth() * 0.33) {
      try { if (pe.asShape().getText().asString().toLowerCase().includes("propuestas")) pe.remove(); } catch(e){}
    }
  });
}
function getMapFileForSheetStrict_(name) {
  const it = DriveApp.getFilesByName(`Mapa_${name}.png`);
  if (it.hasNext()) return it.next();
  throw new Error(`Falta "Mapa_${name}.png".`);
}
function ES_writeSlidesLink_(sheet, row, col, url, folder) {
  sheet.getRange(row, col).setValue("Slides:").setFontWeight("bold");
  sheet.getRange(row, col + 1).setFormula(`=HYPERLINK("${url}","Abrir Slides")`);
  sheet.getRange(row + 1, col).setValue("Carpeta:").setFontWeight("bold");
  sheet.getRange(row + 1, col + 1).setFormula(`=HYPERLINK("${folder}","Abrir carpeta")`);
}