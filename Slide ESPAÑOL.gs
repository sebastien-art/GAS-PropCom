/***************************************************************
 * 1) CONFIGURACIÓN GLOBAL
 ***************************************************************/
const TEMPLATE_PRESENTATION_ID = "1Hgyml-zwmQLsicHcNW8yg38Bf6ecN9AM7ts8kYzmrNE";

const SLIDE_NUM_COVER = 1;
const SLIDE_NUM_TABLE = 6;
const SLIDE_NUM_MAP = 7;

const COVER_FONT = "Arial";
const COVER_COLOR_DARK = "#2b2b2b";
const COVER_COLOR_BLUE = "#1a73e8";
const COVER_SUB_COLOR = "#111111";

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

    const tableImgBlob = renderTablePngViaTempSlides_(tableData);

    const presName = buildPresentationName_(empresa, sheetName, fecha);
    trashFilesByName_(sheetFolder, presName);
    
    const copyFile = DriveApp.getFileById(TEMPLATE_PRESENTATION_ID).makeCopy(presName, sheetFolder);
    const presId = copyFile.getId();
    const pres = SlidesApp.openById(presId);
    const slides = pres.getSlides();

    ES_insertCoverInfo_FINAL(slides[SLIDE_NUM_COVER - 1], pres, { empresa, estadoZona, rangoSup, fecha });
    ES_executeSafeInsert(slides[SLIDE_NUM_TABLE - 1], pres, tableImgBlob, "AUTO_TABLE_IMG");
    ES_executeSafeInsert(slides[SLIDE_NUM_MAP - 1], pres, mapBlob, "AUTO_MAP_IMG");

    pres.saveAndClose();

    PropertiesService.getDocumentProperties().setProperty("LAST_SLIDES_ID", presId);
    
    ES_writeSlidesLink_(sheet, sheet.getLastRow() + 2, 4, copyFile.getUrl(), sheetFolder.getUrl());
    ss.toast("¡Presentación generada! ✅", "Éxito");

  } catch (e) {
    ui.alert("Error Crítico: " + e.toString());
  }
}

/***************************************************************
 * 3) LÓGICA DE INSERCIÓN SEGURA (COPIADA DE CHI/ENG)
 ***************************************************************/
function ES_executeSafeInsert(slide, pres, blob, desc) {
  try {
    slide.getPageElements().forEach(pe => {
      try { 
        if (pe.getDescription() === desc) pe.remove(); 
      } catch(e) {}
    });

    const box = { 
      left: 70, 
      top: 165, 
      width: pres.getPageWidth() - 140, 
      height: pres.getPageHeight() - 260 
    };
    
    const img = slide.insertImage(blob).setDescription(desc);
    const ratio = Math.min(box.width / img.getWidth(), box.height / img.getHeight());
    
    img.setWidth(img.getWidth() * ratio).setHeight(img.getHeight() * ratio);
    img.setLeft(box.left + (box.width - img.getWidth()) / 2).setTop(box.top + (box.height - img.getHeight()) / 2);
    
  } catch (e) {
    Logger.log("Error en inserción: " + e.toString());
  }
}

/***************************************************************
 * 4) PORTADA (LÓGICA CHI/ENG)
 ***************************************************************/
function ES_insertCoverInfo_FINAL(slide, pres, info) {
  ["AUTO_COVER_CLIENT", "AUTO_COVER_SUB"].forEach(d => {
    slide.getPageElements().forEach(pe => { 
      try { if(pe.getDescription() === d) pe.remove(); } catch(e){} 
    });
  });

const x = pres.getPageWidth() * 0.43, w = pres.getPageWidth() * 0.52, y = pres.getPageHeight() * 0.30;

ES_setTxt(slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x, y + 255, w, 105).setDescription("AUTO_COVER_CLIENT"), String(info.empresa).toUpperCase(), 62, true, "#1a73e8");

ES_setTxt(slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x, y + 345, w, 120).setDescription("AUTO_COVER_SUB"), `${info.estadoZona}\n${info.rangoSup}\n${info.fecha}`, 36, false, "#111111");
}

function ES_setTxt(s, t, sz, b, c) {
  const tr = s.getText(); 
  tr.setText(t);
  tr.getTextStyle().setFontFamily("Arial").setFontSize(sz).setBold(b).setForegroundColor(c);
  tr.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
}

/***************************************************************
 * 5) RENDER TABLA (ESTÁNDAR)
 ***************************************************************/
function renderTablePngViaTempSlides_(tableData) {
  const tempPres = SlidesApp.create("__TEMP_TABLE__");
  const tempId = tempPres.getId();

  try {
    const slide = tempPres.getSlides()[0];
    const PW = tempPres.getPageWidth();
    const PH = tempPres.getPageHeight();
    const allRows = [tableData.headers, ...tableData.values];
    const cols = tableData.headers.length;
    const rowCount = allRows.length;
    const colW = (PW - 60) / cols;
    const rowH = Math.min(22, (PH - 120) / rowCount);
    const startX = 30;
    const startY = 40;

    allRows.forEach((row, rIdx) => {
      row.forEach((cell, cIdx) => {
        try {
          const x = startX + cIdx * colW;
          const y = startY + rIdx * rowH;
          const shape = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, x, y, colW, rowH);
          const isHeader = rIdx === 0;
          shape.getFill().setSolidFill(isHeader ? "#1a3a5c" : (rIdx % 2 === 0 ? "#f0f4f8" : "#ffffff"));
          shape.getBorder().getLineFill().setSolidFill("#cccccc");
          shape.getBorder().setWeight(0.5);
          const tf = shape.getText();
          tf.setText(String(cell == null ? "" : cell));
          tf.getTextStyle().setFontFamily("Arial").setFontSize(isHeader ? 7.5 : 6.5).setBold(isHeader).setForegroundColor(isHeader ? "#ffffff" : "#222222");
          tf.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
        } catch(cellErr) {
          Logger.log("Error en celda [" + rIdx + "," + cIdx + "]: " + cellErr.message);
        }
      });
    });

    tempPres.saveAndClose();

    const thumbnailUrl = "https://slides.googleapis.com/v1/presentations/" + tempId + "/pages/" + SlidesApp.openById(tempId).getSlides()[0].getObjectId() + "/thumbnail?thumbnailProperties.mimeType=PNG&thumbnailProperties.thumbnailSize=LARGE";
    const token = ScriptApp.getOAuthToken();
    const resp = UrlFetchApp.fetch(thumbnailUrl, {
      headers: { Authorization: "Bearer " + token },
      muteHttpExceptions: true
    });

    if (resp.getResponseCode() !== 200) {
      throw new Error("Slides thumbnail API falló: " + resp.getContentText().substring(0, 300));
    }

    const json = JSON.parse(resp.getContentText());
    const imgResp = UrlFetchApp.fetch(json.contentUrl);
    return imgResp.getBlob()
      .setName("Tabla_" + new Date().getTime() + ".png")
      .setContentType("image/png");

  } finally {
    DriveApp.getFileById(tempId).setTrashed(true);
  }
}

/***************************************************************
 * 6) HELPERS RESTANTES
 ***************************************************************/
function readFinalTableForSlides_(sheet, meta) {
  const colsDeseadas = ["Operación", "REF", "Estado", "Zona Principal", "Sub Zona", "Desarrollador", "Intermediario", "Parque", "Nave", "M2 de construcción", "Asking price /m2", "Mantenimiento / m2", "Disponibilidad"];
  const headerValues = sheet.getRange(meta.headerRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  const mapaIndices = colsDeseadas.map(nombre => headerValues.indexOf(nombre));
  const values = [];
  const data = sheet.getDataRange().getDisplayValues();
  for (let r = meta.headerRow; r < data.length; r++) {
    const row = data[r];
    if ((row[headerValues.indexOf("REF")] || "") === "" && row[headerValues.indexOf("Operación")] === "") break;
    if (row.join("").includes("/") && row.filter(c => c !== "").length < 3) continue;
    values.push(mapaIndices.map(idx => (idx !== -1 ? row[idx] : "N/A")));
  }
  return { headers: colsDeseadas, values: values };
}

function findFinalTableMeta_(sheet) {
  const finder = sheet.createTextFinder("Operación").matchCase(false).findNext();
  if (!finder) throw new Error('No se encontró la columna "Operación"');
  return { headerRow: finder.getRow(), colPartida: finder.getColumn() };
}

function getMapFileForSheetStrict_(name) {
  const it = DriveApp.getFolderById("1hFZd3Re2q5ScMQQ3h5jRv8k1oW70Fmrv").getFilesByName("MapaCliente_" + name + ".png");
  if (it.hasNext()) return it.next();
  throw new Error("Falta mapa para " + name);
}

function getOrCreateSheetFolderInRoot_(name) { 
  const n = name.replace(/[\/\\:*?"<>|]/g, " ");
  const it = DriveApp.getFoldersByName(n);
  return it.hasNext() ? it.next() : DriveApp.createFolder(n);
}

function trashFilesByName_(folder, name) { const it = folder.getFilesByName(name); while (it.hasNext()) it.next().setTrashed(true); }
function buildPresentationName_(e, s, f) { return `Propuesta - ${e} - ${s} - ${f}`; }
function promptOrThrow_(ui, t, m) { const r = ui.prompt(t, m, ui.ButtonSet.OK_CANCEL); if (r.getSelectedButton() !== ui.Button.OK) throw new Error("Cancelado"); return r.getResponseText().trim(); }

function ES_writeSlidesLink_(sheet, row, col, url, folder) {
  sheet.getRange(row, col, 2, 1).setValues([["Slides:"],["Carpeta:"]]).setFontWeight("bold");
  sheet.getRange(row, col + 1).setFormula(`=HYPERLINK("${url}","Abrir Slides")`);
  sheet.getRange(row + 1, col + 1).setFormula(`=HYPERLINK("${folder}","Abrir carpeta")`);
}