/***************************************************************
 * CREAR EN SLIDES + GENERAR PDF  (MACHOTE YA ES GOOGLE SLIDES)
 *
 * Inserta:
 * - Slide 1: Portada estilo ANTA (Arial 52/62/36)
 * - Slide 6: Imagen de la tabla (NO editable) ✅ (render real a PNG)
 * - Slide 6: Footnote como TEXTO separado ✅ (ya no se encima)
 * - Slide 7: Imagen del mapa (Drive)
 *
 * ✅ FIX DEFINITIVO:
 * - El mapa se busca SOLO en: OUTPUT_ROOT_FOLDER_ID / MAPS_FOLDER_NAME
 * - Por nombre exacto: MAP_FILE_PREFIX + {Pestaña} + ".png"
 * - Si no existe, TRUENA (no mete mapa equivocado)
 *
 * ✅ TABLA:
 * - Usa anchos reales de columnas desde Sheets
 * - Se DETIENE en la PRIMERA columna vacía del header (no toma columnas después)
 *
 * Requiere en 00_Config.gs:
 *  const OUTPUT_ROOT_FOLDER_ID = "...";
 *  const MAPS_FOLDER_NAME = "Mapas Propuestas";
 *  const MAP_FILE_PREFIX = "MapaCliente_";
 ***************************************************************/

// ===== CONFIG =====
const TEMPLATE_PRESENTATION_ID = "1Hgyml-zwmQLsicHcNW8yg38Bf6ecN9AM7ts8kYzmrNE";

const SLIDE_NUM_COVER = 1;
const SLIDE_NUM_TABLE = 6;
const SLIDE_NUM_MAP = 7;

// Guardar último slides generado
const PROP_LAST_SLIDES_ID = "LAST_PROPOSAL_SLIDES_ID";
const PROP_LAST_LINK_ANCHOR = "LAST_PROPOSAL_LINK_ANCHOR"; // "row,col"
const PROP_LAST_SHEET_NAME = "LAST_PROPOSAL_SHEET_NAME";
const PROP_LAST_FOLDER_ID = "LAST_PROPOSAL_FOLDER_ID";

// ===== PORTADA (tipo ANTA) =====
const COVER_FONT = "Arial";
const COVER_COLOR_DARK = "#2b2b2b";
const COVER_COLOR_BLUE = "#1a73e8";
const COVER_SUB_COLOR = "#111111";

// ===== Layout (fallback) =====
const FALLBACK_SIDE_MARGIN = 70;
const FALLBACK_TOP_MARGIN = 165;
const FALLBACK_BOTTOM_MARGIN = 95;
const TITLE_TO_CONTENT_GAP = 22;
const FOOTER_SAFETY_GAP = 16;

// ===== Tabla PNG =====
const HEADER_FONT_SIZE = 11;
const BODY_FONT_SIZE = 12;
const PNG_EXPORT_SCALE = 3;

/***************************************************************
 * 1) CREAR EN SLIDES
 ***************************************************************/
function CrearEnSlides() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const sheetName = sheet.getName();

  const empresa = promptOrThrow_(ui, "Crear en Slides", "Pregunta 1: Empresa (CLIENTE)?");
  const estadoZona = promptOrThrow_(ui, "Crear en Slides", "Pregunta 2: Estado y Zona (ej: CTT, Estado de México)");
  const rangoSup = promptOrThrow_(ui, "Crear en Slides", "Pregunta 3: Rango de superficie (ej: 8 a 10,000 m2)");
  const fecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");

  // ✅ Carpeta destino = nombre de pestaña activa (dentro del ROOT)
  const sheetFolder = getOrCreateSheetFolderInRoot_(sheetName);

  // Detectar tabla final
  const tableMeta = findFinalTableMeta_(sheet);
  const tableData = readFinalTableForSlides_(sheet, tableMeta);

  if (!tableData.values || tableData.values.length === 0) {
    throw new Error('Encontré "Partida" pero no hay filas de datos (Partida numérica + REF).');
  }

  // Anchos reales desde Sheets (dinámico)
  const colCount = tableData.headers.length;
  const sheetColWidthsPx = getSheetTableWidthsPx_(sheet, tableMeta.colPartida, colCount);

  // Render tabla a PNG (SIN footnote)
  const tableImgBlob = renderTablePngViaTempSlides_(tableData, sheetColWidthsPx);

  // ✅ Mapa STRICT (ROOT/Mapas Propuestas + nombre exacto)
  const mapFile = getMapFileForSheetStrict_(sheetName);
  const mapBlob = mapFile.getBlob();

  // Copiar mapa a carpeta de la pestaña (para que quede todo junto)
  const mapCopyName = mapFile.getName();
  trashFilesByName_(sheetFolder, mapCopyName);
  mapFile.makeCopy(mapCopyName, sheetFolder);

  // Crear Slides en carpeta de pestaña
  const presName = buildPresentationName_(empresa, sheetName, fecha);
  trashFilesByName_(sheetFolder, presName);

  const copyFile = DriveApp.getFileById(TEMPLATE_PRESENTATION_ID).makeCopy(presName, sheetFolder);
  const presId = copyFile.getId();
  const presUrl = `https://docs.google.com/presentation/d/${presId}/edit`;
const folderUrl = `https://drive.google.com/drive/folders/${sheetFolder.getId()}`;

// Escribir ligas SIEMPRE
const outRow = sheet.getLastRow() + 2;
const outCol = 4; // Columna D
ES_writeSlidesLink_(sheet, outRow, outCol, presUrl, folderUrl);


  const props = PropertiesService.getDocumentProperties();
  props.setProperty(PROP_LAST_SLIDES_ID, presId);
  props.setProperty(PROP_LAST_LINK_ANCHOR, `${outRow},${outCol}`);
  props.setProperty(PROP_LAST_SHEET_NAME, sheetName);
  props.setProperty(PROP_LAST_FOLDER_ID, sheetFolder.getId());

  // Editar Slides
  try {
    const pres = SlidesApp.openById(presId);
    const slides = pres.getSlides();

    if (slides.length < SLIDE_NUM_MAP) {
      throw new Error(`Tu machote tiene ${slides.length} slides. Necesito al menos ${SLIDE_NUM_MAP}.`);
    }

    ES_insertCoverInfoANTA_(slides[SLIDE_NUM_COVER - 1], pres, { empresa, estadoZona, rangoSup, fecha });
    esp_insertTableImageOnSlide_(slides[SLIDE_NUM_TABLE - 1], pres, tableImgBlob, "");
    esp_insertMapOnSlide_(slides[SLIDE_NUM_MAP - 1], pres, mapBlob);

    pres.saveAndClose();
    ss.toast(`Listo: carpeta "${sheetName}" con Slides + Mapa.`, "Crear en Slides", 6);
  } catch (e) {
    sheet.getRange(outRow + 1, outCol).setValue("ERROR: " + e.message).setFontColor("#b00020");
    throw e;
  }
}

function ES_writeSlidesLink_(sheet, outRow, outCol, presUrl, folderUrl) {
  // Limpia 2 filas x 2 columnas (D:E y dos renglones)
  sheet.getRange(outRow, outCol, 2, 2).clearContent();

  sheet.getRange(outRow, outCol).setValue("Slides:").setFontWeight("bold");
  sheet.getRange(outRow, outCol + 1).setFormula(`=HYPERLINK("${presUrl}","Abrir Slides")`);

  sheet.getRange(outRow + 1, outCol).setValue("Ubicación de carpeta Drive:").setFontWeight("bold");
  sheet.getRange(outRow + 1, outCol + 1).setFormula(`=HYPERLINK("${folderUrl}","Abrir carpeta")`);
}

/***************************************************************
 * PORTADA (Slide 1) estilo ANTA – 52/62/36
 ***************************************************************/
function ES_insertCoverInfoANTA_(slide, pres, info) {
  const { empresa, estadoZona, rangoSup, fecha } = info;

  esp_removePageElementsByDescription_(slide, "AUTO_COVER_TITLE");
  esp_removePageElementsByDescription_(slide, "AUTO_COVER_CLIENT");
  esp_removePageElementsByDescription_(slide, "AUTO_COVER_SUB");

  esp_cleanupCoverRightSideTextSafe_(slide, pres);

  const pageW = pres.getPageWidth();
  const pageH = pres.getPageHeight();

  const x = pageW * 0.43;
  const w = pageW * 0.52;

  const titleH = 170;
  const gap1 = 35;
  const clientH = 105;
  const gap2 = 35;
  const subH = 170;

  const titleY = pageH * 0.30;

  const titleText = "Propuestas de Naves\nIndustriales para";
  const titleBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x, titleY, w, titleH);
  titleBox.setDescription("AUTO_COVER_TITLE");
  esp_setTextBoxStyle_(titleBox, { text: titleText, font: COVER_FONT, size: 52, bold: true, color: COVER_COLOR_DARK });

  const clientY = titleY + titleH + gap1;
  const clientBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x, clientY, w, clientH);
  clientBox.setDescription("AUTO_COVER_CLIENT");
  esp_setTextBoxStyle_(clientBox, { text: String(empresa || "").toUpperCase(), font: COVER_FONT, size: 62, bold: true, color: COVER_COLOR_BLUE });

  const subY = clientY + clientH + gap2;
  const subText = `${estadoZona}\n${rangoSup}\n${fecha}`;
  const subBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x, subY, w, subH);
  subBox.setDescription("AUTO_COVER_SUB");
  esp_setTextBoxStyle_(subBox, { text: subText, font: COVER_FONT, size: 36, bold: false, color: COVER_SUB_COLOR });
}

function esp_setTextBoxStyle_(shape, cfg) {
  const { text, font, size, bold, color } = cfg;

  const tr = shape.getText();
  tr.setText(text);

  const ts = tr.getTextStyle();
  ts.setFontFamily(font);
  ts.setFontSize(size);
  ts.setBold(!!bold);
  ts.setForegroundColor(color);

  esp_safeSetParagraphAlignment_(tr, SlidesApp.ParagraphAlignment.CENTER);
}

function esp_cleanupCoverRightSideTextSafe_(slide, pres) {
  const pageW = pres.getPageWidth();
  const rightThreshold = pageW * 0.33;

  slide.getPageElements().forEach(pe => {
    try {
      if (pe.getPageElementType() !== SlidesApp.PageElementType.SHAPE) return;
      const shape = pe.asShape();
      if (shape.getLeft() < rightThreshold) return;

      let txt = "";
      try { txt = String(shape.getText().asString() || "").trim(); }
      catch (e) { return; }

      if (!txt) return;
      const t = txt.toLowerCase();
      if (t.includes("propuestas") || t.includes("naves") || t.includes("industriales") || t === "para") {
        shape.remove();
      }
    } catch (e) {}
  });
}

/***************************************************************
 * ✅ FIX: setParagraphAlignment a prueba de balas
 ***************************************************************/
function esp_safeSetParagraphAlignment_(textRange, alignment) {
  try {
    const ps = textRange.getParagraphStyle && textRange.getParagraphStyle();
    if (ps && typeof ps.setParagraphAlignment === "function") {
      ps.setParagraphAlignment(alignment);
      return;
    }
  } catch (e) {}

  try {
    const s = textRange.asString ? String(textRange.asString() || "") : "";
    if (!s) return;
    if (typeof textRange.getRange === "function") {
      const r = textRange.getRange(0, Math.max(0, s.length - 1));
      const ps2 = r.getParagraphStyle && r.getParagraphStyle();
      if (ps2 && typeof ps2.setParagraphAlignment === "function") {
        ps2.setParagraphAlignment(alignment);
        return;
      }
    }
  } catch (e2) {}
}

/***************************************************************
 * Calcular caja para contenido (NO tapar título/footer)
 ***************************************************************/
function esp_computeContentBox_(slide, pres) {
  const pageW = pres.getPageWidth();
  const pageH = pres.getPageHeight();

  let left = FALLBACK_SIDE_MARGIN;
  let right = pageW - FALLBACK_SIDE_MARGIN;
  let top = FALLBACK_TOP_MARGIN;
  let bottom = pageH - FALLBACK_BOTTOM_MARGIN;

  let titleBottom = null;
  let footerTop = null;

  slide.getPageElements().forEach(pe => {
    try {
      if (pe.getPageElementType() !== SlidesApp.PageElementType.SHAPE) return;
      const sh = pe.asShape();

      let txt = "";
      try { txt = String(sh.getText().asString() || "").trim(); }
      catch (e) { return; }
      if (!txt) return;

      const tl = txt.toLowerCase();

      if (tl.includes("propuestas") || tl.includes("ubicaciones") || tl.includes("naves disponibles")) {
        const b = sh.getTop() + sh.getHeight();
        titleBottom = (titleBottom == null) ? b : Math.max(titleBottom, b);
      }

      if (tl.includes("+52") || tl.includes("@industrialestatemexico.com") || tl.includes("industrialestatemexico.com")) {
        footerTop = (footerTop == null) ? sh.getTop() : Math.min(footerTop, sh.getTop());
      }
    } catch (e) {}
  });

  if (titleBottom != null) top = Math.max(top, titleBottom + TITLE_TO_CONTENT_GAP);
  if (footerTop != null) bottom = Math.min(bottom, footerTop - FOOTER_SAFETY_GAP);

  return {
    left,
    top,
    width: Math.max(50, right - left),
    height: Math.max(50, bottom - top)
  };
}

/***************************************************************
 * Slide 6: insertar tabla como IMAGEN + footnote separado
 ***************************************************************/
function esp_insertTableImageOnSlide_(slide, pres, imgBlob, footnoteText) {
  slide.getPageElements().forEach(pe => {
    try { if (pe.getPageElementType() === SlidesApp.PageElementType.TABLE) pe.remove(); } catch (e) {}
  });
  esp_removePageElementsByDescription_(slide, "AUTO_TABLE_IMG");
  esp_removePageElementsByDescription_(slide, "AUTO_TABLE_FOOTNOTE");

  const box = esp_computeContentBox_(slide, pres);

  const footH = 22;
  const gap = 8;

  const imgBox = {
    left: box.left,
    top: box.top,
    width: box.width,
    height: Math.max(50, box.height - footH - gap)
  };

  const img = slide.insertImage(imgBlob);
  img.setDescription("AUTO_TABLE_IMG");

  const iw = img.getWidth();
  const ih = img.getHeight();
  const scale = Math.min(imgBox.width / iw, imgBox.height / ih);

  img.setWidth(iw * scale);
  img.setHeight(ih * scale);
  img.setLeft(imgBox.left + (imgBox.width - img.getWidth()) / 2);
  img.setTop(imgBox.top + (imgBox.height - img.getHeight()) / 2);

  // Footnote separado (no se encima)
  const footY = imgBox.top + imgBox.height + gap;
  const foot = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, box.left, footY, box.width, footH);
  foot.setDescription("AUTO_TABLE_FOOTNOTE");

  foot.getText().setText(footnoteText || "");
  const ts = foot.getText().getTextStyle();
  ts.setFontFamily("Arial").setFontSize(12).setItalic(true).setForegroundColor("#000000");
  esp_safeSetParagraphAlignment_(foot.getText(), SlidesApp.ParagraphAlignment.LEFT);
}

/***************************************************************
 * Slide 7: insertar MAPA como IMAGEN
 ***************************************************************/
function esp_insertMapOnSlide_(slide, pres, mapBlob) {
  esp_removePageElementsByDescription_(slide, "AUTO_MAP");

  const box = esp_computeContentBox_(slide, pres);

  const img = slide.insertImage(mapBlob);
  img.setDescription("AUTO_MAP");

  const iw = img.getWidth();
  const ih = img.getHeight();
  const scale = Math.min(box.width / iw, box.height / ih);

  img.setWidth(iw * scale);
  img.setHeight(ih * scale);
  img.setLeft(box.left + (box.width - img.getWidth()) / 2);
  img.setTop(box.top + (box.height - img.getHeight()) / 2);
}

/***************************************************************
 * TABLA FINAL: detectar último "Partida" que tenga "REF" cerca
 ***************************************************************/
function findFinalTableMeta_(sheet) {
  const matches = sheet.createTextFinder("Partida")
    .matchCase(false)
    .matchEntireCell(false)
    .findAll();

  if (!matches || matches.length === 0) throw new Error('No encontré el encabezado "Partida".');

  let best = null;
  for (const cell of matches) {
    const r = cell.getRow();
    const c = cell.getColumn();
    const w = Math.min(25, sheet.getLastColumn() - c + 1);
    const rowVals = sheet.getRange(r, c, 1, w).getDisplayValues()[0].map(v => String(v || "").trim().toLowerCase());
    const hasRef = rowVals.includes("ref");
    if (hasRef) if (!best || r > best.getRow()) best = cell;
  }
  if (!best) best = matches.reduce((a, b) => (a.getRow() > b.getRow() ? a : b));

  const headerRow = best.getRow();
  const colPartida = best.getColumn();
  const headers = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getDisplayValues()[0].map(v => String(v || "").trim());
  return { headerRow, colPartida, headers };
}

/***************************************************************
 * Leer tabla para Slides (✅ se detiene en primera columna vacía)
 ***************************************************************/
function readFinalTableForSlides_(sheet, meta) {
  const headerRow = meta.headerRow;
  const startCol = meta.colPartida;

  function norm_(s) {
    return String(s || "")
      .trim()
      .toLowerCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/[^a-z0-9]/g, "");
  }

  const maxCol = sheet.getLastColumn();
  const headerSlice = sheet.getRange(headerRow, startCol, 1, maxCol - startCol + 1).getDisplayValues()[0];

  // ✅ Tomar hasta la PRIMERA celda vacía del header
  let end = headerSlice.length;
  for (let i = 0; i < headerSlice.length; i++) {
    if (String(headerSlice[i] || "").trim() === "") { end = i; break; }
  }
  if (end <= 0) throw new Error('Encontré "Partida" pero no pude leer headers a la derecha.');

  const headers = headerSlice.slice(0, end).map(h => String(h || "").trim());

  const idxPartida = headers.findIndex(h => norm_(h) === "partida");
  let idxRef = headers.findIndex(h => norm_(h) === "ref");
  if (idxRef < 0) idxRef = Math.min(1, headers.length - 1);

  const values = [];
  const lastRow = sheet.getLastRow();

  for (let r = headerRow + 1; r <= lastRow; r++) {
    const row = sheet.getRange(r, startCol, 1, headers.length).getDisplayValues()[0];

    const partida = String(row[idxPartida >= 0 ? idxPartida : 0] || "").trim();
    const ref = String(row[idxRef] || "").trim();

    if (!/^\d+$/.test(partida)) break;
    if (!ref) break;

    values.push(row);
  }

  return { headers, values };
}

/***************************************************************
 * Anchos reales de la tabla final en Sheets (px)
 ***************************************************************/
function getSheetTableWidthsPx_(sheet, startCol, count) {
  const out = [];
  for (let i = 0; i < count; i++) out.push(sheet.getColumnWidth(startCol + i));
  return out;
}

/***************************************************************
 * Render tabla a PNG (SIN footnote)
 ***************************************************************/
function renderTablePngViaTempSlides_(tableData, sheetColWidthsPx) {
  const HEADER_BG = "#9fc5e8";

  const tmpPres = SlidesApp.create(`TMP_TableRender_${Date.now()}`);
  const tmpId = tmpPres.getId();
  const slide = tmpPres.getSlides()[0];

  slide.getPageElements().forEach(pe => { try { pe.remove(); } catch (e) {} });

  const pageW = tmpPres.getPageWidth();
  const pageH = tmpPres.getPageHeight();

  const margin = 18;
  const left = margin;
  const top = margin;
  const width = pageW - margin * 2;
  const height = pageH - margin * 2;

  const rows = tableData.values.length + 1;
  const cols = tableData.headers.length;

  // ✅ Firma correcta: insertTable(rows, cols, left, top, width, height)
  const table = slide.insertTable(rows, cols, left, top, width, height);

  function setCellBordersBlack_(cell) {
    const getters = [
      () => cell.getBorderTop(),
      () => cell.getBorderBottom(),
      () => cell.getBorderLeft(),
      () => cell.getBorderRight()
    ];
    getters.forEach(fn => {
      try {
        const b = fn();
        b.getLineFill().setSolidFill("#000000");
        b.setWeight(1);
      } catch (e) {}
    });
  }

  // Encabezados
  for (let c = 0; c < cols; c++) {
    const cell = table.getCell(0, c);
    cell.getText().setText(String(tableData.headers[c] || ""));
    try { cell.getFill().setSolidFill(HEADER_BG); } catch (e) {}

    try {
      const ts = cell.getText().getTextStyle();
      ts.setFontFamily("Arial").setFontSize(HEADER_FONT_SIZE).setBold(true).setForegroundColor("#000000");
      esp_safeSetParagraphAlignment_(cell.getText(), SlidesApp.ParagraphAlignment.CENTER);
    } catch (e) {}

    setCellBordersBlack_(cell);
  }

  // Cuerpo
  for (let r = 0; r < tableData.values.length; r++) {
    for (let c = 0; c < cols; c++) {
      const cell = table.getCell(r + 1, c);
      cell.getText().setText(String(tableData.values[r][c] ?? ""));

      try {
        const ts = cell.getText().getTextStyle();
        ts.setFontFamily("Arial").setFontSize(BODY_FONT_SIZE).setBold(false).setForegroundColor("#000000");
        esp_safeSetParagraphAlignment_(cell.getText(), SlidesApp.ParagraphAlignment.CENTER);
      } catch (e) {}

      setCellBordersBlack_(cell);
    }
  }

  tmpPres.saveAndClose();

  const tableObjectId = getFirstTableObjectId_(tmpId);

  // ===== Anchos de columna vía API (con clamp >= 32pt) =====
  if (sheetColWidthsPx && sheetColWidthsPx.length === cols && tableObjectId) {
    const MIN_COL_PT = 32;

    const sumPx = sheetColWidthsPx.reduce((a, b) => a + (Number(b) || 0), 0) || 1;
    let colWidthsPt = sheetColWidthsPx.map(px => width * ((Number(px) || 0) / sumPx));

    colWidthsPt = clampColumnWidths_(colWidthsPt, width, MIN_COL_PT);

    updateTableColumnWidthsViaSlidesApi_(tmpId, tableObjectId, colWidthsPt);
    Utilities.sleep(600);
  }

  // Export PNG
  const pageId = getFirstSlideObjectIdFromPresentation_(tmpId);
  const url = `https://docs.google.com/presentation/d/${tmpId}/export/png?pageid=${pageId}&scale=${PNG_EXPORT_SCALE}`;
  const resp = UrlFetchApp.fetch(url, {
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });

  if (resp.getResponseCode() !== 200) {
    try { DriveApp.getFileById(tmpId).setTrashed(true); } catch (e) {}
    throw new Error(`No pude exportar PNG. HTTP ${resp.getResponseCode()}: ${resp.getContentText()}`);
  }

  const png = resp.getBlob().setName("tabla_final.png");
  try { DriveApp.getFileById(tmpId).setTrashed(true); } catch (e) {}
  return png;
}

/***************************************************************
 * Clamp columnas: asegura >= minPt y suma = totalWidthPt
 ***************************************************************/
function clampColumnWidths_(widthsPt, totalWidthPt, minPt) {
  const n = widthsPt.length;

  if (minPt * n > totalWidthPt) {
    return widthsPt.map(() => Math.max(minPt, totalWidthPt / n));
  }

  let w = widthsPt.slice();
  let deficit = 0;

  for (let i = 0; i < n; i++) {
    if (w[i] < minPt) {
      deficit += (minPt - w[i]);
      w[i] = minPt;
    }
  }

  if (deficit > 0) {
    const adjustable = [];
    let adjustableRoom = 0;

    for (let i = 0; i < n; i++) {
      const room = w[i] - minPt;
      if (room > 0) {
        adjustable.push(i);
        adjustableRoom += room;
      }
    }

    if (adjustableRoom > 0) {
      adjustable.forEach(i => {
        const room = w[i] - minPt;
        const reduce = deficit * (room / adjustableRoom);
        w[i] = Math.max(minPt, w[i] - reduce);
      });
    }
  }

  let sum = w.reduce((a, b) => a + b, 0);
  const diff = totalWidthPt - sum;

  let maxIdx = 0;
  for (let i = 1; i < n; i++) if (w[i] > w[maxIdx]) maxIdx = i;

  w[maxIdx] = Math.max(minPt, w[maxIdx] + diff);

  return w;
}

/***************************************************************
 * Google Slides API: updateTableColumnProperties (batchUpdate)
 ***************************************************************/
function updateTableColumnWidthsViaSlidesApi_(presentationId, tableObjectId, colWidthsPt) {
  const requests = [];
  for (let i = 0; i < colWidthsPt.length; i++) {
    requests.push({
      updateTableColumnProperties: {
        objectId: tableObjectId,
        columnIndices: [i],
        tableColumnProperties: {
          columnWidth: { magnitude: colWidthsPt[i], unit: "PT" }
        },
        fields: "columnWidth"
      }
    });
  }

  const url = `https://slides.googleapis.com/v1/presentations/${presentationId}:batchUpdate`;
  const resp = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({ requests }),
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });

  if (resp.getResponseCode() !== 200) {
    throw new Error("Slides API batchUpdate (col widths) falló: HTTP " + resp.getResponseCode() + " | " + resp.getContentText());
  }
}

/***************************************************************
 * (AQUÍ ESTÁ EL ÚNICO CAMBIO)
 * Google Slides API: updateTableRowProperties (batchUpdate)
 * ✅ Se usa minRowHeight (NO rowHeight)
 ***************************************************************/
function updateTableRowHeightsViaSlidesApi_(presentationId, tableObjectId, rowHeightsPt) {
  const requests = [];
  for (let i = 0; i < rowHeightsPt.length; i++) {
    requests.push({
      updateTableRowProperties: {
        objectId: tableObjectId,
        rowIndices: [i],
        tableRowProperties: {
          minRowHeight: { magnitude: rowHeightsPt[i], unit: "PT" }
        },
        fields: "minRowHeight"
      }
    });
  }

  const url = `https://slides.googleapis.com/v1/presentations/${presentationId}:batchUpdate`;
  const resp = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({ requests }),
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });

  if (resp.getResponseCode() !== 200) {
    throw new Error("Slides API batchUpdate (row heights) falló: HTTP " + resp.getResponseCode() + " | " + resp.getContentText());
  }
}

/***************************************************************
 * Obtener objectId del primer TABLE (Slides API GET)
 ***************************************************************/
function getFirstTableObjectId_(presentationId) {
  const pres = fetchSlidesPresentation_(presentationId);
  const slide = pres.slides && pres.slides[0];
  if (!slide || !slide.pageElements) return null;

  for (const pe of slide.pageElements) {
    if (pe.table) return pe.objectId;
  }
  return null;
}

/***************************************************************
 * Obtener objectId de la primera slide (pageid) para export/png
 ***************************************************************/
function getFirstSlideObjectIdFromPresentation_(presentationId) {
  const pres = fetchSlidesPresentation_(presentationId);
  if (!pres.slides || !pres.slides[0]) throw new Error("No pude leer la slide temporal.");
  return pres.slides[0].objectId;
}

function fetchSlidesPresentation_(presentationId) {
  const url = `https://slides.googleapis.com/v1/presentations/${presentationId}`;
  const resp = UrlFetchApp.fetch(url, {
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });
  if (resp.getResponseCode() !== 200) {
    throw new Error("Slides API GET falló: HTTP " + resp.getResponseCode() + " | " + resp.getContentText());
  }
  return JSON.parse(resp.getContentText());
}

/***************************************************************
 * ✅ MAPA STRICT: SOLO ROOT / MAPS_FOLDER_NAME
 ***************************************************************/
function getOrCreateMapsFolderInRoot_() {
  const root = DriveApp.getFolderById(OUTPUT_ROOT_FOLDER_ID);
  const it = root.getFoldersByName(MAPS_FOLDER_NAME);
  return it.hasNext() ? it.next() : root.createFolder(MAPS_FOLDER_NAME);
}

function getMapFileForSheetStrict_(sheetName) {
  const folder = getOrCreateMapsFolderInRoot_();
  const expected = `${MAP_FILE_PREFIX}${sheetName}.png`;

  const it = folder.getFilesByName(expected);
  if (it.hasNext()) return it.next();

  throw new Error(
    `No encontré el mapa exacto para la pestaña "${sheetName}".\n` +
    `Debe existir en ROOT/${MAPS_FOLDER_NAME} con el nombre:\n- ${expected}\n\n` +
    `Tip: corre primero "GenerarMapaCliente" en esa pestaña.`
  );
}

/***************************************************************
 * Limpieza por description (AUTO_*)
 ***************************************************************/
function esp_removePageElementsByDescription_(slide, desc) {
  slide.getPageElements().forEach(pe => {
    try { if (pe.getDescription && pe.getDescription() === desc) pe.remove(); } catch (e) {}
  });
}

/***************************************************************
 * ✅ Folder helpers: todo dentro de OUTPUT_ROOT_FOLDER_ID
 ***************************************************************/
function getOrCreateFolderInRoot_(name) {
  const root = DriveApp.getFolderById(OUTPUT_ROOT_FOLDER_ID);
  const it = root.getFoldersByName(name);
  return it.hasNext() ? it.next() : root.createFolder(name);
}

function getOrCreateSheetFolderInRoot_(sheetName) {
  return getOrCreateFolderInRoot_(sanitizeDriveName_(sheetName));
}

function sanitizeDriveName_(name) {
  return String(name || "")
    .replace(/[\/\\:*?"<>|]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function trashFilesByName_(folder, fileName) {
  const it = folder.getFilesByName(fileName);
  while (it.hasNext()) {
    try { it.next().setTrashed(true); } catch (e) {}
  }
}

function buildPresentationName_(empresa, sheetName, fecha) {
  const safeEmpresa = sanitizeDriveName_(empresa);
  const safeSheet = sanitizeDriveName_(sheetName);
  return `Propuesta - ${safeEmpresa} - ${safeSheet} - ${fecha}`;
}

/***************************************************************
 * UI helper
 ***************************************************************/
function promptOrThrow_(ui, title, message) {
  const res = ui.prompt(title, message, ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) throw new Error("Operación cancelada.");
  const val = String(res.getResponseText() || "").trim();
  if (!val) throw new Error("Campo vacío. Cancelado.");
  return val;
}
