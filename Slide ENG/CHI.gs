/***************************************************************
 * SLIDES_ENG_CHI  (MACHOTE = GOOGLE SLIDES TEMPLATE PROVIDED)
 *
 * Uses template:
 *  https://docs.google.com/presentation/d/1xA4FVsxwUxrH9DPhwkfR8vYB5lD4lPYyuuSFi8H3UYY/edit
 *
 * Inserts:
 * - Slide 1: Cover (Arial 52/62/36)
 * - Slide 6: Table image (PNG render, not editable) + footnote as separate text
 * - Slide 7: Map image (Drive)
 *
 * STRICT MAP:
 * - Looks ONLY in: OUTPUT_ROOT_FOLDER_ID / MAPS_FOLDER_NAME
 * - Exact name: MAP_FILE_PREFIX + {SheetName} + ".png"
 * - If missing: throws error (no wrong map)
 *
 * TABLE:
 * - Reads widths from Sheets
 * - Stops at the first empty header cell (doesn't include columns after blank)
 *
 * Requires (in your 00_Config.gs):
 *  const OUTPUT_ROOT_FOLDER_ID = "...";
 *  const MAPS_FOLDER_NAME = "Mapas Propuestas";
 *  const MAP_FILE_PREFIX = "MapaCliente_";
 ***************************************************************/

// ===== CONFIG (wrapped to avoid duplicate const collisions across files) =====
const SLIDES_ENG_CHI_CFG = {
  TEMPLATE_PRESENTATION_ID: "1xA4FVsxwUxrH9DPhwkfR8vYB5lD4lPYyuuSFi8H3UYY",

  SLIDE_NUM_COVER: 1,
  SLIDE_NUM_TABLE: 6,
  SLIDE_NUM_MAP: 7,

  // Save last generated slides
  PROP_LAST_SLIDES_ID: "LAST_PROPOSAL_SLIDES_ID",
  PROP_LAST_LINK_ANCHOR: "LAST_PROPOSAL_LINK_ANCHOR", // "row,col"
  PROP_LAST_SHEET_NAME: "LAST_PROPOSAL_SHEET_NAME",
  PROP_LAST_FOLDER_ID: "LAST_PROPOSAL_FOLDER_ID",

  // Cover
  COVER_FONT: "Arial",
  COVER_COLOR_DARK: "#2b2b2b",
  COVER_COLOR_BLUE: "#1a73e8",
  COVER_SUB_COLOR: "#111111",

  // Layout (fallback)
  FALLBACK_SIDE_MARGIN: 70,
  FALLBACK_TOP_MARGIN: 165,
  FALLBACK_BOTTOM_MARGIN: 95,
  TITLE_TO_CONTENT_GAP: 22,
  FOOTER_SAFETY_GAP: 16,

  // Table PNG
  HEADER_FONT_SIZE: 11,
  BODY_FONT_SIZE: 12,
  PNG_EXPORT_SCALE: 3,

  };

/***************************************************************
 * 1) CREATE SLIDES (ENG/CHI)
 * (NOTE: function names in Apps Script can't contain "/"
 * so it's Slides_ENG_CHI)
 ***************************************************************/
function Slides_ENG_CHI() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const sheetName = sheet.getName();

  const empresa = promptOrThrow_(
    ui,
    "Create Slides / 生成幻灯片",
    "Q1: Nombre de Empresa"
  );
  const estadoZona = promptOrThrow_(
    ui,
    "Create Slides / 生成幻灯片",
    "Q2: Zona y Estado"
  );
  const rangoSup = promptOrThrow_(
    ui,
    "Create Slides / 生成幻灯片",
    "Q3: Superficie"
  );

  const fecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");

  // Destination folder = active sheet name (inside ROOT)
  const sheetFolder = getOrCreateSheetFolderInRoot_(sheetName);

  // Detect final table (supports Spanish "Partida" OR ENG/CHI "Item")
  const tableMeta = findFinalTableMeta_(sheet);
  const tableData = readFinalTableForSlides_(sheet, tableMeta);

  if (!tableData.values || tableData.values.length === 0) {
    throw new Error('Found the table header but no data rows (numeric Item/Partida + REF).');
  }

  // Real column widths from Sheets (dynamic)
  const colCount = tableData.headers.length;
  const sheetColWidthsPx = getSheetTableWidthsPx_(sheet, tableMeta.colPartida, colCount);

  // Render table to PNG (NO footnote inside the image)
  const tableImgBlob = renderTablePngViaTempSlides_(tableData, sheetColWidthsPx);

  // STRICT map (ROOT/Mapas Propuestas + exact name)
  const mapFile = getMapFileForSheetStrict_(sheetName);
  const mapBlob = mapFile.getBlob();
  Logger.log("MAP name=%s type=%s bytes=%s",
  mapFile.getName(),
  mapBlob.getContentType(),
  mapBlob.getBytes().length
);

  // Copy map into the sheet folder (keep everything together)
  const mapCopyName = mapFile.getName();
  trashFilesByName_(sheetFolder, mapCopyName);
  mapFile.makeCopy(mapCopyName, sheetFolder);

  // Create Slides in sheet folder
  const presName = buildPresentationName_(empresa, sheetName, fecha);
  trashFilesByName_(sheetFolder, presName);

  const copyFile = DriveApp.getFileById(SLIDES_ENG_CHI_CFG.TEMPLATE_PRESENTATION_ID).makeCopy(presName, sheetFolder);
  const presId = copyFile.getId();
  const presUrl = `https://docs.google.com/presentation/d/${presId}/edit`;
  const folderUrl = `https://drive.google.com/drive/folders/${sheetFolder.getId()}`;

  // Always write links
  const outRow = sheet.getLastRow() + 2;
  const outCol = 4; // Column D
  writeSlidesLink_(sheet, outRow, outCol, presUrl, folderUrl);

  const props = PropertiesService.getDocumentProperties();
  props.setProperty(SLIDES_ENG_CHI_CFG.PROP_LAST_SLIDES_ID, presId);
  props.setProperty(SLIDES_ENG_CHI_CFG.PROP_LAST_LINK_ANCHOR, `${outRow},${outCol}`);
  props.setProperty(SLIDES_ENG_CHI_CFG.PROP_LAST_SHEET_NAME, sheetName);
  props.setProperty(SLIDES_ENG_CHI_CFG.PROP_LAST_FOLDER_ID, sheetFolder.getId());

  // Edit Slides
  try {
    const pres = SlidesApp.openById(presId);
    const slides = pres.getSlides();

    if (slides.length < SLIDES_ENG_CHI_CFG.SLIDE_NUM_MAP) {
      throw new Error(`Template has ${slides.length} slides. Need at least ${SLIDES_ENG_CHI_CFG.SLIDE_NUM_MAP}.`);
    }

    insertCoverInfoANTA_(slides[SLIDES_ENG_CHI_CFG.SLIDE_NUM_COVER - 1], pres, { empresa, estadoZona, rangoSup, fecha });
    insertTableImageOnSlide_(slides[SLIDES_ENG_CHI_CFG.SLIDE_NUM_TABLE - 1], pres, tableImgBlob, SLIDES_ENG_CHI_CFG.TABLE_FOOTNOTE_TEXT);
    insertMapOnSlide_(slides[SLIDES_ENG_CHI_CFG.SLIDE_NUM_MAP - 1], pres, mapBlob);

    pres.saveAndClose();
    ss.toast(`Done: folder "${sheetName}" with Slides + Map. / 完成：已生成幻灯片与地图`, "Slides_ENG_CHI", 6);
  } catch (e) {
    sheet.getRange(outRow + 1, outCol).setValue("ERROR: " + e.message).setFontColor("#b00020");
    throw e;
  }
}

function writeSlidesLink_(sheet, outRow, outCol, presUrl, folderUrl) {
  // Clear 2 rows x 2 cols (D:E, two rows)
  sheet.getRange(outRow, outCol, 2, 2).clearContent();

  sheet.getRange(outRow, outCol).setValue("Slides / 幻灯片:").setFontWeight("bold");
  sheet.getRange(outRow, outCol + 1).setFormula(`=HYPERLINK("${presUrl}","Open Slides / 打开幻灯片")`);

  sheet.getRange(outRow + 1, outCol).setValue("Drive folder / 文件夹:").setFontWeight("bold");
  sheet.getRange(outRow + 1, outCol + 1).setFormula(`=HYPERLINK("${folderUrl}","Open folder / 打开文件夹")`);
}

/***************************************************************
 * COVER (Slide 1) – 52/62/36
 ***************************************************************/
function insertCoverInfoANTA_(slide, pres, info) {
  const { empresa, estadoZona, rangoSup, fecha } = info;

  removePageElementsByDescription_(slide, "AUTO_COVER_TITLE");
  removePageElementsByDescription_(slide, "AUTO_COVER_CLIENT");
  removePageElementsByDescription_(slide, "AUTO_COVER_SUB");

  cleanupCoverRightSideTextSafe_(slide, pres);

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

  const titleText = "Industrial property proposals for\n工业厂房/仓库方案 - 客户";
  const titleBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x, titleY, w, titleH);
  titleBox.setDescription("AUTO_COVER_TITLE");
  setTextBoxStyle_(titleBox, { text: titleText, font: SLIDES_ENG_CHI_CFG.COVER_FONT, size: 52, bold: true, color: SLIDES_ENG_CHI_CFG.COVER_COLOR_DARK });

  const clientY = titleY + titleH + gap1;
  const clientBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x, clientY, w, clientH);
  clientBox.setDescription("AUTO_COVER_CLIENT");
  setTextBoxStyle_(clientBox, {
    text: String(empresa || "").toUpperCase(),
    font: SLIDES_ENG_CHI_CFG.COVER_FONT,
    size: 62,
    bold: true,
    color: SLIDES_ENG_CHI_CFG.COVER_COLOR_BLUE
  });

  const subY = clientY + clientH + gap2;
  const subText = `${estadoZona}\n${rangoSup}\n${fecha}`;
  const subBox = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x, subY, w, subH);
  subBox.setDescription("AUTO_COVER_SUB");
  setTextBoxStyle_(subBox, { text: subText, font: SLIDES_ENG_CHI_CFG.COVER_FONT, size: 36, bold: false, color: SLIDES_ENG_CHI_CFG.COVER_SUB_COLOR });
}

function setTextBoxStyle_(shape, cfg) {
  const { text, font, size, bold, color } = cfg;

  const tr = shape.getText();
  tr.setText(text);

  const ts = tr.getTextStyle();
  ts.setFontFamily(font);
  ts.setFontSize(size);
  ts.setBold(!!bold);
  ts.setForegroundColor(color);

  safeSetParagraphAlignment_(tr, SlidesApp.ParagraphAlignment.CENTER);
}

function cleanupCoverRightSideTextSafe_(slide, pres) {
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

      // Remove common template words if they exist on the right side
      if (
        t.includes("propuestas") || t.includes("naves") || t.includes("industriales") ||
        t.includes("proposal") || t.includes("industrial") || t.includes("warehouse") ||
        t === "para"
      ) {
        shape.remove();
      }
    } catch (e) {}
  });
}

/***************************************************************
 * SAFE: setParagraphAlignment
 ***************************************************************/
function safeSetParagraphAlignment_(textRange, alignment) {
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
 * Compute content box (avoid title/footer)
 ***************************************************************/
function computeContentBox_(slide, pres) {
  const pageW = pres.getPageWidth();
  const pageH = pres.getPageHeight();

  let left = SLIDES_ENG_CHI_CFG.FALLBACK_SIDE_MARGIN;
  let right = pageW - SLIDES_ENG_CHI_CFG.FALLBACK_SIDE_MARGIN;
  let top = SLIDES_ENG_CHI_CFG.FALLBACK_TOP_MARGIN;
  let bottom = pageH - SLIDES_ENG_CHI_CFG.FALLBACK_BOTTOM_MARGIN;

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

      if (tl.includes("propuestas") || tl.includes("ubicaciones") || tl.includes("naves disponibles") ||
          tl.includes("proposals") || tl.includes("available") || tl.includes("locations")) {
        const b = sh.getTop() + sh.getHeight();
        titleBottom = (titleBottom == null) ? b : Math.max(titleBottom, b);
      }

      if (tl.includes("+52") || tl.includes("@industrialestatemexico.com") || tl.includes("industrialestatemexico.com")) {
        footerTop = (footerTop == null) ? sh.getTop() : Math.min(footerTop, sh.getTop());
      }
    } catch (e) {}
  });

  if (titleBottom != null) top = Math.max(top, titleBottom + SLIDES_ENG_CHI_CFG.TITLE_TO_CONTENT_GAP);
  if (footerTop != null) bottom = Math.min(bottom, footerTop - SLIDES_ENG_CHI_CFG.FOOTER_SAFETY_GAP);

  return {
    left,
    top,
    width: Math.max(50, right - left),
    height: Math.max(50, bottom - top)
  };
}

/***************************************************************
 * Slide 6: insert table IMAGE + separate footnote
 ***************************************************************/
function insertTableImageOnSlide_(slide, pres, imgBlob, footnoteText) {
  slide.getPageElements().forEach(pe => {
    try { if (pe.getPageElementType() === SlidesApp.PageElementType.TABLE) pe.remove(); } catch (e) {}
  });
  removePageElementsByDescription_(slide, "AUTO_TABLE_IMG");
  removePageElementsByDescription_(slide, "AUTO_TABLE_FOOTNOTE");

  const box = computeContentBox_(slide, pres);

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

  // Footnote (separate, no overlap)
  const footY = imgBox.top + imgBox.height + gap;
  const foot = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, box.left, footY, box.width, footH);
  foot.setDescription("AUTO_TABLE_FOOTNOTE");

  foot.getText().setText(footnoteText || "");
  const ts = foot.getText().getTextStyle();
  ts.setFontFamily("Arial").setFontSize(12).setItalic(true).setForegroundColor("#000000");
  safeSetParagraphAlignment_(foot.getText(), SlidesApp.ParagraphAlignment.LEFT);
}

/***************************************************************
 * Slide 7: insert MAP image
 ***************************************************************/
function insertMapOnSlide_(slide, pres, mapBlob) {
  removePageElementsByDescription_(slide, "AUTO_MAP");

  // (Opcional) elimina imágenes anteriores para evitar que algo lo tape
  slide.getPageElements().forEach(pe => {
    try {
      if (pe.getPageElementType() === SlidesApp.PageElementType.IMAGE) pe.remove();
    } catch (e) {}
  });

  const box = computeContentBox_(slide, pres);

  // Validación básica del blob
  const ct = String(mapBlob.getContentType() || "");
  if (!ct.startsWith("image/")) {
    throw new Error(`El blob del mapa NO es imagen. contentType=${ct}`);
  }

  const img = slide.insertImage(mapBlob);
  img.setDescription("AUTO_MAP");

  // A veces Slides regresa 0x0 inmediatamente; reintenta
  Utilities.sleep(200);

  let iw = img.getWidth();
  let ih = img.getHeight();

  if (!iw || !ih) {
    Utilities.sleep(600);
    iw = img.getWidth();
    ih = img.getHeight();
  }

  // Fallback duro: si sigue en 0, colócalo “a caja completa”
  if (!iw || !ih) {
    img.setLeft(box.left);
    img.setTop(box.top);
    img.setWidth(box.width);
    img.setHeight(box.height);
    return;
  }

  const scale = Math.min(box.width / iw, box.height / ih);

  img.setWidth(iw * scale);
  img.setHeight(ih * scale);
  img.setLeft(box.left + (box.width - img.getWidth()) / 2);
  img.setTop(box.top + (box.height - img.getHeight()) / 2);
}

/***************************************************************
 * FINAL TABLE: detect last "Partida" OR "Item" that has "REF" near
 ***************************************************************/
function findFinalTableMeta_(sheet) {
  const matchesPartida = sheet.createTextFinder("Partida")
    .matchCase(false)
    .matchEntireCell(false)
    .findAll() || [];

  const matchesItem = sheet.createTextFinder("Item")
    .matchCase(false)
    .matchEntireCell(false)
    .findAll() || [];

  const matches = matchesPartida.concat(matchesItem);
  if (!matches || matches.length === 0) throw new Error('Could not find table header "Partida" or "Item".');

  let best = null;
  for (const cell of matches) {
    const r = cell.getRow();
    const c = cell.getColumn();
    const w = Math.min(25, sheet.getLastColumn() - c + 1);
    const rowVals = sheet.getRange(r, c, 1, w).getDisplayValues()[0].map(v => String(v || "").trim().toLowerCase());
    const hasRef = rowVals.includes("ref") || rowVals.some(v => v === "ref / 参考编号" || v.includes("ref"));
    if (hasRef) if (!best || r > best.getRow()) best = cell;
  }
  if (!best) best = matches.reduce((a, b) => (a.getRow() > b.getRow() ? a : b));

  const headerRow = best.getRow();
  const colPartida = best.getColumn();
  const headers = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getDisplayValues()[0].map(v => String(v || "").trim());
  return { headerRow, colPartida, headers };
}

/***************************************************************
 * Read table for Slides (stops at first empty header cell)
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

  // take until first empty header cell
  let end = headerSlice.length;
  for (let i = 0; i < headerSlice.length; i++) {
    if (String(headerSlice[i] || "").trim() === "") { end = i; break; }
  }
  if (end <= 0) throw new Error('Found the header cell but could not read headers to the right.');

  const headers = headerSlice.slice(0, end).map(h => String(h || "").trim());

  const idxPartida = headers.findIndex(h => {
    const n = norm_(h);
    return n === "partida" || n === "item";
  });

  let idxRef = headers.findIndex(h => norm_(h) === "ref");
  if (idxRef < 0) {
    // fallback: look for any header containing "ref"
    idxRef = headers.findIndex(h => String(h || "").toLowerCase().includes("ref"));
  }
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
 * Read real column widths from Sheets (px)
 ***************************************************************/
function getSheetTableWidthsPx_(sheet, startCol, count) {
  const out = [];
  for (let i = 0; i < count; i++) out.push(sheet.getColumnWidth(startCol + i));
  return out;
}

/***************************************************************
 * Render table to PNG (NO footnote)
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

  // Correct signature: insertTable(rows, cols, left, top, width, height)
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

  // Header row
  for (let c = 0; c < cols; c++) {
    const cell = table.getCell(0, c);
    cell.getText().setText(String(tableData.headers[c] || ""));
    try { cell.getFill().setSolidFill(HEADER_BG); } catch (e) {}

    try {
      const ts = cell.getText().getTextStyle();
      ts.setFontFamily("Arial").setFontSize(SLIDES_ENG_CHI_CFG.HEADER_FONT_SIZE).setBold(true).setForegroundColor("#000000");
      safeSetParagraphAlignment_(cell.getText(), SlidesApp.ParagraphAlignment.CENTER);
    } catch (e) {}

    setCellBordersBlack_(cell);
  }

  // Body
  for (let r = 0; r < tableData.values.length; r++) {
    for (let c = 0; c < cols; c++) {
      const cell = table.getCell(r + 1, c);
      cell.getText().setText(String(tableData.values[r][c] ?? ""));

      try {
        const ts = cell.getText().getTextStyle();
        ts.setFontFamily("Arial").setFontSize(SLIDES_ENG_CHI_CFG.BODY_FONT_SIZE).setBold(false).setForegroundColor("#000000");
        safeSetParagraphAlignment_(cell.getText(), SlidesApp.ParagraphAlignment.CENTER);
      } catch (e) {}

      setCellBordersBlack_(cell);
    }
  }

  tmpPres.saveAndClose();

  const tableObjectId = getFirstTableObjectId_(tmpId);

  // Column widths via API (clamp >= 32pt)
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
  const url = `https://docs.google.com/presentation/d/${tmpId}/export/png?pageid=${pageId}&scale=${SLIDES_ENG_CHI_CFG.PNG_EXPORT_SCALE}`;
  const resp = UrlFetchApp.fetch(url, {
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });

  if (resp.getResponseCode() !== 200) {
    try { DriveApp.getFileById(tmpId).setTrashed(true); } catch (e) {}
    throw new Error(`PNG export failed. HTTP ${resp.getResponseCode()}: ${resp.getContentText()}`);
  }

  const png = resp.getBlob().setName("final_table.png");
  try { DriveApp.getFileById(tmpId).setTrashed(true); } catch (e) {}
  return png;
}

/***************************************************************
 * Clamp columns: ensure >= minPt and sum = totalWidthPt
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
 * Slides API: updateTableColumnProperties (batchUpdate)
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
    throw new Error("Slides API batchUpdate (col widths) failed: HTTP " + resp.getResponseCode() + " | " + resp.getContentText());
  }
}

/***************************************************************
 * Slides API: updateTableRowProperties (minRowHeight)
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
    throw new Error("Slides API batchUpdate (row heights) failed: HTTP " + resp.getResponseCode() + " | " + resp.getContentText());
  }
}

/***************************************************************
 * Get objectId of the first TABLE (Slides API GET)
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
 * Get objectId of first slide (pageid) for export/png
 ***************************************************************/
function getFirstSlideObjectIdFromPresentation_(presentationId) {
  const pres = fetchSlidesPresentation_(presentationId);
  if (!pres.slides || !pres.slides[0]) throw new Error("Could not read temp slide.");
  return pres.slides[0].objectId;
}

function fetchSlidesPresentation_(presentationId) {
  const url = `https://slides.googleapis.com/v1/presentations/${presentationId}`;
  const resp = UrlFetchApp.fetch(url, {
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });
  if (resp.getResponseCode() !== 200) {
    throw new Error("Slides API GET failed: HTTP " + resp.getResponseCode() + " | " + resp.getContentText());
  }
  return JSON.parse(resp.getContentText());
}

/***************************************************************
 * STRICT MAP: ONLY ROOT / MAPS_FOLDER_NAME
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
    `Map not found for sheet "${sheetName}". / 未找到该工作表对应地图。\n` +
    `It must exist in ROOT/${MAPS_FOLDER_NAME} named exactly:\n- ${expected}\n\n` +
    `Tip: run your map generator first on that sheet. / 提示：请先在该工作表运行地图生成。`
  );
}

/***************************************************************
 * Cleanup by description (AUTO_*)
 ***************************************************************/
function removePageElementsByDescription_(slide, desc) {
  slide.getPageElements().forEach(pe => {
    try { if (pe.getDescription && pe.getDescription() === desc) pe.remove(); } catch (e) {}
  });
}

/***************************************************************
 * Folder helpers (inside OUTPUT_ROOT_FOLDER_ID)
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
  return `Proposal - ${safeEmpresa} - ${safeSheet} - ${fecha}`;
}

/***************************************************************
 * UI helper
 ***************************************************************/
function promptOrThrow_(ui, title, message) {
  const res = ui.prompt(title, message, ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) throw new Error("Operation cancelled. / 操作已取消。");
  const val = String(res.getResponseText() || "").trim();
  if (!val) throw new Error("Empty value. Cancelled. / 输入为空，已取消。");
  return val;
}
