function CrearTablaFinal_ENG_CHI() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  // Hide gridlines
  sheet.setHiddenGridlines(true);

  const HEADER_ROW = 1;

  // Final table starts at column D
  const START_COL_TABLE = 4; // D

  // White padding frame (for screenshot)
  const PAD_ROWS = 2;
  const PAD_LEFT = 1;
  const PAD_RIGHT = 1;

  // Header color
  const HEADER_BG = "#9fc5e8";

  // ✅ Footnote (ENG/CHI only)
  const FOOTNOTE_TEXT =
    "* Price (shell) before negotiation and technical project + maintenance + VAT / *价格为谈判及技术方案之前的报价，另加维护费和增值税（VAT）。";

  // ✅ Flexible note (ENG/CHI only)
  const FLEX_TEXT =
    "** Flexible area: we can deliver the number of m² according to your requirement, within the range shown (from the minimum leasable area to the maximum leasable area). / ** 面积灵活：我们可根据您的需求在所示范围内交付相应的平方米数（从最小可租面积到最大可租面积）。";

  // Source headers (as they exist in your sheet)
  const SRC_HEADERS = {
    ENVIAR: "ENVIAR",
    REF: "REF",
    ESTADO: "Estado",
    ZONA: "Zona Principal",
    SUBZONA: "Sub Zona",
    M2CONST: "M2 de construcción",
    ASKING: "Asking price /m2",
    MANT: "Mantenimiento / m2",
    DISP: "Disponibilidad",
    COORD: "Coordenadas",
    PARQUE: "Parque"
  };

  // ✅ Destination headers (ENG/CHI only)
  const FINAL_HEADERS = [
    "Item / 项目",
    "REF / 参考编号",
    "State / 州",
    "Main zone / 主区域",
    "Sub-zone / 子区域",
    "Built area (m²) / 建筑面积（平方米）",
    "Asking price per m² * / 要价（每平方米）*",
    "Maintenance per m² / 维护费（每平方米）",
    "Availability / 可用性",
    "", // separator column
    "Coordinates",
    "Park"
  ];

  const SEP_INDEX = FINAL_HEADERS.indexOf("") + 1; // 10
  const DISP_INDEX = FINAL_HEADERS.indexOf("Availability / 可用性") + 1; // 9
  const COORD_INDEX = FINAL_HEADERS.indexOf("Coordinates") + 1; // 11
  const M2CONST_INDEX = FINAL_HEADERS.indexOf("Built area (m²) / 建筑面积（平方米）") + 1;

  // ✅ Footnotes merged only until last column BEFORE separator
  const FOOTNOTE_MERGE_COLS = SEP_INDEX - 1;

  // ✅ Separator not empty to avoid visual overflow
  const SEP_FILL = " ";

  const norm = (s) => String(s ?? "").replace(/\s+/g, " ").trim().toLowerCase();

  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return;

  const headerValues = sheet.getRange(HEADER_ROW, 1, 1, lastCol).getValues()[0];
  const headerMap = new Map();
  headerValues.forEach((h, i) => {
    const key = norm(h);
    if (key) headerMap.set(key, i + 1);
  });

  const getCol = (headerName) => headerMap.get(norm(headerName)) || null;

  const colENVIAR = getCol(SRC_HEADERS.ENVIAR);
  const colREF = getCol(SRC_HEADERS.REF);
  const colESTADO = getCol(SRC_HEADERS.ESTADO);
  const colZONA = getCol(SRC_HEADERS.ZONA);
  const colSUBZONA = getCol(SRC_HEADERS.SUBZONA);
  const colM2CONST = getCol(SRC_HEADERS.M2CONST);
  const colASKING = getCol(SRC_HEADERS.ASKING);
  const colMANT = getCol(SRC_HEADERS.MANT);
  const colDISP = getCol(SRC_HEADERS.DISP);

  // ✅ FIX: acepta Coordenadas o Coordinates (y variantes)
  const colCOORD =
    getCol(SRC_HEADERS.COORD) ||
    getCol("Coordinates") ||
    getCol("Coord") ||
    getCol("GPS") ||
    getCol("Coordenadas (GPS)");

  // ✅ FIX: acepta Parque o Park
  const colPARQUE =
    getCol(SRC_HEADERS.PARQUE) ||
    getCol("Park");

  const required = [
    ["ENVIAR", colENVIAR],
    ["REF", colREF],
    ["Estado", colESTADO],
    ["Zona Principal", colZONA],
    ["Sub Zona", colSUBZONA],
    ["M2 de construcción", colM2CONST],
    ["Asking price /m2", colASKING],
    ["Mantenimiento / m2", colMANT],
    ["Disponibilidad", colDISP],
    ["Coordenadas/Coordinates", colCOORD],
    ["Parque/Park", colPARQUE]
  ];

  const missing = required.filter(([, c]) => !c).map(([n]) => n);
  if (missing.length) {
    throw new Error("Missing headers in row 1: " + missing.join(", "));
  }

  const scanFrom = HEADER_ROW + 1;
  const scanTo = sheet.getLastRow();
  if (scanTo < scanFrom) return;

  // Find lastDataRow based on REF column
  const refColValues = sheet.getRange(scanFrom, colREF, scanTo - HEADER_ROW, 1).getValues();
  let lastDataRow = HEADER_ROW;
  for (let i = refColValues.length - 1; i >= 0; i--) {
    if (String(refColValues[i][0] ?? "").trim() !== "") {
      lastDataRow = scanFrom + i;
      break;
    }
  }
  if (lastDataRow === HEADER_ROW) return;

  const dataNumRows = lastDataRow - HEADER_ROW;
  const data = sheet.getRange(scanFrom, 1, dataNumRows, lastCol).getValues();

  const out = [];
  let item = 1;

  for (let r = 0; r < data.length; r++) {
    const row = data[r];
    const enviarVal = String(row[colENVIAR - 1] ?? "").trim().toUpperCase();
    if (enviarVal === "OK") {
      out.push([
        item++,
        row[colREF - 1],
        row[colESTADO - 1],
        row[colZONA - 1],
        row[colSUBZONA - 1],
        row[colM2CONST - 1],
        row[colASKING - 1],
        row[colMANT - 1],
        row[colDISP - 1],
        SEP_FILL, // separator
        row[colCOORD - 1],   // ✅ ahora sí toma Coordinates/Coordenadas
        row[colPARQUE - 1]
      ]);
    }
  }

  const startRowTable = lastDataRow + 4;
  const numColsTable = FINAL_HEADERS.length;
  const numRowsTable = 1 + out.length;

  // +1 row footnote, +1 row flex note
  const totalContentRows = numRowsTable + 2;

  const clearStartRow = Math.max(1, startRowTable - PAD_ROWS);
  const clearStartCol = Math.max(1, START_COL_TABLE - PAD_LEFT);

  const blockRows = totalContentRows + (PAD_ROWS * 2);
  const blockCols = numColsTable + PAD_LEFT + PAD_RIGHT;

  const neededEndRow = clearStartRow + blockRows - 1;
  const maxRows = sheet.getMaxRows();
  if (neededEndRow > maxRows) sheet.insertRowsAfter(maxRows, neededEndRow - maxRows);

  // Clear + white background
  const blockRange = sheet.getRange(clearStartRow, clearStartCol, blockRows, blockCols);

  // ✅ break previous merges
  blockRange.breakApart();

  blockRange.clearContent();
  blockRange.setBackground("#ffffff");
  blockRange.setBorder(false, false, false, false, false, false);

  // Headers
  const headerRange = sheet.getRange(startRowTable, START_COL_TABLE, 1, numColsTable);
  headerRange.setValues([FINAL_HEADERS]);
  headerRange.setFontWeight("bold");
  headerRange.setBackground(HEADER_BG);
  headerRange.setFontColor("#000000");
  headerRange.setWrap(true);
  headerRange.setHorizontalAlignment("center");
  headerRange.setVerticalAlignment("middle");

  // Data
  if (out.length) {
    sheet.getRange(startRowTable + 1, START_COL_TABLE, out.length, numColsTable).setValues(out);
  }

  // Table borders + alignment
  const tableRange = sheet.getRange(startRowTable, START_COL_TABLE, Math.max(1, numRowsTable), numColsTable);
  tableRange.setBorder(true, true, true, true, true, true);
  tableRange.setHorizontalAlignment("center");
  tableRange.setVerticalAlignment("middle");

  // Overflow wrap only for data
  if (out.length) {
    const dataRange = sheet.getRange(startRowTable + 1, START_COL_TABLE, out.length, numColsTable);
    dataRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW);
  }

  // Separator: no borders/format
  const rowsForBorders = Math.max(1, numRowsTable);
  const sepAbsCol = START_COL_TABLE + SEP_INDEX - 1;
  const sepColRange = sheet.getRange(startRowTable, sepAbsCol, rowsForBorders, 1);
  sepColRange.clearFormat();
  sepColRange.setBorder(false, false, false, false, false, false);

  // Reinforce borders on both sides of separator
  const dispAbsCol = START_COL_TABLE + DISP_INDEX - 1;
  const coordAbsCol = START_COL_TABLE + COORD_INDEX - 1;
  sheet.getRange(startRowTable, dispAbsCol, rowsForBorders, 1).setBorder(null, null, null, true, null, null);
  sheet.getRange(startRowTable, coordAbsCol, rowsForBorders, 1).setBorder(null, true, null, null, null, null);

  // Footnote row (merge only before separator)
  const footnoteRow = startRowTable + numRowsTable;
  const footnoteRange = sheet.getRange(footnoteRow, START_COL_TABLE, 1, FOOTNOTE_MERGE_COLS);
  footnoteRange.breakApart();
  footnoteRange.merge();
  footnoteRange.setValue(FOOTNOTE_TEXT);
  footnoteRange.setBackground("#ffffff");
  footnoteRange.setFontStyle("italic");
  footnoteRange.setHorizontalAlignment("left");
  footnoteRange.setVerticalAlignment("middle");
  footnoteRange.setWrap(true);

  // Flex note row
  const flexRow = footnoteRow + 1;
  const flexRange = sheet.getRange(flexRow, START_COL_TABLE, 1, FOOTNOTE_MERGE_COLS);
  flexRange.breakApart();
  flexRange.merge();
  flexRange.setValue(FLEX_TEXT);
  flexRange.setBackground("#ffffff");
  flexRange.setFontStyle("italic");
  flexRange.setHorizontalAlignment("left");
  flexRange.setVerticalAlignment("middle");
  flexRange.setWrap(true);

  // Column widths
  sheet.setColumnWidths(START_COL_TABLE, numColsTable, 100);
  if (PAD_LEFT > 0) sheet.setColumnWidths(START_COL_TABLE - PAD_LEFT, PAD_LEFT, 100);
  if (PAD_RIGHT > 0) sheet.setColumnWidths(START_COL_TABLE + numColsTable, PAD_RIGHT, 100);

  // Row sizing
  sheet.setRowHeights(startRowTable, totalContentRows, 22);
  sheet.autoResizeRows(startRowTable, totalContentRows);

  // ✅ Built area column: blue + thousand separator, no decimals (data only)
  if (out.length && M2CONST_INDEX > 0) {
    const m2AbsCol = START_COL_TABLE + M2CONST_INDEX - 1;
    const m2DataRange = sheet.getRange(startRowTable + 1, m2AbsCol, out.length, 1);
    m2DataRange.setFontColor("#1a73e8");
    m2DataRange.setNumberFormat("#,##0");
  }

  sheet.setActiveRange(headerRange);
}
