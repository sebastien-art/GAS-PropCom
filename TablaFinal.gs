function CrearTablaFinal() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  // Quita gridlines (las líneas grises de la hoja)
  sheet.setHiddenGridlines(true);

  const HEADER_ROW = 1;

  // Tabla final inicia en columna D
  const START_COL_TABLE = 4; // D

  // Marco blanco alrededor (para screenshot)
  const PAD_ROWS = 2;
  const PAD_LEFT = 1;
  const PAD_RIGHT = 1;

  // Leyenda
  const FOOTNOTE_TEXT = "* Precio (shell) antes de negociación y proyecto técnico + mantenimiento + IVA";

  // Color encabezados (Google Sheets “Light blue 3”)
  const HEADER_BG = "#9fc5e8";

  // Encabezados origen
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

  // ✅ Encabezados destino
  const FINAL_HEADERS = [
    "Partida",
    "REF",
    "Estado",
    "Zona Principal",
    "Sub Zona",
    "M2 de construcción",
    "Asking price / m2 *",
    "Mantenimiento / m2",
    "Disponibilidad",
    "", // columna vacía separadora
    "Coordenadas",
    "Parque"
  ];

  const SEP_INDEX = FINAL_HEADERS.indexOf("") + 1; // 10
  const DISP_INDEX = FINAL_HEADERS.indexOf("Disponibilidad") + 1; // 9
  const COORD_INDEX = FINAL_HEADERS.indexOf("Coordenadas") + 1;   // 11
  const M2CONST_INDEX = FINAL_HEADERS.indexOf("M2 de construcción") + 1;

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
  const colCOORD = getCol(SRC_HEADERS.COORD);
  const colPARQUE = getCol(SRC_HEADERS.PARQUE);

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
    ["Coordenadas", colCOORD],
    ["Parque", colPARQUE]
  ];

  const missing = required.filter(([, c]) => !c).map(([n]) => n);
  if (missing.length) {
    throw new Error("No encontré estos encabezados en la fila 1: " + missing.join(", "));
  }

  const scanFrom = HEADER_ROW + 1;
  const scanTo = sheet.getLastRow();
  if (scanTo < scanFrom) return;

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
  let partida = 1;

  for (let r = 0; r < data.length; r++) {
    const row = data[r];
    const enviarVal = String(row[colENVIAR - 1] ?? "").trim().toUpperCase();
    if (enviarVal === "OK") {
      out.push([
        partida++,
        row[colREF - 1],
        row[colESTADO - 1],
        row[colZONA - 1],
        row[colSUBZONA - 1],
        row[colM2CONST - 1],
        row[colASKING - 1],
        row[colMANT - 1],
        row[colDISP - 1],
        "", // separador vacío
        row[colCOORD - 1],
        row[colPARQUE - 1]
      ]);
    }
  }

  const startRowTable = lastDataRow + 4;
  const numColsTable = FINAL_HEADERS.length;
  const numRowsTable = 1 + out.length;

  // +1 fila para la leyenda
  const totalContentRows = numRowsTable + 1;

  const clearStartRow = Math.max(1, startRowTable - PAD_ROWS);
  const clearStartCol = Math.max(1, START_COL_TABLE - PAD_LEFT);

  const blockRows = totalContentRows + (PAD_ROWS * 2);
  const blockCols = numColsTable + PAD_LEFT + PAD_RIGHT;

  const neededEndRow = clearStartRow + blockRows - 1;
  const maxRows = sheet.getMaxRows();
  if (neededEndRow > maxRows) sheet.insertRowsAfter(maxRows, neededEndRow - maxRows);

  // Limpia + deja blanco todo el bloque (tabla + marco)
  const blockRange = sheet.getRange(clearStartRow, clearStartCol, blockRows, blockCols);

  // ✅ FIX: rompe merges previos para que no "desaparezcan" encabezados
  blockRange.breakApart();

  blockRange.clearContent();
  blockRange.setBackground("#ffffff");
  blockRange.setBorder(false, false, false, false, false, false);

  // Encabezados (Light blue 3)
  const headerRange = sheet.getRange(startRowTable, START_COL_TABLE, 1, numColsTable);
  headerRange.setValues([FINAL_HEADERS]);
  headerRange.setFontWeight("bold");
  headerRange.setBackground(HEADER_BG);
  headerRange.setFontColor("#000000");

  // Encabezados: wrap (NO overflow)
  headerRange.setWrap(true);
  headerRange.setHorizontalAlignment("center");
  headerRange.setVerticalAlignment("middle");

  // Datos
  if (out.length) {
    sheet.getRange(startRowTable + 1, START_COL_TABLE, out.length, numColsTable).setValues(out);
  }

  // Bordes base de toda la tabla
  const tableRange = sheet.getRange(startRowTable, START_COL_TABLE, Math.max(1, numRowsTable), numColsTable);
  tableRange.setBorder(true, true, true, true, true, true);

  // Centrar TODA la tabla
  tableRange.setHorizontalAlignment("center");
  tableRange.setVerticalAlignment("middle");

  // ✅ Textwrapping OVERFLOW solo en datos (no encabezados)
  if (out.length) {
    const dataRange = sheet.getRange(startRowTable + 1, START_COL_TABLE, out.length, numColsTable);
    dataRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW);
  }

  // ====== SEPARADOR: sin color ni bordes ======
  const rowsForBorders = Math.max(1, numRowsTable);
  const sepAbsCol = START_COL_TABLE + SEP_INDEX - 1;

  const sepColRange = sheet.getRange(startRowTable, sepAbsCol, rowsForBorders, 1);
  sepColRange.clearFormat();
  sepColRange.setBorder(false, false, false, false, false, false);

  // Re-forzar bordes que se pierden por el separador
  const dispAbsCol = START_COL_TABLE + DISP_INDEX - 1;
  const coordAbsCol = START_COL_TABLE + COORD_INDEX - 1;

  const dispRange = sheet.getRange(startRowTable, dispAbsCol, rowsForBorders, 1);
  const coordRange = sheet.getRange(startRowTable, coordAbsCol, rowsForBorders, 1);

  dispRange.setBorder(null, null, null, true, null, null);
  coordRange.setBorder(null, true, null, null, null, null);

  // Leyenda (1 fila abajo de la tabla)
  const footnoteRow = startRowTable + numRowsTable;
  const footnoteRange = sheet.getRange(footnoteRow, START_COL_TABLE, 1, numColsTable);
  footnoteRange.breakApart();
  footnoteRange.merge();
  footnoteRange.setValue(FOOTNOTE_TEXT);
  footnoteRange.setBackground("#ffffff");
  footnoteRange.setFontStyle("italic");
  footnoteRange.setHorizontalAlignment("left");
  footnoteRange.setVerticalAlignment("middle");
  footnoteRange.setWrap(true);

  // Ancho 100 px tabla + márgenes laterales
  sheet.setColumnWidths(START_COL_TABLE, numColsTable, 100);
  if (PAD_LEFT > 0) sheet.setColumnWidths(START_COL_TABLE - PAD_LEFT, PAD_LEFT, 100);
  if (PAD_RIGHT > 0) sheet.setColumnWidths(START_COL_TABLE + numColsTable, PAD_RIGHT, 100);

  // "Fit to data" (lo más cercano): autoResizeRows sobre headers + datos + leyenda
  sheet.setRowHeights(startRowTable, totalContentRows, 22);
  sheet.autoResizeRows(startRowTable, totalContentRows);

  const flexText =
    "** Superficie Flexible: podemos entregar la cantidad de m² de acuerdo a su requerimiento, de acuerdo al rango mostrado (desde la superficie mínima rentable hasta la máxima rentable).";

  /***************************************************************
   * CAMBIO PEDIDO:
   * - Texto en azul en "M2 de construcción" (NO encabezado)
   * - Formato número con separador de millar y sin decimales
   ***************************************************************/
  if (out.length && M2CONST_INDEX > 0) {
    const m2AbsCol = START_COL_TABLE + M2CONST_INDEX - 1;
    const m2DataRange = sheet.getRange(startRowTable + 1, m2AbsCol, out.length, 1);
    m2DataRange.setFontColor("#1a73e8");
    m2DataRange.setNumberFormat("#,##0");
  }
}
