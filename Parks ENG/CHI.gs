function Park_ENG_CHI() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const ui = SpreadsheetApp.getUi();

  sheet.setHiddenGridlines(true);

  const HEADER_ROW = 1;
  const START_COL_TABLE = 4; // D

  const PAD_ROWS = 2;
  const PAD_LEFT = 1;
  const PAD_RIGHT = 1;

  const HEADER_BG = "#9fc5e8";

  // ✅ Leyendas (ENG/CHI)
  const FOOTNOTE_TEXT =
    "* Price (shell) before negotiation and technical project + maintenance + VAT / *价格为谈判及技术方案之前的报价，另加维护费和增值税（VAT）。";

  const FLEX_TEXT =
    "** Flexible area: we can deliver the number of m² according to your requirement, within the range shown (from the minimum leasable area to the maximum leasable area). / ** 面积灵活：我们可根据您的需求在所示范围内交付相应的平方米数（从最小可租面积到最大可租面积）。";

  // Encabezados origen (agrego alias "Operación" por si así viene tu hoja)
  const SRC_HEADERS = {
    ENVIAR: ["ENVIAR", "Operación", "Operacion"],
    REF: ["REF"],
    ESTADO: ["Estado"],
    ZONA: ["Zona Principal"],
    SUBZONA: ["Sub Zona", "Subzona", "Sub-Zona", "Subzona Principal"],
    M2CONST: ["M2 de construcción", "M2 de construccion", "m2 de construcción", "m2 de construccion"],
    ASKING: ["Asking price /m2", "Asking price/m2", "Asking price / m2"],
    MANT: ["Mantenimiento / m2", "Mantenimiento/m2", "Mantenimiento /m2", "Maintenance / m2", "Maintenance/m2"],
    DISP: ["Disponibilidad"],
    DESARROLLADOR: ["Desarrollador"],
    PARQUE: ["Parque"],
    COORD: ["Coordenadas", "Coordenadas (GPS)", "GPS", "Coord", "Coordinates"]
  };

  // ✅ Encabezados destino (lo que quieres en lugar del español)
  const FINAL_HEADERS = [
    "Item / 项目",
    "REF / 参考编号",
    "State / 州",
    "Sub-zone / 子区域",
    "Suggested area (m²) / 建议面积（平方米）",
    "Minimum leasable area (m²) ** / 最小可租面积（平方米）**",
    "Maximum leasable area (m²) ** / 最大可租面积（平方米）**",
    "Asking price per m² * / 要价（每平方米）*",
    "Availability / 可用性",
    "", // separador
    "Coordinates",
    "Park"
  ];

  // Índices (1-based dentro de FINAL_HEADERS)
  const SUG_INDEX = FINAL_HEADERS.indexOf("Suggested area (m²) / 建议面积（平方米）") + 1;
  const MIN_INDEX = FINAL_HEADERS.indexOf("Minimum leasable area (m²) ** / 最小可租面积（平方米）**") + 1;
  const MAX_INDEX = FINAL_HEADERS.indexOf("Maximum leasable area (m²) ** / 最大可租面积（平方米）**") + 1;
  const SEP_INDEX = FINAL_HEADERS.indexOf("") + 1;
  const DISP_INDEX = FINAL_HEADERS.indexOf("Availability / 可用性") + 1;
  const COORD_INDEX = FINAL_HEADERS.indexOf("Coordinates") + 1;

  // ✅ Auto-ajustar ancho SOLO en Sub-zone
  const SUBZONA_OUT_INDEX = FINAL_HEADERS.indexOf("Sub-zone / 子区域") + 1;

  /************ Helpers encabezados (tolerante + sin acentos) ************/
  const stripAccents = (s) =>
    String(s ?? "")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "");

  const makeKeys = (s) => {
    let t = stripAccents(String(s ?? "").trim().toLowerCase()).replace(/\s+/g, " ");
    const k1 = t;
    const k2 = t.replace(/\s+/g, "");
    const k3 = k2.replace(/\//g, "");
    return [k1, k2, k3];
  };

  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return;

  const headerValues = sheet.getRange(HEADER_ROW, 1, 1, lastCol).getValues()[0];
  const headerMap = new Map();
  headerValues.forEach((h, i) => {
    makeKeys(h).forEach((k) => {
      if (k && !headerMap.has(k)) headerMap.set(k, i + 1);
    });
  });

  const getCol = (namesArray) => {
    for (const name of namesArray) {
      for (const k of makeKeys(name)) {
        if (headerMap.has(k)) return headerMap.get(k);
      }
    }
    return null;
  };

  const colENVIAR = getCol(SRC_HEADERS.ENVIAR);
  const colREF = getCol(SRC_HEADERS.REF);
  const colESTADO = getCol(SRC_HEADERS.ESTADO);
  const colSUBZONA = getCol(SRC_HEADERS.SUBZONA);
  const colM2CONST = getCol(SRC_HEADERS.M2CONST);
  const colASKING = getCol(SRC_HEADERS.ASKING);
  const colDISP = getCol(SRC_HEADERS.DISP);
  const colDEV = getCol(SRC_HEADERS.DESARROLLADOR);
  const colPARQUE = getCol(SRC_HEADERS.PARQUE);
  const colCOORD = getCol(SRC_HEADERS.COORD); // opcional

  const required = [
    ["ENVIAR", colENVIAR],
    ["REF", colREF],
    ["Estado", colESTADO],
    ["Sub Zona", colSUBZONA],
    ["M2 de construcción", colM2CONST],
    ["Asking price /m2", colASKING],
    ["Disponibilidad", colDISP],
    ["Desarrollador", colDEV],
    ["Parque", colPARQUE]
  ];

  const missing = required.filter(([, c]) => !c).map(([n]) => n);
  if (missing.length) {
    throw new Error("No encontré estos encabezados en la fila 1: " + missing.join(", "));
  }

  /************ Helper numérico ************/
  const toNumber = (v) => {
    if (v === null || v === "") return null;
    if (typeof v === "number") return v;
    if (v instanceof Date) return null;

    let s = String(v).trim();
    if (!s) return null;

    s = s.replace(/[^\d.,-]/g, "");
    const comma = s.lastIndexOf(",");
    const dot = s.lastIndexOf(".");
    if (comma > -1 && dot > -1) {
      if (dot > comma) s = s.replace(/,/g, "");
      else s = s.replace(/\./g, "").replace(",", ".");
    } else if (comma > -1 && dot === -1) {
      s = s.replace(/,/g, "");
    }
    const n = parseFloat(s);
    return isNaN(n) ? null : n;
  };

  const scanFrom = HEADER_ROW + 1;
  const scanTo = sheet.getLastRow();
  if (scanTo < scanFrom) return;

  // Detecta lastDataRow por REF (origen)
  const refVals = sheet.getRange(scanFrom, colREF, scanTo - HEADER_ROW, 1).getValues();
  let lastDataRow = HEADER_ROW;
  for (let i = refVals.length - 1; i >= 0; i--) {
    if (String(refVals[i][0] ?? "").trim() !== "") {
      lastDataRow = scanFrom + i;
      break;
    }
  }
  if (lastDataRow === HEADER_ROW) {
    ui.alert("No encontré datos en la tabla origen (REF vacío).");
    return;
  }

  const dataNumRows = lastDataRow - HEADER_ROW;
  const data = sheet.getRange(scanFrom, 1, dataNumRows, lastCol).getValues();

  // Filtra ENVIAR == OK
  const picked = [];
  for (let r = 0; r < data.length; r++) {
    const row = data[r];
    const enviarVal = stripAccents(String(row[colENVIAR - 1] ?? "")).trim().toUpperCase();
    if (/^OK\b/.test(enviarVal)) picked.push(row);
  }
  if (!picked.length) {
    ui.alert("No encontré filas con ENVIAR = OK.");
    return;
  }

  /************ Agrupación por (Desarrollador, Parque) ************/
  const keyOf = (row) => {
    const dev = String(row[colDEV - 1] ?? "").trim();
    const parque = String(row[colPARQUE - 1] ?? "").trim();
    if (!dev || !parque) return null;
    return stripAccents(dev).toLowerCase() + "||" + stripAccents(parque).toLowerCase();
  };

  const counts = new Map();
  picked.forEach((row) => {
    const k = keyOf(row);
    if (!k) return;
    counts.set(k, (counts.get(k) || 0) + 1);
  });

  // Guarda min/max + refs + coords por grupo
  const groups = new Map();
  picked.forEach((row) => {
    const k = keyOf(row);
    if (!k) return;

    if (!groups.has(k)) {
      groups.set(k, {
        baseRow: row,
        minM2: null,
        maxM2: null,
        refs: [],
        coords: []
      });
    }
    const g = groups.get(k);

    const m2 = toNumber(row[colM2CONST - 1]);
    if (m2 !== null) {
      g.minM2 = (g.minM2 === null) ? m2 : Math.min(g.minM2, m2);
      g.maxM2 = (g.maxM2 === null) ? m2 : Math.max(g.maxM2, m2);
    }

    const ref = String(row[colREF - 1] ?? "").trim();
    if (ref && !g.refs.includes(ref)) g.refs.push(ref);

    if (colCOORD) {
      const c = String(row[colCOORD - 1] ?? "").trim();
      if (c && !g.coords.includes(c)) g.coords.push(c);
    }
  });

  /************ Construcción de salida en orden ************/
  const out = [];
  const emittedGroup = new Set();
  let itemNum = 1;

  for (const row of picked) {
    const k = keyOf(row);
    const canGroup = k && (counts.get(k) || 0) >= 2;

    if (canGroup) {
      if (emittedGroup.has(k)) continue;
      emittedGroup.add(k);

      const g = groups.get(k);
      const base = g.baseRow;

      const minVal = g.minM2 !== null ? Math.round(g.minM2) : "";
      const maxVal = g.maxM2 !== null ? Math.round(g.maxM2) : "";

      // Suggested: manual salvo si no hay opción o min=max
      let sugVal = "";
      if (minVal !== "" && maxVal !== "" && minVal === maxVal) {
        sugVal = minVal;
      } else if (minVal === "" && maxVal === "") {
        const baseM2 = toNumber(base[colM2CONST - 1]);
        sugVal = baseM2 !== null ? Math.round(baseM2) : "";
      }

      const refsJoined = g.refs.length ? g.refs.join(", ") : (base[colREF - 1] ?? "");
      const coordsJoined = g.coords.length ? g.coords.join("\n\n") : "";

      out.push([
        itemNum++,
        refsJoined,
        base[colESTADO - 1],
        base[colSUBZONA - 1],
        sugVal,
        minVal,
        maxVal,
        base[colASKING - 1],
        base[colDISP - 1],
        "",
        coordsJoined,
        base[colPARQUE - 1]
      ]);
    } else {
      const m2 = toNumber(row[colM2CONST - 1]);
      const m2Round = (m2 !== null) ? Math.round(m2) : "";

      const coordSingle = colCOORD ? String(row[colCOORD - 1] ?? "").trim() : "";

      out.push([
        itemNum++,
        row[colREF - 1],
        row[colESTADO - 1],
        row[colSUBZONA - 1],
        m2Round,
        m2Round,
        m2Round,
        row[colASKING - 1],
        row[colDISP - 1],
        "",
        coordSingle,
        row[colPARQUE - 1]
      ]);
    }
  }

  /************ Render ************/
  const startRowTable = lastDataRow + 4;
  const numColsTable = FINAL_HEADERS.length;
  const numRowsTable = 1 + out.length;
  const totalContentRows = numRowsTable + 2; // +2 leyendas

  const clearStartRow = Math.max(1, startRowTable - PAD_ROWS);
  const clearStartCol = Math.max(1, START_COL_TABLE - PAD_LEFT);

  const blockRows = totalContentRows + (PAD_ROWS * 2);
  const blockCols = numColsTable + PAD_LEFT + PAD_RIGHT;

  const neededEndRow = clearStartRow + blockRows - 1;
  const maxRows = sheet.getMaxRows();
  if (neededEndRow > maxRows) sheet.insertRowsAfter(maxRows, neededEndRow - maxRows);

  const blockRange = sheet.getRange(clearStartRow, clearStartCol, blockRows, blockCols);
  blockRange.breakApart();
  blockRange.clear();
  blockRange.setBackground("#ffffff");
  blockRange.setBorder(false, false, false, false, false, false);

  // Encabezados
  const headerRange = sheet.getRange(startRowTable, START_COL_TABLE, 1, numColsTable);
  headerRange.setValues([FINAL_HEADERS]);
  headerRange.setFontWeight("bold");
  headerRange.setBackground(HEADER_BG);
  headerRange.setWrap(true);
  headerRange.setHorizontalAlignment("center");
  headerRange.setVerticalAlignment("middle");

  // Datos
  if (out.length) {
    sheet.getRange(startRowTable + 1, START_COL_TABLE, out.length, numColsTable).setValues(out);
  }

  // Bordes y alineación
  const tableRange = sheet.getRange(startRowTable, START_COL_TABLE, Math.max(1, numRowsTable), numColsTable);
  tableRange.setBorder(true, true, true, true, true, true);
  tableRange.setHorizontalAlignment("center");
  tableRange.setVerticalAlignment("middle");

  // Wrap overflow
  tableRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW);
  headerRange.setWrap(true);

  // Separador sin bordes / sin formato
  const rowsForBorders = Math.max(1, numRowsTable);
  const sepAbsCol = START_COL_TABLE + SEP_INDEX - 1;
  const sepColRange = sheet.getRange(startRowTable, sepAbsCol, rowsForBorders, 1);
  sepColRange.clearFormat();
  sepColRange.setBorder(false, false, false, false, false, false);

  // Reforzar bordes a ambos lados del separador
  const dispAbsCol = START_COL_TABLE + DISP_INDEX - 1;
  const coordAbsCol = START_COL_TABLE + COORD_INDEX - 1;
  sheet.getRange(startRowTable, dispAbsCol, rowsForBorders, 1).setBorder(null, null, null, true, null, null);
  sheet.getRange(startRowTable, coordAbsCol, rowsForBorders, 1).setBorder(null, true, null, null, null, null);

  // Formato numérico m2 y azul solo en suggested (sin encabezado)
  if (out.length) {
    const sugAbsCol = START_COL_TABLE + SUG_INDEX - 1;
    const minAbsCol = START_COL_TABLE + MIN_INDEX - 1;
    const maxAbsCol = START_COL_TABLE + MAX_INDEX - 1;

    sheet.getRange(startRowTable + 1, sugAbsCol, out.length, 1).setNumberFormat("#,##0");
    sheet.getRange(startRowTable + 1, minAbsCol, out.length, 1).setNumberFormat("#,##0");
    sheet.getRange(startRowTable + 1, maxAbsCol, out.length, 1).setNumberFormat("#,##0");

    sheet.getRange(startRowTable + 1, sugAbsCol, out.length, 1).setFontColor("#1a73e8");
  }

  // ✅ Merge notes only up to the last column BEFORE the separator
  const NOTE_COLS = SEP_INDEX - 1;

  // Footnote
  const footnoteRow = startRowTable + numRowsTable;
  sheet.getRange(footnoteRow, START_COL_TABLE, 1, numColsTable).breakApart();
  const footnoteRange = sheet.getRange(footnoteRow, START_COL_TABLE, 1, NOTE_COLS);
  footnoteRange.merge();
  footnoteRange.setValue(FOOTNOTE_TEXT);
  footnoteRange.setBackground("#ffffff");
  footnoteRange.setFontStyle("italic");
  footnoteRange.setHorizontalAlignment("left");
  footnoteRange.setVerticalAlignment("middle");
  footnoteRange.setWrap(true);

  // Flex note
  const flexRow = footnoteRow + 1;
  sheet.getRange(flexRow, START_COL_TABLE, 1, numColsTable).breakApart();
  const flexRange = sheet.getRange(flexRow, START_COL_TABLE, 1, NOTE_COLS);
  flexRange.merge();
  flexRange.setValue(FLEX_TEXT);
  flexRange.setBackground("#ffffff");
  flexRange.setFontStyle("italic");
  flexRange.setHorizontalAlignment("left");
  flexRange.setVerticalAlignment("middle");
  flexRange.setWrap(true);

  // Tamaños
  sheet.setColumnWidths(START_COL_TABLE, numColsTable, 110);
  if (PAD_LEFT > 0) sheet.setColumnWidths(START_COL_TABLE - PAD_LEFT, PAD_LEFT, 110);
  if (PAD_RIGHT > 0) sheet.setColumnWidths(START_COL_TABLE + numColsTable, PAD_RIGHT, 110);

  // Auto-ajuste SOLO Sub-zone
  const subZonaAbsCol = START_COL_TABLE + SUBZONA_OUT_INDEX - 1;
  sheet.autoResizeColumn(subZonaAbsCol);

  sheet.setRowHeights(startRowTable, totalContentRows, 22);
  sheet.autoResizeRows(startRowTable, totalContentRows);

  sheet.setActiveRange(headerRange);
}
