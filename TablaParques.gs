function TablaParques() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  sheet.setHiddenGridlines(true);

  const HEADER_ROW = 1;
  const START_COL_TABLE = 4; 
  const HEADER_BG = "#9fc5e8";

  const SRC_HEADERS = {
    ENVIAR: ["ENVIAR", "Operación"], FICHA: ["Ficha", "OK"], REF: ["REF"],
    ESTADO: ["Estado"], ZONA: ["Zona Principal"], SUBZONA: ["Sub Zona"],
    M2CONST: ["M2 de construcción", "m2 de construccion"], ASKING: ["Asking price /m2"],
    MANT: ["Mantenimiento / m2"], DISP: ["Disponibilidad"],
    DESARROLLADOR: ["Desarrollador"], PARQUE: ["Parque"], COORD: ["Coordenadas"]
  };

  const FINAL_HEADERS = [
    "Partida", "REF", "Estado", "Zona Principal", "Sub Zona",
    "Superficie sugerida (m2)", "Superficie mínima rentable (m2) **", "Superficie máxima rentable (m2) **",
    "Asking price / m2 *", "Mantenimiento / m2", "Disponibilidad",
    "", // SEPARADOR
    "Coordenadas", "Parque", "Desarrollador" 
  ];

  const REF_IDX = 2, ESTADO_IDX = 3, ZONA_IDX = 4, SUBZONA_IDX = 5, SUG_IDX = 6;
  const MIN_IDX = 7, MAX_IDX = 8, ASKING_IDX = 9, MANT_IDX = 10, DISP_IDX = 11;
  const SEP_IDX = 12, COORD_IDX = 13, PARQUE_IDX = 14, DEV_IDX = 15;

  /************ PROCESAMIENTO ************/
  const stripAccents = (s) => String(s ?? "").normalize("NFD").replace(/[\u0300-\u036f]/g, "");
  const makeKeys = (s) => {
    let t = stripAccents(String(s ?? "").trim().toLowerCase()).replace(/\s+/g, " ");
    return [t, t.replace(/\s+/g, ""), t.replace(/\//g, "").replace(/\s+/g, "")];
  };

  const lastColSource = sheet.getLastColumn();
  const headerValues = sheet.getRange(HEADER_ROW, 1, 1, lastColSource).getValues()[0];
  const headerMap = new Map();
  headerValues.forEach((h, i) => { makeKeys(h).forEach(k => { if (k) headerMap.set(k, i + 1); }); });

  const getCol = (names) => {
    for (const n of names) { for (const k of makeKeys(n)) { if (headerMap.has(k)) return headerMap.get(k); } }
    return null;
  };

  const c = {
    enviar: getCol(SRC_HEADERS.ENVIAR), ficha: getCol(SRC_HEADERS.FICHA),
    ref: getCol(SRC_HEADERS.REF), estado: getCol(SRC_HEADERS.ESTADO),
    zona: getCol(SRC_HEADERS.ZONA), subzona: getCol(SRC_HEADERS.SUBZONA),
    m2: getCol(SRC_HEADERS.M2CONST), asking: getCol(SRC_HEADERS.ASKING),
    mant: getCol(SRC_HEADERS.MANT), disp: getCol(SRC_HEADERS.DISP),
    dev: getCol(SRC_HEADERS.DESARROLLADOR), parque: getCol(SRC_HEADERS.PARQUE),
    coord: getCol(SRC_HEADERS.COORD)
  };

  const scanTo = sheet.getLastRow();
  const rangeData = sheet.getRange(HEADER_ROW + 1, 1, scanTo - HEADER_ROW, lastColSource);
  const dataValues = rangeData.getValues();
  const dataRichText = rangeData.getRichTextValues();

  const picked = [];
  for (let r = 0; r < dataValues.length; r++) {
    if (/^OK\b/.test(stripAccents(String(dataValues[r][c.enviar - 1])).trim().toUpperCase())) {
      let url = dataRichText[r][c.ficha - 1].getLinkUrl() || "";
      picked.push({ row: dataValues[r], url: url });
    }
  }

  if (!picked.length) return;

  const groups = new Map();
  picked.forEach(p => {
    const k = stripAccents(p.row[c.dev-1] + "||" + p.row[c.parque-1]).toLowerCase();
    if (!groups.has(k)) groups.set(k, { items: [], minM2: null, maxM2: null });
    const g = groups.get(k); g.items.push(p);
    let val = parseFloat(String(p.row[c.m2-1]).replace(/[^\d.]/g, ""));
    if (!isNaN(val)) {
      g.minM2 = (g.minM2 === null) ? val : Math.min(g.minM2, val);
      g.maxM2 = (g.maxM2 === null) ? val : Math.max(g.maxM2, val);
    }
  });

  const tableData = [];
  const emitted = new Set();
  let partida = 1;

  picked.forEach(p => {
    const k = stripAccents(p.row[c.dev-1] + "||" + p.row[c.parque-1]).toLowerCase();
    const g = groups.get(k);
    if (g.items.length >= 2) {
      if (emitted.has(k)) return;
      emitted.add(k);
      const b = g.items[0].row;
      const uniqueCoords = [...new Set(g.items.map(i => String(i.row[c.coord-1]).trim()).filter(x => x))].join(", ");
      tableData.push({
        partida: partida++, isGroup: true, items: g.items,
        values: [
          b[c.estado-1], b[c.zona-1], b[c.subzona-1],
          (g.minM2 === g.maxM2 ? Math.round(g.minM2) : ""), Math.round(g.minM2), Math.round(g.maxM2),
          b[c.asking-1], b[c.mant-1], b[c.disp-1], "", 
          uniqueCoords, b[c.parque-1], b[c.dev-1]
        ]
      });
    } else {
      const m2 = Math.round(parseFloat(String(p.row[c.m2-1]).replace(/[^\d.]/g, ""))) || "";
      tableData.push({
        partida: partida++, isGroup: false, item: p,
        values: [
          p.row[c.estado-1], p.row[c.zona-1], p.row[c.subzona-1],
          m2, m2, m2, p.row[c.asking-1], p.row[c.mant-1], p.row[c.disp-1], "", 
          p.row[c.coord-1], p.row[c.parque-1], p.row[c.dev-1]
        ]
      });
    }
  });

  /************ RENDERIZADO ************/
  const startRow = scanTo + 4;
  sheet.getRange(startRow - 1, START_COL_TABLE - 1, tableData.length + 30, FINAL_HEADERS.length + 5).breakApart().clear().setBackground("#ffffff");

  sheet.getRange(startRow, START_COL_TABLE, 1, FINAL_HEADERS.length).setValues([FINAL_HEADERS])
       .setFontWeight("bold").setBackground(HEADER_BG).setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(true);

  tableData.forEach((rowObj, i) => {
    const rowIdx = startRow + 1 + i;
    sheet.getRange(rowIdx, START_COL_TABLE).setValue(rowObj.partida);
    const refCell = sheet.getRange(rowIdx, START_COL_TABLE + 1);
    let richValue = SpreadsheetApp.newRichTextValue();
    if (rowObj.isGroup) {
      let fullText = rowObj.items.map(item => String(item.row[c.ref-1]).trim()).join(", ");
      richValue.setText(fullText);
      let cursor = 0;
      rowObj.items.forEach(item => {
        let txt = String(item.row[c.ref-1]).trim();
        if (item.url) richValue.setLinkUrl(cursor, cursor + txt.length, item.url);
        cursor += txt.length + 2;
      });
    } else {
      let txt = String(rowObj.item.row[c.ref-1]).trim();
      richValue.setText(txt);
      if (rowObj.item.url) richValue.setLinkUrl(0, txt.length, rowObj.item.url);
    }
    refCell.setRichTextValue(richValue.build());
    sheet.getRange(rowIdx, START_COL_TABLE + 2, 1, rowObj.values.length).setValues([rowObj.values]);
  });

  // BORDES GENERALES
  const tableRange = sheet.getRange(startRow, START_COL_TABLE, tableData.length + 1, FINAL_HEADERS.length);
  tableRange.setBorder(true, true, true, true, true, true).setVerticalAlignment("middle").setHorizontalAlignment("center");
  
  // ANCHOS SOLICITADOS
  // 110 px: Estado, Superficies, Asking, Mantenimiento, Disponibilidad, Coordenadas, Parque
  [ESTADO_IDX, SUG_IDX, MIN_IDX, MAX_IDX, ASKING_IDX, MANT_IDX, DISP_IDX, COORD_IDX, PARQUE_IDX].forEach(idx => {
    sheet.setColumnWidth(START_COL_TABLE + idx - 1, 110);
  });
  // 150 px: Zona Principal y Sub Zona
  [ZONA_IDX, SUBZONA_IDX].forEach(idx => { sheet.setColumnWidth(START_COL_TABLE + idx - 1, 150); });
  sheet.setColumnWidth(START_COL_TABLE, 62); // Partida
  sheet.setColumnWidth(START_COL_TABLE + 1, 42); // REF

  // ALINEACIONES IZQUIERDA
  [REF_IDX, COORD_IDX, PARQUE_IDX, DEV_IDX].forEach(idx => {
    sheet.getRange(startRow + 1, START_COL_TABLE + idx - 1, tableData.length, 1).setHorizontalAlignment("left");
  });

  // FORMATO REF
  sheet.getRange(startRow + 1, START_COL_TABLE + 1, tableData.length, 1).setFontLine("underline").setFontColor("#1155cc");

  // AJUSTE SEPARADOR (Espacio voluntario con bordes correctos)
  const sepRange = sheet.getRange(startRow, START_COL_TABLE + SEP_IDX - 1, tableData.length + 1, 1);
  sepRange.clearFormat().setBackground(null).setBorder(null, false, null, false, false, false);
  // Asegurar que Disponibilidad tenga borde derecho y Coordenadas borde izquierdo
  sheet.getRange(startRow, START_COL_TABLE + DISP_IDX - 1, tableData.length + 1, 1).setBorder(null, null, null, true, null, null);
  sheet.getRange(startRow, START_COL_TABLE + COORD_IDX - 1, tableData.length + 1, 1).setBorder(null, true, null, null, null, null);
  
  sheet.setRowHeights(startRow, tableData.length + 1, 21);
  sheet.getRange(startRow + 1, START_COL_TABLE + 5, tableData.length, 5).setNumberFormat("#,##0.00");
}