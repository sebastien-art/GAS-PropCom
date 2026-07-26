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
    DESARROLLADOR: ["Desarrollador"], PARQUE: ["Parque"], COORD: ["Coordenadas"],
    INTERMEDIARIO: ["Intermediario", "Broker", "Agent"],
    UBICACION: ["Ubicación", "Ubicacion"]
  };

  const FINAL_HEADERS = [
    "Partida", "REF", "Zona Principal", "Sub Zona",
    "Superficie sugerida (m2)", "Superficie mínima disponible (m2) **","Superficie máxima disponible (m2) **",
    "Precio por m2 *", "Mantenimiento / m2", "Disponibilidad",
    "", 
    "Coordenadas", "Parque", "Desarrollador", "Intermediario", "Ubicación"
  ];

  const REF_IDX = 2;
  const ZONA_IDX = 3;
  const SUBZONA_IDX = 4;
  const MIN_IDX = 6;
  const MAX_IDX = 7;
  const ASKING_IDX = 8;
  const MANT_IDX = 9;
  const DISP_IDX = 10;
  const SEP_IDX = 11;
  const COORD_IDX = 12;
  const PARQUE_IDX = 13;
  const DEV_IDX = 14;
  const BROKER_IDX = 15;
  const UBIC_IDX = 16;

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
    coord: getCol(SRC_HEADERS.COORD), intermediario: getCol(SRC_HEADERS.INTERMEDIARIO),
    ubic: getCol(SRC_HEADERS.UBICACION)
  };

  const scanTo = sheet.getLastRow();
  const rangeData = sheet.getRange(HEADER_ROW + 1, 1, scanTo - HEADER_ROW, lastColSource);
  const dataValues = rangeData.getValues();
  const dataRichText = rangeData.getRichTextValues();

  const picked = [];
  for (let r = 0; r < dataValues.length; r++) {
    if (String(dataValues[r][c.ref - 1]).trim() !== "") {
      let richText = dataRichText[r][c.ficha - 1];
      let url = richText ? (richText.getLinkUrl() || "") : "";
      picked.push({ row: dataValues[r], url: url, originalIndex: r });
    }
  }

  if (!picked.length) return;

  const groups = new Map();
  picked.forEach((p, idx) => {
    const devVal = String(p.row[c.dev-1] || "").trim();
    const parqVal = String(p.row[c.parque-1] || "").trim();
    let k = (!devVal && !parqVal) ? "UNIQUE_" + idx : stripAccents(devVal + "||" + parqVal).toLowerCase();

    if (!groups.has(k)) groups.set(k, { items: [], minM2: null, maxM2: null, isAgrupable: (!!devVal || !!parqVal) });
    const g = groups.get(k); 
    g.items.push(p);
    let val = parseFloat(String(p.row[c.m2-1]).replace(/[^\d.]/g, ""));
    if (!isNaN(val)) {
      g.minM2 = (g.minM2 === null) ? val : Math.min(g.minM2, val);
      g.maxM2 = (g.maxM2 === null) ? val : Math.max(g.maxM2, val);
    }
  });

  const tableData = [];
  const emittedKeys = new Set();
  let partida = 1;

  picked.forEach((p, idx) => {
    const devVal = String(p.row[c.dev-1] || "").trim();
    const parqVal = String(p.row[c.parque-1] || "").trim();
    let k = (!devVal && !parqVal) ? "UNIQUE_" + idx : stripAccents(devVal + "||" + parqVal).toLowerCase();
    const g = groups.get(k);

    if (g.isAgrupable && g.items.length >= 2) {
      if (emittedKeys.has(k)) return;
      emittedKeys.add(k);
      const b = g.items[0].row;
      const uniqueCoords = [...new Set(g.items.map(i => String(i.row[c.coord-1]).trim()).filter(x => x))].join(", ");
      const uniqueUbic = [...new Set(g.items.map(i => c.ubic ? String(i.row[c.ubic-1]).trim() : "").filter(x => x))].join(", ");
      tableData.push({
        partida: partida++, isGroup: true, items: g.items,
        values: [
          b[c.zona-1], b[c.subzona-1],
          (g.minM2 === g.maxM2 ? Math.round(g.minM2) : ""), Math.round(g.minM2), Math.round(g.maxM2),
          b[c.asking-1], b[c.mant-1], b[c.disp-1], "", 
          uniqueCoords, b[c.parque-1], b[c.dev-1], b[c.intermediario-1], uniqueUbic
        ]
      });
    } else {
      const m2 = Math.round(parseFloat(String(p.row[c.m2-1]).replace(/[^\d.]/g, ""))) || "";
      const ubicVal = c.ubic ? String(p.row[c.ubic-1]).trim() : "";
      tableData.push({
        partida: partida++, isGroup: false, item: p,
        values: [
          p.row[c.zona-1], p.row[c.subzona-1],
          m2, m2, m2, p.row[c.asking-1], p.row[c.mant-1], p.row[c.disp-1], "", 
          p.row[c.coord-1], p.row[c.parque-1], p.row[c.dev-1], p.row[c.intermediario-1], ubicVal
        ]
      });
    }
  });

  /************ RENDERIZADO ************/
  const startRow = scanTo + 4;
  sheet.getRange(startRow - 1, START_COL_TABLE - 1, tableData.length + 50, FINAL_HEADERS.length + 5).breakApart().clear().setBackground("#ffffff");

  const headerRange = sheet.getRange(startRow, START_COL_TABLE, 1, FINAL_HEADERS.length);
  headerRange.setValues([FINAL_HEADERS])
       .setFontWeight("bold").setBackground(HEADER_BG).setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(true);

  const finalHeadersRow = FINAL_HEADERS.map(h => stripAccents(h).toLowerCase());
  const sugColOffset = finalHeadersRow.indexOf(stripAccents("Superficie sugerida (m2)").toLowerCase());
  const colSuperficieSugerida = (sugColOffset !== -1) ? START_COL_TABLE + sugColOffset : null;

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

  const tableRange = sheet.getRange(startRow, START_COL_TABLE, tableData.length + 1, FINAL_HEADERS.length);
  tableRange.setBorder(true, true, true, true, true, true).setVerticalAlignment("middle").setHorizontalAlignment("center");
  
  if (colSuperficieSugerida) {
    sheet.getRange(startRow + 1, colSuperficieSugerida, tableData.length, 1).setFontColor("#1155cc").setFontWeight("bold");
  }

  const cols110 = [6, MIN_IDX, MAX_IDX, ASKING_IDX, MANT_IDX, DISP_IDX, COORD_IDX, PARQUE_IDX, BROKER_IDX, UBIC_IDX];
  cols110.forEach(idx => { sheet.setColumnWidth(START_COL_TABLE + idx - 1, 110); });
  [ZONA_IDX, SUBZONA_IDX].forEach(idx => { sheet.setColumnWidth(START_COL_TABLE + idx - 1, 150); });
  sheet.setColumnWidth(START_COL_TABLE, 62); 
  sheet.setColumnWidth(START_COL_TABLE + 1, 42); 

  [REF_IDX, COORD_IDX, PARQUE_IDX, DEV_IDX, BROKER_IDX, UBIC_IDX].forEach(idx => {
    sheet.getRange(startRow + 1, START_COL_TABLE + idx - 1, tableData.length, 1).setHorizontalAlignment("left");
  });

  sheet.getRange(startRow + 1, START_COL_TABLE + 1, tableData.length, 1).setFontLine("underline").setFontColor("#1155cc");

  const sepRange = sheet.getRange(startRow, START_COL_TABLE + SEP_IDX - 1, tableData.length + 1, 1);
  sepRange.clearFormat().setBackground(null).setBorder(null, false, null, false, false, false);
  sheet.getRange(startRow, START_COL_TABLE + DISP_IDX - 1, tableData.length + 1, 1).setBorder(null, null, null, true, null, null);
  sheet.getRange(startRow, START_COL_TABLE + COORD_IDX - 1, tableData.length + 1, 1).setBorder(null, true, null, null, null, null);
  
  const numFormats = {
    "Superficie sugerida (m2)":            "#,##0",
    "Superficie mínima disponible (m2) **":  "#,##0",
    "Superficie máxima disponible (m2) **":  "#,##0",
    "Precio por m2 *":                 "#,##0.00",
    "Mantenimiento / m2":                  "#,##0.00",
  };
  Object.entries(numFormats).forEach(([header, fmt]) => {
    const offset = FINAL_HEADERS.indexOf(header);
    if (offset !== -1) {
      sheet.getRange(startRow + 1, START_COL_TABLE + offset, tableData.length, 1).setNumberFormat(fmt);
    }
  });

  const footnoteRow = startRow + tableData.length + 2;
  sheet.getRange(footnoteRow, START_COL_TABLE, 1, SEP_IDX - 1)
    .merge().setValue("* Precios y disponibilidades sujetos a cambios sin previo aviso.")
    .setFontSize(10).setFontStyle("italic").setHorizontalAlignment("left").setWrap(true);
  
  sheet.getRange(footnoteRow + 1, START_COL_TABLE, 1, SEP_IDX - 1)
    .merge().setValue("** Área flexible: podemos entregar el número de m² según su requerimiento, dentro del rango mostrado.")
    .setFontSize(10).setFontStyle("italic").setHorizontalAlignment("left").setWrap(true);
}