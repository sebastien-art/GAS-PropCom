function Park_ENG_CHI() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  sheet.setHiddenGridlines(true);

  const HEADER_ROW = 1;
  const START_COL_TABLE = 4; // D
  const HEADER_BG = "#9fc5e8";

  const FOOTNOTE_TEXT =
    "* Price (shell) before negotiation and technical project + maintenance + VAT / *价格为谈判及技术方案之前的报价，另加维护费和增值税（VAT）。";

  const FLEX_TEXT =
    "** Flexible area: we can deliver the number of m² according to your requirement, within the range shown (from the minimum leasable area to the maximum leasable area). / ** 面积灵活：我们可根据您的需求在所示范围内交付相应的平方米数（从最小可租面积到最大可租面积）。";

  const SRC_HEADERS = {
    ENVIAR: ["ENVIAR", "Operación", "SEND"],
    FICHA: ["Ficha", "OK", "Link"],
    REF: ["REF", "参考编号"],
    ESTADO: ["Estado", "Status", "状态"],
    ZONA: ["Zona Principal", "Main Zone", "区域"],
    SUBZONA: ["Sub Zona", "Sub Zone"],
    M2CONST: ["M2 de construcción", "m2 de construccion", "Building Area", "建筑面积"],
    ASKING: ["Asking price /m2", "租金"],
    MANT: ["Mantenimiento / m2", "Maintenance", "物业"],
    DISP: ["Disponibilidad", "Availability"],
    DESARROLLADOR: ["Desarrollador", "Developer", "开发商"],
    PARQUE: ["Parque", "Park", "园区"],
    COORD: ["Coordenadas", "Coordinates", "坐标"],
    INTERMEDIARIO: ["Intermediario", "Broker", "Agent"] // Agregado
  };

  const FINAL_HEADERS = [
    "Item / 项目",
    "REF / 参考编号",
    "State / 州",
    "Main zone / 主区域",
    "Sub-zone / 子区域",
    "Suggested Area (m2) / 建议面积",
    "Min Area (m2) ** / 最小面积",
    "Max Area (m2) ** / 最大面积",
    "Asking price * / 要价*",
    "Maint. / 维护费",
    "Availability / 可用性",
    "", // separator column
    "Coordinates / 坐标",
    "Industrial Park / 工业园",
    "Developer / 开发商",
    "Broker / 中介" // Agregado
  ];

  const REF_IDX = 2, ESTADO_IDX = 3, ZONA_IDX = 4, SUBZONA_IDX = 5;
  const MIN_IDX = 7, MAX_IDX = 8, ASKING_IDX = 9, MANT_IDX = 10, DISP_IDX = 11;
  const SEP_IDX = 12, COORD_IDX = 13, PARQUE_IDX = 14, DEV_IDX = 15, BROKER_IDX = 16;

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
    coord: getCol(SRC_HEADERS.COORD), intermediario: getCol(SRC_HEADERS.INTERMEDIARIO)
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
      picked.push({ row: dataValues[r], url: url });
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
      tableData.push({
        partida: partida++, isGroup: true, items: g.items,
        values: [
          b[c.estado-1], b[c.zona-1], b[c.subzona-1],
          (g.minM2 === g.maxM2 ? Math.round(g.minM2) : ""), Math.round(g.minM2), Math.round(g.maxM2),
          b[c.asking-1], b[c.mant-1], b[c.disp-1], " ", 
          uniqueCoords, b[c.parque-1], b[c.dev-1], b[c.intermediario-1]
        ]
      });
    } else if (!g.isAgrupable || g.items.length < 2) {
      const m2 = Math.round(parseFloat(String(p.row[c.m2-1]).replace(/[^\d.]/g, ""))) || "";
      tableData.push({
        partida: partida++, isGroup: false, item: p,
        values: [
          p.row[c.estado-1], p.row[c.zona-1], p.row[c.subzona-1],
          m2, m2, m2, p.row[c.asking-1], p.row[c.mant-1], p.row[c.disp-1], " ", 
          p.row[c.coord-1], p.row[c.parque-1], p.row[c.dev-1], p.row[c.intermediario-1]
        ]
      });
    }
  });

  const startRow = scanTo + 4;
  sheet.getRange(startRow - 1, START_COL_TABLE - 1, tableData.length + 50, FINAL_HEADERS.length + 5).breakApart().clear().setBackground("#ffffff");

  const headerRange = sheet.getRange(startRow, START_COL_TABLE, 1, FINAL_HEADERS.length);
  headerRange.setValues([FINAL_HEADERS])
       .setFontWeight("bold").setBackground(HEADER_BG).setFontColor("#000000").setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(true);

  const finalHeadersRow = FINAL_HEADERS.map(h => stripAccents(h).toLowerCase());
  const sugColOffset = finalHeadersRow.indexOf(stripAccents("Suggested Area (m2) / 建议面积").toLowerCase());
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
    sheet.getRange(startRow + 1, colSuperficieSugerida, tableData.length, 1).setFontColor("#1a73e8").setFontWeight("bold");
  }

  const cols110 = [ESTADO_IDX, 6, MIN_IDX, MAX_IDX, ASKING_IDX, MANT_IDX, DISP_IDX, COORD_IDX, PARQUE_IDX, BROKER_IDX];
  cols110.forEach(idx => {
    sheet.setColumnWidth(START_COL_TABLE + idx - 1, 110);
  });
  [ZONA_IDX, SUBZONA_IDX].forEach(idx => { 
    sheet.setColumnWidth(START_COL_TABLE + idx - 1, 150); 
  });
  sheet.setColumnWidth(START_COL_TABLE, 62); 
  sheet.setColumnWidth(START_COL_TABLE + 1, 42); 

  [REF_IDX, COORD_IDX, PARQUE_IDX, DEV_IDX, BROKER_IDX].forEach(idx => {
    sheet.getRange(startRow + 1, START_COL_TABLE + idx - 1, tableData.length, 1).setHorizontalAlignment("left");
  });

  sheet.getRange(startRow + 1, START_COL_TABLE + 1, tableData.length, 1).setFontLine("underline").setFontColor("#1a73e8");

  const sepRange = sheet.getRange(startRow, START_COL_TABLE + SEP_IDX - 1, tableData.length + 1, 1);
  sepRange.clearFormat().setBackground(null).setBorder(null, false, null, false, false, false);
  sheet.getRange(startRow, START_COL_TABLE + DISP_IDX - 1, tableData.length + 1, 1).setBorder(null, null, null, true, null, null);
  sheet.getRange(startRow, START_COL_TABLE + COORD_IDX - 1, tableData.length + 1, 1).setBorder(null, true, null, null, null, null);
  
  sheet.getRange(startRow + 1, START_COL_TABLE + 5, tableData.length, 3).setNumberFormat("#,##0");
  sheet.getRange(startRow + 1, START_COL_TABLE + 8, tableData.length, 2).setNumberFormat("#,##0.00");

  const footnoteRow = startRow + tableData.length + 1;
  const footnoteRange = sheet.getRange(footnoteRow, START_COL_TABLE, 1, SEP_IDX - 1);
  footnoteRange.merge().setValue(FOOTNOTE_TEXT).setBackground("#ffffff").setFontStyle("italic").setHorizontalAlignment("left").setWrap(true);
  
  const flexRow = footnoteRow + 1;
  const flexRange = sheet.getRange(flexRow, START_COL_TABLE, 1, SEP_IDX - 1);
  flexRange.merge().setValue(FLEX_TEXT).setBackground("#ffffff").setFontStyle("italic").setHorizontalAlignment("left").setWrap(true);

  SpreadsheetApp.flush();
  sheet.setActiveRange(headerRange);
}