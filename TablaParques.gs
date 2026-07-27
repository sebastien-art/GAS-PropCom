function TablaParques() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  sheet.setHiddenGridlines(true);

  const HEADER_ROW = 1;
  const START_COL_TABLE = 4; // Columna Partida (D)
  const HEADER_BG = "#9fc5e8";

  const SRC_HEADERS = {
    ENVIAR: ["ENVIAR"], FICHA: ["Ficha", "OK"], REF: ["REF"],
    OPERACION: ["Operación", "Operacio", "Operació"],
    ESTADO: ["Estado"], ZONA: ["Zona Principal"], SUBZONA: ["Sub Zona"],
    M2CONST: ["M2 de construcción", "m2 de construccion"], ASKING: ["Asking price /m2"],
    MANT: ["Mantenimiento / m2"], DISP: ["Disponibilidad"],
    DESARROLLADOR: ["Desarrollador"], PARQUE: ["Parque"], COORD: ["Coordenadas"],
    INTERMEDIARIO: ["Intermediario", "Broker", "Agent"],
    UBICACION: ["Ubicación", "Ubicacion"]
  };

  const COL_FICHA = START_COL_TABLE - 2; // Columna B (2)

  // 1. "Operación" agregada después de "Ubicación"
  const FINAL_HEADERS = [
    "Partida", "REF", "Zona Principal", "Sub Zona",
    "Superficie sugerida (m2)", "Superficie mínima disponible (m2) **","Superficie máxima disponible (m2) **",
    "Precio por m2 *", "Disponibilidad",
    "", "Precio total", "Mantenimiento / m2", "Estado",
    "Coordenadas", "Parque", "Desarrollador", "Intermediario", "Ubicación", "Operación"
  ];

  const REF_IDX = 2;
  const ZONA_IDX = 3;
  const SUBZONA_IDX = 4;
  const MIN_IDX = 6;
  const MAX_IDX = 7;
  const ASKING_IDX = 8;
  const DISP_IDX = 9;
  const SEP_IDX = 10;
  const PRECIO_TOTAL_IDX = 11;
  const MANT_IDX = 12;
  const ESTADO_IDX = 13;
  const COORD_IDX = 14;
  const PARQUE_IDX = 15;
  const DEV_IDX = 16;
  const BROKER_IDX = 17;
  const UBIC_IDX = 18;
  const OP_IDX = 19; // Índice asignado a Operación

  const stripAccents = (s) => String(s ?? "").normalize("NFD").replace(/[\u0300-\u036f]/g, "");
  const makeKeys = (s) => {
    let t = stripAccents(String(s ?? "").trim().toLowerCase()).replace(/\s+/g, " ");
    return [t, t.replace(/\s+/g, ""), t.replace(/\//g, "").replace(/\s+/g, "")];
  };

  const parsePrecioMoneda = (val) => {
    let str = String(val ?? "").trim();
    if (!str) return { num: 0, currency: "" };
    let isUSD = str.toUpperCase().includes("USD");
    let isMXN = str.toUpperCase().includes("MXN");
    let clean = str.replace(/[^\d.]/g, "");
    let num = parseFloat(clean) || 0;
    let currency = isUSD ? "USD" : (isMXN ? "MXN" : "");
    return { num: num, currency: currency };
  };

  const formatPrecioTotal = (sugM2, rawPrice) => {
    let m2 = parseFloat(String(sugM2).replace(/[^\d.]/g, "")) || 0;
    let p = parsePrecioMoneda(rawPrice);
    if (!m2 || !p.num) return "";
    let total = Math.round(m2 * p.num);
    let totalStr = total.toLocaleString("en-US");
    if (p.currency === "USD") return "USD " + totalStr;
    if (p.currency === "MXN") return totalStr + " MXN";
    return totalStr;
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
    ref: getCol(SRC_HEADERS.REF), operacion: getCol(SRC_HEADERS.OPERACION),
    estado: getCol(SRC_HEADERS.ESTADO),
    zona: getCol(SRC_HEADERS.ZONA), subzona: getCol(SRC_HEADERS.SUBZONA),
    m2: getCol(SRC_HEADERS.M2CONST), asking: getCol(SRC_HEADERS.ASKING),
    mant: getCol(SRC_HEADERS.MANT), disp: getCol(SRC_HEADERS.DISP),
    dev: getCol(SRC_HEADERS.DESARROLLADOR), parque: getCol(SRC_HEADERS.PARQUE),
    coord: getCol(SRC_HEADERS.COORD), intermediario: getCol(SRC_HEADERS.INTERMEDIARIO),
    ubic: getCol(SRC_HEADERS.UBICACION)
  };

  const lastRowAll = sheet.getLastRow();
  let scanTo = HEADER_ROW;
  
  for (let r = HEADER_ROW + 1; r <= lastRowAll; r++) {
    const valColFicha = String(sheet.getRange(r, COL_FICHA).getValue()).trim();
    if (valColFicha.toLowerCase() === "ficha") {
      break;
    }
    scanTo = r;
  }

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
      const opVal = c.operacion ? String(b[c.operacion-1] || "").trim() : "";
      const estadoVal = c.estado ? String(b[c.estado-1] || "").trim() : "";
      
      const sugM2 = (g.minM2 === g.maxM2 ? Math.round(g.minM2) : "");
      const precioTotal = formatPrecioTotal(sugM2, b[c.asking-1]);

      tableData.push({
        partida: partida++, isGroup: true, items: g.items,
        values: [
          b[c.zona-1], b[c.subzona-1],
          sugM2, Math.round(g.minM2), Math.round(g.maxM2),
          b[c.asking-1], b[c.disp-1], "", 
          precioTotal, b[c.mant-1], estadoVal,
          uniqueCoords, b[c.parque-1], b[c.dev-1], b[c.intermediario-1], uniqueUbic, opVal
        ]
      });
    } else {
      const m2 = Math.round(parseFloat(String(p.row[c.m2-1]).replace(/[^\d.]/g, ""))) || "";
      const precioTotal = formatPrecioTotal(m2, p.row[c.asking-1]);
      const ubicVal = c.ubic ? String(p.row[c.ubic-1]).trim() : "";
      const opVal = c.operacion ? String(p.row[c.operacion-1] || "").trim() : "";
      const estadoVal = c.estado ? String(p.row[c.estado-1] || "").trim() : "";

      tableData.push({
        partida: partida++, isGroup: false, item: p,
        values: [
          p.row[c.zona-1], p.row[c.subzona-1],
          m2, m2, m2, p.row[c.asking-1], p.row[c.disp-1], "", 
          precioTotal, p.row[c.mant-1], estadoVal,
          p.row[c.coord-1], p.row[c.parque-1], p.row[c.dev-1], p.row[c.intermediario-1], ubicVal, opVal
        ]
      });
    }
  });

  /************ RENDERIZADO ************/
  const startRow = lastRowAll + 4;

  // Encabezado Ficha
  sheet.getRange(startRow, COL_FICHA).setValue("Ficha")
       .setFontWeight("bold").setBackground(HEADER_BG).setHorizontalAlignment("center").setVerticalAlignment("middle").setFontLine("none").setFontColor("#000000");

  // Tabla Principal
  const headerRange = sheet.getRange(startRow, START_COL_TABLE, 1, FINAL_HEADERS.length);
  headerRange.setValues([FINAL_HEADERS])
       .setFontWeight("bold").setBackground(HEADER_BG).setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(true);

  const finalHeadersRow = FINAL_HEADERS.map(h => stripAccents(h).toLowerCase());
  const sugColOffset = finalHeadersRow.indexOf(stripAccents("Superficie sugerida (m2)").toLowerCase());
  const colSuperficieSugerida = (sugColOffset !== -1) ? START_COL_TABLE + sugColOffset : null;

  tableData.forEach((rowObj, i) => {
    const rowIdx = startRow + 1 + i;
    sheet.getRange(rowIdx, START_COL_TABLE).setValue(rowObj.partida);
    
    const refCell = sheet.getRange(rowIdx, START_COL_TABLE + 1); // Columna REF (Col E / 5)
    const fichaCell = sheet.getRange(rowIdx, COL_FICHA);        // Columna FICHA (Col B / 2)

    refCell.clearDataValidations();
    fichaCell.clearDataValidations();
    refCell.setValue("");
    fichaCell.setValue(""); 

    if (rowObj.isGroup) {
      let options = ["Seleccionar..."];
      rowObj.items.forEach(item => {
        let refStr = String(item.row[c.ref-1]).trim();
        let m2Str = String(item.row[c.m2-1]).trim();
        let estadoFicha = item.url ? "FICHA OK" : "SIN FICHA";
        
        let label = refStr + (m2Str ? " (" + m2Str + " m2)" : "") + " - " + estadoFicha;
        options.push(label);
      });
      options.push("CREAR");

      let rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(options, true)
        .setAllowInvalid(true)
        .build();

      let hasLinks = rowObj.items.some(item => !!item.url);

      if (hasLinks) {
        let fullText = rowObj.items.map(item => String(item.row[c.ref-1]).trim()).join(", ");
        let richValue = SpreadsheetApp.newRichTextValue().setText(fullText);
        
        let cursor = 0;
        rowObj.items.forEach(item => {
          let txt = String(item.row[c.ref-1]).trim();
          if (item.url) {
            richValue.setLinkUrl(cursor, cursor + txt.length, item.url);
          }
          cursor += txt.length + 2;
        });

        fichaCell.setFontLine("underline").setFontColor("#1155cc");
        fichaCell.setRichTextValue(richValue.build());

        refCell.setDataValidation(rule);
        refCell.setValue("Seleccionar...");
        refCell.setFontLine("none").setFontColor("#000000");

      } else {
        fichaCell.setDataValidation(rule);
        fichaCell.setValue("Seleccionar...");
        fichaCell.setFontLine("none").setFontColor("#000000");

        refCell.setFontLine("none").setFontColor("#000000");
      }

    } else {
      let refTxt = String(rowObj.item.row[c.ref-1]).trim();
      let richValue = SpreadsheetApp.newRichTextValue().setText(refTxt);
      if (rowObj.item.url) {
        richValue.setLinkUrl(0, refTxt.length, rowObj.item.url);
        refCell.setFontLine("underline").setFontColor("#1155cc");
      } else {
        refCell.setFontLine("none").setFontColor("#000000");
      }
      refCell.setRichTextValue(richValue.build());
    }

    sheet.getRange(rowIdx, START_COL_TABLE + 2, 1, rowObj.values.length).setValues([rowObj.values]);
  });

  const tableRange = sheet.getRange(startRow, START_COL_TABLE, tableData.length + 1, FINAL_HEADERS.length);
  tableRange.setBorder(true, true, true, true, true, true).setVerticalAlignment("middle").setHorizontalAlignment("center");

  const fichaRange = sheet.getRange(startRow, COL_FICHA, tableData.length + 1, 1);
  fichaRange.setBorder(true, true, true, true, true, true)
            .setVerticalAlignment("middle")
            .setHorizontalAlignment("left")
            .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  
  if (colSuperficieSugerida) {
    sheet.getRange(startRow + 1, colSuperficieSugerida, tableData.length, 1).setFontColor("#1155cc").setFontWeight("bold");
  }

  // ANCHOS DE COLUMNAS
  sheet.setColumnWidth(COL_FICHA, 160);
  sheet.setColumnWidth(START_COL_TABLE, 62);          // Partida (Col D)
  sheet.setColumnWidth(START_COL_TABLE + 1, 70);       // REF (Col E)

  const cols110 = [MIN_IDX, MAX_IDX, ASKING_IDX, MANT_IDX, ESTADO_IDX, DISP_IDX, PRECIO_TOTAL_IDX, COORD_IDX, PARQUE_IDX, BROKER_IDX, UBIC_IDX, OP_IDX];
  cols110.forEach(idx => { sheet.setColumnWidth(START_COL_TABLE + idx - 1, 110); });
  [ZONA_IDX, SUBZONA_IDX].forEach(idx => { sheet.setColumnWidth(START_COL_TABLE + idx - 1, 150); });

  [REF_IDX, ESTADO_IDX, COORD_IDX, PARQUE_IDX, DEV_IDX, BROKER_IDX, UBIC_IDX, OP_IDX].forEach(idx => {
    sheet.getRange(startRow + 1, START_COL_TABLE + idx - 1, tableData.length, 1).setHorizontalAlignment("left");
  });

  const sepRange = sheet.getRange(startRow, START_COL_TABLE + SEP_IDX - 1, tableData.length + 1, 1);
  sepRange.clearFormat().setBackground(null).setBorder(null, false, null, false, false, false);
  sheet.getRange(startRow, START_COL_TABLE + DISP_IDX - 1, tableData.length + 1, 1).setBorder(null, null, null, true, null, null);
  sheet.getRange(startRow, START_COL_TABLE + PRECIO_TOTAL_IDX - 1, tableData.length + 1, 1).setBorder(true, true, true, true, true, true);
  
  const numFormats = {
    "Superficie sugerida (m2)":            "#,##0",
    "Superficie mínima disponible (m2) **": "#,##0",
    "Superficie máxima disponible (m2) **": "#,##0"
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

/*************************************************************
 * RECALCULO Y MANEJO DE DESPLEGABLES EN INTERACCIÓN (onEdit)
 *************************************************************/

function onEdit(e) {
  if (!e) return;
  const range = e.range;
  const sheet = range.getSheet();
  
  const editedRow = range.getRow();
  const editedCol = range.getColumn();
  
  // 1. Encontrar la fila de encabezado de la tabla específica
  const lastCol = sheet.getLastColumn();
  let headerRowIdx = -1;

  for (let r = editedRow; r >= 1; r--) {
    let cellVal = String(sheet.getRange(r, 4).getValue()).trim().toLowerCase(); // Columna Partida (D)
    if (cellVal === "partida") {
      headerRowIdx = r;
      break;
    }
  }

  // Si no pertenece a una tabla generada, ignorar
  if (headerRowIdx === -1 || editedRow <= headerRowIdx) return;

  // 2. Mapear columnas exactas escaneando directamente el rango de encabezados
  const headerValues = sheet.getRange(headerRowIdx, 1, 1, lastCol).getValues()[0];
  
  let colFicha = 2; // Columna B
  let colRef = -1;
  let colSugM2 = -1;
  let colAsking = -1;
  let colTotal = -1;

  headerValues.forEach((h, idx) => {
    let text = String(h).trim().toLowerCase();
    if (text === "ref") colRef = idx + 1;
    if (text.includes("superficie sugerida")) colSugM2 = idx + 1;
    if (text.includes("precio por m2") || text.includes("asking")) colAsking = idx + 1;
    if (text === "precio total") colTotal = idx + 1;
  });

  if (colSugM2 === -1 || colAsking === -1 || colTotal === -1) return;

  const selectedValue = String(range.getValue()).trim();

  // Función auxiliar para buscar si la REF seleccionada tiene un link en la columna FICHA
  const aplicarRefConLink = (refTargetCell, cleanRefStr) => {
    let fichaCell = sheet.getRange(editedRow, colFicha);
    let fichaRichText = fichaCell.getRichTextValue();
    let targetUrl = "";

    if (fichaRichText) {
      let runs = fichaRichText.getRuns();
      for (let run of runs) {
        if (run.getText().trim() === cleanRefStr && run.getLinkUrl()) {
          targetUrl = run.getLinkUrl();
          break;
        }
      }
    }

    if (targetUrl) {
      let richVal = SpreadsheetApp.newRichTextValue()
        .setText(cleanRefStr)
        .setLinkUrl(0, cleanRefStr.length, targetUrl)
        .build();
      refTargetCell.setRichTextValue(richVal);
      refTargetCell.setFontLine("underline").setFontColor("#1155cc");
    } else {
      refTargetCell.setValue(cleanRefStr);
      refTargetCell.setFontLine("none").setFontColor("#000000");
    }
  };

  // A) INTERACCIÓN EN COLUMNA REF
  if (editedCol === colRef && selectedValue && selectedValue !== "Seleccionar...") {
    let refMatch = selectedValue.match(/^([^\s(]+)/);
    let m2Match = selectedValue.match(/\(([\d,.]+)\s*m2\)/i);

    if (refMatch && refMatch[1]) {
      let cleanRef = refMatch[1];
      aplicarRefConLink(range, cleanRef);
    }

    if (m2Match && m2Match[1]) {
      let numM2 = parseFloat(m2Match[1].replace(/,/g, "")) || 0;
      if (numM2 > 0) {
        sheet.getRange(editedRow, colSugM2).setValue(numM2);
      }
    }
    // Quitar icono del desplegable
    range.clearDataValidations();
  }

  // B) INTERACCIÓN EN COLUMNA FICHA
  if (editedCol === colFicha && selectedValue && selectedValue !== "Seleccionar...") {
    let refMatch = selectedValue.match(/^([^\s(]+)/);
    let m2Match = selectedValue.match(/\(([\d,.]+)\s*m2\)/i);

    if (refMatch && refMatch[1] && colRef !== -1) {
      let refCell = sheet.getRange(editedRow, colRef);
      let cleanRef = refMatch[1];
      aplicarRefConLink(refCell, cleanRef);
      refCell.clearDataValidations();
    }

    if (m2Match && m2Match[1]) {
      let numM2 = parseFloat(m2Match[1].replace(/,/g, "")) || 0;
      if (numM2 > 0) {
        sheet.getRange(editedRow, colSugM2).setValue(numM2);
      }
    }
    // Quitar icono del desplegable
    range.clearDataValidations();
  }

  // C) CÁLCULO DE PRECIO TOTAL (Al cambiar Superficie Sugerida o por selección)
  let rawSugM2 = sheet.getRange(editedRow, colSugM2).getValue();
  let numSugM2 = parseFloat(String(rawSugM2).replace(/[^\d.]/g, "")) || 0;
  
  let rawAsking = sheet.getRange(editedRow, colAsking).getValue();
  let strAsking = String(rawAsking ?? "").trim();

  let cellTotal = sheet.getRange(editedRow, colTotal);

  if (!numSugM2 || !strAsking) {
    cellTotal.setValue("");
    return;
  }

  let isUSD = strAsking.toUpperCase().includes("USD");
  let isMXN = strAsking.toUpperCase().includes("MXN");
  let cleanAsking = parseFloat(strAsking.replace(/[^\d.]/g, "")) || 0;

  if (cleanAsking > 0) {
    let totalNum = Math.round(numSugM2 * cleanAsking);
    let totalFormatted = totalNum.toLocaleString("en-US");

    if (isUSD) totalFormatted = "USD " + totalFormatted;
    else if (isMXN) totalFormatted = totalFormatted + " MXN";

    cellTotal.setValue(totalFormatted);
  } else {
    cellTotal.setValue("");
  }
}