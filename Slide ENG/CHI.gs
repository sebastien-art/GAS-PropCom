/***************************************************************
 * SLIDES_ENG_CHI - VERSIÓN ULTRA COMPATIBLE
 ***************************************************************/

function Slides_ENG_CHI() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const sheetName = sheet.getName();

  try {
    const empresa = promptOrThrow_(ui, "Create Slides / 生成幻灯片", "Q1: Nombre CLIENTE");
    const estadoZona = promptOrThrow_(ui, "Create Slides / 生成幻灯片", "Q2: Zona y Estado");
    const rangoSup = promptOrThrow_(ui, "Create Slides / 生成幻灯片", "Q3: Superficie deseada");
    const fecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");

    const sheetFolder = getOrCreateSheetFolderInRoot_(sheetName);
    registrarPropuesta_(sheetName, sheetFolder.getId());
    
    const tableMeta = findFinalTableMeta_INT(sheet);
    const tableData = readFinalTableForSlides_INT(sheet, tableMeta);

    if (!tableData.values || tableData.values.length === 0) throw new Error('No se encontraron datos en la Tabla.');

    const mapFile = getMapFileForSheetStrict_INT(sheetName);
    const mapBlob = mapFile.getBlob().setContentType("image/png");

    const tableImgBlob = renderTablePngViaTempSlides_INT(tableData, tableMeta);

    const presName = `Proposal - ${empresa} - ${sheetName} - ${fecha}`;
    trashFilesByName_(sheetFolder, presName);
    
    const TEMPLATE_ID = "1xA4FVsxwUxrH9DPhwkfR8vYB5lD4lPYyuuSFi8H3UYY";
    const copyFile = DriveApp.getFileById(TEMPLATE_ID).makeCopy(presName, sheetFolder);
    const presId = copyFile.getId();
    
    const pres = SlidesApp.openById(presId);
    const slides = pres.getSlides();

    INT_insertCoverInfo_FINAL(slides[0], pres, { empresa, estadoZona, rangoSup, fecha });
    
    INT_executeSafeInsert(slides[5], pres, tableImgBlob, "AUTO_TABLE_IMG");
    
    INT_executeSafeInsert(slides[6], pres, mapBlob, "AUTO_MAP_IMG");

    pres.saveAndClose();

    PropertiesService.getDocumentProperties().setProperty("LAST_SLIDES_ID", presId);
    PropertiesService.getDocumentProperties().setProperty("LAST_SHEET_NAME", sheetName);
    INT_writeLinks_FINAL(sheet, sheet.getLastRow() + 2, 4, copyFile.getUrl(), sheetFolder.getUrl());
    
    ss.toast("✅ Slide Creado", "Exito");

  } catch (e) {
    ui.alert("Error Crítico: " + e.toString());
  }
}

/***************************************************************
 * MOTORES DE APOYO (CORREGIDOS)
 ***************************************************************/

function renderTablePngViaTempSlides_INT(tableData, tableMeta) {
  const tempPres = SlidesApp.create("__TEMP_TABLE__");
  const tempId = tempPres.getId();

  try {
    const slide = tempPres.getSlides()[0];
    const PW = tempPres.getPageWidth();
    const PH = tempPres.getPageHeight();

    const headers = tableData.headers;
    const rows    = tableData.values;
    const allRows = [headers, ...rows];

    const cols     = headers.length;
    const rowCount = allRows.length;
    const colW     = (PW - 40) / cols;
    const rowH     = Math.min(30, (PH - 40) / rowCount);
    const startX   = 20;
    const startY   = 20;

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
          tf.getTextStyle()
            .setFontFamily("Arial")
            .setFontSize(isHeader ? 9 : 8)
            .setBold(isHeader)
            .setForegroundColor(isHeader ? "#ffffff" : "#222222");
          tf.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

        } catch(cellErr) {
          Logger.log("Error en celda [" + rIdx + "," + cIdx + "]: " + cellErr.message);
        }
      });
    });

    tempPres.saveAndClose();

    const thumbnailUrl = "https://slides.googleapis.com/v1/presentations/" +
      tempId + "/pages/" + SlidesApp.openById(tempId).getSlides()[0].getObjectId() +
      "/thumbnail?thumbnailProperties.mimeType=PNG&thumbnailProperties.thumbnailSize=LARGE";

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

function columnToLetter_(col) {
  let letter = "";
  while (col > 0) {
    const mod = (col - 1) % 26;
    letter = String.fromCharCode(65 + mod) + letter;
    col = Math.floor((col - 1) / 26);
  }
  return letter;
}

function INT_executeSafeInsert(slide, pres, blob, desc) {
  try {
    slide.getPageElements().forEach(pe => {
      try { if (pe.getDescription() === desc) pe.remove(); } catch(e) {}
    });

    const box = { left: 70, top: 165, width: pres.getPageWidth() - 140, height: pres.getPageHeight() - 260 };
    
    const img = slide.insertImage(blob).setDescription(desc);
    
    const ratio = Math.min(box.width / img.getWidth(), box.height / img.getHeight());
    img.setWidth(img.getWidth() * ratio).setHeight(img.getHeight() * ratio);
    img.setLeft(box.left + (box.width - img.getWidth()) / 2).setTop(box.top + (box.height - img.getHeight()) / 2);
  } catch (e) {
    throw new Error(`Fallo al insertar ${desc}. Google Slides no acepta el formato generado.`);
  }
}

function getMapFileForSheetStrict_INT(name) {
  const fileName = MAP_FILE_PREFIX + name + ".png"; 
  const folderIt = DriveApp.getFoldersByName(MAPS_FOLDER_NAME);
  if (!folderIt.hasNext()) throw new Error("Falta carpeta: " + MAPS_FOLDER_NAME);
  const folder = folderIt.next();
  const fileIt = folder.getFilesByName(fileName);
  if (!fileIt.hasNext()) throw new Error("Falta \"" + fileName + "\" en Drive.");
  return fileIt.next();
}

function findFinalTableMeta_INT(sheet) {
  const finder = sheet.createTextFinder("Item / 项目").matchCase(false);
  const matches = finder.findAll();
  if (!matches.length) throw new Error('No encontré la tabla (Item / 项目).');
  let best = matches[matches.length - 1]; 
  return { headerRow: best.getRow(), colStart: best.getColumn() };
}

function readFinalTableForSlides_INT(sheet, meta) {
  const headers = sheet.getRange(meta.headerRow, meta.colStart, 1, 15).getDisplayValues()[0];
  let end = headers.findIndex(h => h.trim() === "");
  const finalHeaders = headers.slice(0, end > 0 ? end : headers.length);
  const values = [];
  for (let r = meta.headerRow + 1; r <= sheet.getLastRow(); r++) {
    const row = sheet.getRange(r, meta.colStart, 1, finalHeaders.length).getDisplayValues()[0];
    if (!/^\d+$/.test(row[0]) || !row[1]) break;
    values.push(row);
  }
  return { headers: finalHeaders, values };
}

function INT_insertCoverInfo_FINAL(slide, pres, info) {
  ["AUTO_COVER_CLIENT", "AUTO_COVER_SUB"].forEach(d => {
    slide.getPageElements().forEach(pe => { try { if(pe.getDescription() === d) pe.remove(); } catch(e){} });
  });
  const x = pres.getPageWidth() * 0.43, w = pres.getPageWidth() * 0.52, y = pres.getPageHeight() * 0.30;
  INT_setTxt(slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x, y + 205, w, 105).setDescription("AUTO_COVER_CLIENT"), String(info.empresa).toUpperCase(), 62, true, "#1a73e8");
  INT_setTxt(slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x, y + 345, w, 170).setDescription("AUTO_COVER_SUB"), `${info.estadoZona}\n${info.rangoSup}\n${info.fecha}`, 36, false, "#111111");
}

function INT_setTxt(s, t, sz, b, c) {
  const tr = s.getText(); tr.setText(t);
  tr.getTextStyle().setFontFamily("Arial").setFontSize(sz).setBold(b).setForegroundColor(c);
  tr.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
}

function INT_writeLinks_FINAL(sheet, row, col, url, folder) {
  sheet.getRange(row, col).setValue("Slides:").setFontWeight("bold");
  sheet.getRange(row, col + 1).setFormula(`=HYPERLINK("${url}","Open Slides")`);
  sheet.getRange(row + 1, col).setValue("Folder:").setFontWeight("bold");
  sheet.getRange(row + 1, col + 1).setFormula(`=HYPERLINK("${folder}","Open Folder")`);
}