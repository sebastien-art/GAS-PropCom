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
    const tabla = leerTablaConLinksPorNombre_(sheet, tableMeta.headerRow, tableMeta.colStart, 15, HEADERS_TABLA_SLIDE_ENG_CHI, "REF / 参考编号");

    if (!tabla.filas.length) throw new Error('No se encontraron datos en la Tabla.');
    if (!advertirSiTablaGrande_(ui, tabla.filas.length)) throw new Error('Cancelado.');

    const mapFiles = getMapFilesForSheetStrict_INT(sheetName);

    const presName = `Proposal - ${empresa} - ${sheetName} - ${fecha}`;
    trashFilesByName_(sheetFolder, presName);
    
    const TEMPLATE_ID = "1xA4FVsxwUxrH9DPhwkfR8vYB5lD4lPYyuuSFi8H3UYY";
    const copyFile = DriveApp.getFileById(TEMPLATE_ID).makeCopy(presName, sheetFolder);
    const presId = copyFile.getId();
    
    const pres = SlidesApp.openById(presId);
    const slides = pres.getSlides();

    INT_insertCoverInfo_FINAL(slides[0], pres, { empresa, estadoZona, rangoSup, fecha });

    // Si la tabla tiene más de FILAS_POR_SLIDE_TABLA filas, se reparte en varios
    // slides (ver insertarTablaPaginada_ en "TablaNativaSlides.js").
    //
    // AJUSTE 10 (ver CAMBIO-1-TABLA-NATIVA-SLIDES.txt): con FILAS_POR_SLIDE_TABLA
    // subido a 15, el recuadro de siempre se quedaba corto — se sube el
    // título real del slide (no solo el recuadro) para ganarle espacio a la
    // tabla, y se recorta un poco el margen hacia el pie de página. Hay que
    // mover el título ANTES de insertarTablaPaginada_() (que duplica este
    // slide para las páginas siguientes), así el título ya movido se copia
    // solo a cada página nueva.
    moverTituloArribaParaTabla_(slides[5], DESPLAZAR_TITULO_PT);
    const tableBoxBase = {
      left: 70,
      top: 165 - DESPLAZAR_TITULO_PT,
      width: pres.getPageWidth() - 140,
      height: (pres.getPageHeight() - 260) + DESPLAZAR_TITULO_PT + RECORTAR_MARGEN_INFERIOR_PT
    };

    // Leyenda que ya existe debajo de la tabla en el Sheet — se lee ANTES de
    // armar la tabla para poder reservarle su espacio desde el principio (ver
    // boxConEspacioLeyenda_ en "TablaNativaSlides.js") y que no se pise con la
    // última fila.
    const leyenda = leerLeyendaTabla_(sheet, tableMeta.headerRow, tableMeta.colStart, tabla.filas.length);
    const tableBox = boxConEspacioLeyenda_(tableBoxBase, leyenda.length > 0);

    const paginasTabla = insertarTablaPaginada_(slides[5], pres, tabla.headers, tabla.filas, tabla.refColOffset, tableBox, "AUTO_TABLE_NATIVE");

    // Va en TODOS los slides de tabla generados (a pedido del usuario:
    // antes solo iba en el último), justo debajo del recuadro reservado.
    paginasTabla.forEach(pagina => {
      agregarLeyendaTabla_(pagina.slide, pres, leyenda, tableBox.top + tableBox.height, pagina.bottomParaLeyenda);
    });

    insertarMapaSeccionado_(pres, slides[6], mapFiles, "AUTO_MAP_IMG");

    pres.saveAndClose();

    // Ajustar el ancho de columnas (Partida/REF angostas, Zona/Sub Zona anchas) y
    // el alto real de filas (encabezado más alto que las filas de datos, para que
    // no se estire solo y pise la leyenda) de la tabla en CADA slide de tabla
    // generado. Debe hacerse DESPUÉS de guardar — ver ajustarAnchosColumna_()/
    // ajustarAltosFila_() en "TablaNativaSlides.js" para el detalle de por qué.
    // refEsLargo: si algún REF real supera UMBRAL_REF_LARGO_CHARS, la columna
    // REF se ensancha un poco y las de Superficie/Area ceden ese ancho — se
    // calcula UNA vez para toda la tabla (no por página) para que todas las
    // páginas queden parejas.
    const refEsLargo = calcularRefEsLargo_(tabla.filas, tabla.refColOffset);
    paginasTabla.forEach(pagina => {
      ajustarAnchosColumna_(presId, pagina.objectId, tabla.headers, tableBox.width, refEsLargo);
      ajustarAltosFila_(presId, pagina.objectId, pagina.numFilasDatos, pagina.rowH, pagina.headerRowH);
    });

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

// Busca el/los mapa(s) "definitivo(s)" que se insertan en las Slides (ver
// MAPA_DEFINITIVO_PREFIX en "00_Config.js"). El usuario probó las 3
// alternativas de mapa (círculos, zonas por municipio, calor) y primero
// eligió el de calor (AJUSTE 6), pero volvió a círculos/abanico el mismo
// día al ver que el degradado de calor tapaba los números de "Partida" en
// propiedades cercanas (AJUSTE 7) — ver CAMBIO-2-CIRCULOS-MAPA-UBICACION.txt.
//
// Devuelve SIEMPRE un arreglo de archivos: 1 elemento en el caso normal
// (mapa único), o 2-4 cuando GenerarMapaCliente() sección el mapa por
// quedar demasiado apiñado (ver UMBRAL_MANCHA_CONTIGUA en "00_Config.js").
// AJUSTE 31: GenerarMapaCliente() puede seccionar en 2 MITADES (si alcanzan)
// o en 4 cuadrantes (si no) — se buscan los archivos de CUALQUIERA de las 2
// formas (ver TODAS_LAS_ETIQUETAS_SECCION en "00_Config.js"), sin asumir
// cuál se usó. Quien llama debe pasar el resultado a
// insertarMapaSeccionado_(), que maneja los 3 casos (1, 2 o 4 archivos).
function getMapFilesForSheetStrict_INT(name) {
  const root = DriveApp.getFolderById(OUTPUT_ROOT_FOLDER_ID);
  const folderIt = root.getFoldersByName(MAPS_FOLDER_NAME);
  if (!folderIt.hasNext()) throw new Error("Falta carpeta: " + MAPS_FOLDER_NAME);
  const folder = folderIt.next();

  const fileNameUnico = MAPA_DEFINITIVO_PREFIX + name + ".png";
  const fileItUnico = folder.getFilesByName(fileNameUnico);
  if (fileItUnico.hasNext()) return [fileItUnico.next()];

  const archivos = [];
  TODAS_LAS_ETIQUETAS_SECCION.forEach(et => {
    const fileName = MAPA_DEFINITIVO_PREFIX + name + "_" + et + ".png";
    const fileIt = folder.getFilesByName(fileName);
    if (fileIt.hasNext()) archivos.push(fileIt.next());
  });
  if (archivos.length) return archivos;

  throw new Error("Falta \"" + fileNameUnico + "\" (o sus secciones) en Drive. Corre primero \"2a. 🔶 Generar Mapa Comercial\" del menú para esta pestaña.");
}

// Inserta el/los mapa(s) en el slide de "Ubicaciones de las naves
// disponibles". Con 1 solo archivo se comporta exactamente como antes
// (imagen centrada ocupando todo el recuadro reservado). Con 2 a 4
// archivos (mapa seccionado, ver getMapFilesForSheetStrict_INT()), se
// colocan de a 2 LADO A LADO por slide: los primeros 2 en slideMapa, y si
// hay una 3ra/4ta sección, se DUPLICA slideMapa (mismo mecanismo que
// insertarTablaPaginada_() usa para las páginas extra de la tabla,
// Slide.duplicate() inserta la copia justo después del original) para
// colocarlas ahí.
function insertarMapaSeccionado_(pres, slideMapa, files, desc) {
  const box = { left: 70, top: 165, width: pres.getPageWidth() - 140, height: pres.getPageHeight() - 260 };

  if (files.length <= 1) {
    if (files.length === 1) {
      INT_executeSafeInsert(slideMapa, pres, files[0].getBlob().setContentType("image/png"), desc);
    }
    return;
  }

  const GAP_PT = 10;
  const halfW = (box.width - GAP_PT) / 2;
  const boxIzq = { left: box.left, top: box.top, width: halfW, height: box.height };
  const boxDer = { left: box.left + halfW + GAP_PT, top: box.top, width: halfW, height: box.height };

  insertarImagenEnCaja_INT(slideMapa, files[0].getBlob().setContentType("image/png"), boxIzq, desc + "_1");
  if (files[1]) insertarImagenEnCaja_INT(slideMapa, files[1].getBlob().setContentType("image/png"), boxDer, desc + "_2");

  if (files[2]) {
    const slideMapa2 = slideMapa.duplicate();
    // Duplicar el slide copia también las imágenes ya puestas en el
    // primero — se quitan antes de insertar las de esta página.
    [desc + "_1", desc + "_2"].forEach(d => {
      slideMapa2.getPageElements().forEach(pe => { try { if (pe.getDescription() === d) pe.remove(); } catch (e) {} });
    });
    if (files[3]) {
      insertarImagenEnCaja_INT(slideMapa2, files[2].getBlob().setContentType("image/png"), boxIzq, desc + "_1");
      insertarImagenEnCaja_INT(slideMapa2, files[3].getBlob().setContentType("image/png"), boxDer, desc + "_2");
    } else {
      // Una sola sección en este 2do slide (3 secciones en total): se
      // centra con el mismo tamaño que las demás en vez de dejarla
      // pegada a la izquierda con la mitad derecha vacía.
      const boxCentro = { left: box.left + (box.width - halfW) / 2, top: box.top, width: halfW, height: box.height };
      insertarImagenEnCaja_INT(slideMapa2, files[2].getBlob().setContentType("image/png"), boxCentro, desc + "_1");
    }
  }
}

// Igual que INT_executeSafeInsert pero con una caja arbitraria (no siempre
// el recuadro completo del slide) — usado para acomodar 2 mapas lado a
// lado en insertarMapaSeccionado_().
function insertarImagenEnCaja_INT(slide, blob, box, desc) {
  slide.getPageElements().forEach(pe => {
    try { if (pe.getDescription() === desc) pe.remove(); } catch (e) {}
  });
  const img = slide.insertImage(blob).setDescription(desc);
  const ratio = Math.min(box.width / img.getWidth(), box.height / img.getHeight());
  img.setWidth(img.getWidth() * ratio).setHeight(img.getHeight() * ratio);
  img.setLeft(box.left + (box.width - img.getWidth()) / 2).setTop(box.top + (box.height - img.getHeight()) / 2);
}

function findFinalTableMeta_INT(sheet) {
  const finder = sheet.createTextFinder("Item / 项目").matchCase(false);
  const matches = finder.findAll();
  if (!matches.length) throw new Error('No encontré la tabla (Item / 项目).');
  let best = matches[matches.length - 1]; 
  return { headerRow: best.getRow(), colStart: best.getColumn() };
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