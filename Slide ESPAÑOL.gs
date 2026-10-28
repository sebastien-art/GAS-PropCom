/***************************************************************
 * 1) CONFIGURACIÓN GLOBAL
 ***************************************************************/
const TEMPLATE_PRESENTATION_ID = "1Hgyml-zwmQLsicHcNW8yg38Bf6ecN9AM7ts8kYzmrNE";

const SLIDE_NUM_COVER = 1;

/***************************************************************
 * 2) FUNCIÓN PRINCIPAL
 ***************************************************************/
function CrearEnSlides() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const sheetName = sheet.getName();

  try {
    const empresa = promptOrThrow_(ui, "Crear en Slides", "Pregunta 1: Empresa (CLIENTE)?");
    const estadoZona = promptOrThrow_(ui, "Crear en Slides", "Pregunta 2: Estado y Zona?");
    const rangoSup = promptOrThrow_(ui, "Crear en Slides", "Pregunta 3: Rango de superficie?");
    const fecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");

    // 1. Crear o buscar carpeta para guardar la presentación
    const sheetFolder = getOrCreateSheetFolderInRoot_(sheetName);

    // 2. Definir nombre y limpiar archivos anteriores con el mismo nombre
    const presName = buildPresentationName_(empresa, sheetName, fecha);
    trashFilesByName_(sheetFolder, presName);

    // 3. Copiar la plantilla e instanciar Slides
    const copyFile = DriveApp.getFileById(TEMPLATE_PRESENTATION_ID).makeCopy(presName, sheetFolder);
    const presId = copyFile.getId();
    const pres = SlidesApp.openById(presId);
    const slides = pres.getSlides();

    // 4. Llenar los datos de la Portada
    ES_insertCoverInfo_FINAL(slides[SLIDE_NUM_COVER - 1], pres, { empresa, estadoZona, rangoSup, fecha });

    // 5. Insertar la tabla "Partida" como tabla nativa de Slides (con hipervínculos reales a las fichas).
    // Si tiene más de FILAS_POR_SLIDE_TABLA filas, se reparte en varios slides
    // (ver insertarTablaPaginada_ en "TablaNativaSlides.js").
    const tableMeta = findFinalTableMeta_ES(sheet);
    const tabla = leerTablaConLinksPorNombre_(sheet, tableMeta.headerRow, tableMeta.colStart, 20, HEADERS_TABLA_SLIDE_ES, "REF");
    let paginasTabla = [];

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
    let tableBox = tableBoxBase;
    let leyenda = [];
    if (tabla.filas.length && advertirSiTablaGrande_(ui, tabla.filas.length)) {
      // Leyenda ("*"/"**") que ya existe debajo de la tabla en el Sheet — se
      // había pasado por alto al construir la tabla nativa (ver el comentario
      // de leerLeyendaTabla_ en "TablaNativaSlides.js"). Se lee ANTES de armar
      // la tabla para poder reservarle su espacio desde el principio (ver
      // boxConEspacioLeyenda_) y que no se pise con la última fila.
      leyenda = leerLeyendaTabla_(sheet, tableMeta.headerRow, tableMeta.colStart, tabla.filas.length);
      tableBox = boxConEspacioLeyenda_(tableBoxBase, leyenda.length > 0);

      paginasTabla = insertarTablaPaginada_(slides[5], pres, tabla.headers, tabla.filas, tabla.refColOffset, tableBox, "AUTO_TABLE_NATIVE");

      // Va en TODOS los slides de tabla generados (a pedido del usuario:
      // antes solo iba en el último), justo debajo del recuadro reservado.
      paginasTabla.forEach(pagina => {
        agregarLeyendaTabla_(pagina.slide, pres, leyenda, tableBox.top + tableBox.height, pagina.bottomParaLeyenda);
      });
    }

    // 6. Insertar el mapa de ubicaciones en el slide "Ubicaciones de las naves disponibles".
    //
    // NOTA — agregado el 2026-09-15 a pedido explícito del usuario, fuera del alcance
    // original del requerimiento "Mejora visual en Google Maps API" (que solo pedía
    // cambiar el pin exacto por un círculo de área). Hasta este cambio, CrearEnSlides()
    // nunca insertaba el mapa en ningún slide: ese slide 7 quedaba vacío y solo la
    // versión inglés/chino (Slides_ENG_CHI(), en "Slide ENG/CHI.js") lo hacía. Se agrega
    // aquí para que ambos idiomas queden consistentes.
    //
    // Se reutilizan tal cual las funciones getMapFilesForSheetStrict_INT() e
    // insertarMapaSeccionado_(), ya definidas y probadas en "Slide ENG/CHI.js", en vez de
    // duplicar la lógica de búsqueda/inserción de la imagen — el prefijo "_INT" es solo
    // el nombre histórico del archivo donde viven, no significa que sean exclusivas de
    // esa versión; ambos archivos ya comparten helpers de esta forma (getOrCreateSheetFolderInRoot_,
    // trashFilesByName_, promptOrThrow_, etc.).
    //
    // Requiere que ya se haya corrido GenerarMapaCliente() sobre esta misma pestaña
    // (si no, getMapFilesForSheetStrict_INT() lanza un error claro indicando qué falta,
    // y el catch de más abajo lo muestra como "Error Crítico" — igual que ya sucede en
    // Slides_ENG_CHI() cuando falta el mapa). Si GenerarMapaCliente() seccionó el mapa
    // en varias capturas (mapa único demasiado apiñado, ver
    // UMBRAL_MANCHA_CONTIGUA en "00_Config.js"), insertarMapaSeccionado_() las
    // reparte de a 2 lado a lado, duplicando este slide si hace falta una 2da página.
    const mapFiles = getMapFilesForSheetStrict_INT(sheetName);
    insertarMapaSeccionado_(pres, slides[6], mapFiles, "AUTO_MAP_IMG");

    pres.saveAndClose();

    // 7. Ajustar el ancho de columnas (Partida/REF angostas, Zona/Sub Zona
    // anchas) y el alto real de filas (encabezado más alto que las filas de
    // datos, para que no se estire solo y pise la leyenda) de la tabla en
    // CADA slide de tabla generado. Debe hacerse DESPUÉS de guardar — ver
    // ajustarAnchosColumna_()/ajustarAltosFila_() en "TablaNativaSlides.js"
    // para el detalle de por qué.
    // refEsLargo: si algún REF real supera UMBRAL_REF_LARGO_CHARS (ver
    // "TablaNativaSlides.js"), la columna REF se ensancha un poco y las de
    // Superficie ceden ese ancho — se calcula UNA vez para toda la tabla
    // (no por página) para que todas las páginas queden parejas.
    const refEsLargo = calcularRefEsLargo_(tabla.filas, tabla.refColOffset);
    paginasTabla.forEach(pagina => {
      ajustarAnchosColumna_(presId, pagina.objectId, tabla.headers, tableBox.width, refEsLargo);
      ajustarAltosFila_(presId, pagina.objectId, pagina.numFilasDatos, pagina.rowH, pagina.headerRowH);
    });

    // 8. Guardar ID y escribir links en la hoja
    PropertiesService.getDocumentProperties().setProperty("LAST_SLIDES_ID", presId);
    ES_writeSlidesLink_(sheet, sheet.getLastRow() + 2, 4, copyFile.getUrl(), sheetFolder.getUrl());

    ss.toast("¡Presentación generada! ✅", "Éxito");

  } catch (e) {
    ui.alert("Error Crítico: " + e.toString());
  }
}

/***************************************************************
 * 3) LÓGICA DE LA PORTADA
 ***************************************************************/
function ES_insertCoverInfo_FINAL(slide, pres, info) {
  ["AUTO_COVER_CLIENT", "AUTO_COVER_SUB"].forEach(d => {
    slide.getPageElements().forEach(pe => {
      try { if(pe.getDescription() === d) pe.remove(); } catch(e) {}
    });
  });

  const x = pres.getPageWidth() * 0.43;
  const w = pres.getPageWidth() * 0.52;
  const y = pres.getPageHeight() * 0.30;

  ES_setTxt(
    slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x, y + 255, w, 105).setDescription("AUTO_COVER_CLIENT"),
    String(info.empresa).toUpperCase(), 
    62, 
    true, 
    "#1a73e8"
  );

  ES_setTxt(
    slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, x, y + 345, w, 120).setDescription("AUTO_COVER_SUB"),
    `${info.estadoZona}\n${info.rangoSup}\n${info.fecha}`, 
    36, 
    false, 
    "#111111"
  );
}

function ES_setTxt(s, t, sz, b, c) {
  const tr = s.getText();
  tr.setText(t);
  tr.getTextStyle().setFontFamily("Arial").setFontSize(sz).setBold(b).setForegroundColor(c);
  tr.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
}

/***************************************************************
 * 4) HELPERS
 ***************************************************************/
function getOrCreateSheetFolderInRoot_(name) {
  const n = name.replace(/[\/\\:*?"<>|]/g, " ");
  const it = DriveApp.getFoldersByName(n);
  return it.hasNext() ? it.next() : DriveApp.createFolder(n);
}
/*
function trashFilesByName_(folder, name) { 
  const it = folder.getFilesByName(name); 
  while (it.hasNext()) it.next().setTrashed(true); 
}
*/

function buildPresentationName_(e, s, f) { 
  return `Propuesta - ${e} - ${s} - ${f}`; 
}

function findFinalTableMeta_ES(sheet) {
  const finder = sheet.createTextFinder("Partida").matchCase(false).matchEntireCell(true);
  const matches = finder.findAll();
  if (!matches.length) throw new Error('No encontré la tabla (encabezado "Partida"). Corre primero "Generar Tabla".');
  const best = matches[matches.length - 1];
  return { headerRow: best.getRow(), colStart: best.getColumn() };
}

function promptOrThrow_(ui, t, m) {
  const r = ui.prompt(t, m, ui.ButtonSet.OK_CANCEL); 
  if (r.getSelectedButton() !== ui.Button.OK) throw new Error("Cancelado"); 
  return r.getResponseText().trim(); 
}

function ES_writeSlidesLink_(sheet, row, col, url, folder) {
  sheet.getRange(row, col, 2, 1).setValues([["Slides:"],["Carpeta:"]]).setFontWeight("bold");
  sheet.getRange(row, col + 1).setFormula(`=HYPERLINK("${url}","Abrir Slides")`);
  sheet.getRange(row + 1, col + 1).setFormula(`=HYPERLINK("${folder}","Abrir carpeta")`);
}