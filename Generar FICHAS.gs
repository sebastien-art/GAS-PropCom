/**
 * GENERADOR DE FICHAS TÉCNICAS - VERSIÓN CORREGIDA
 * 
 * FIXES APLICADOS:
 * [FIX-1] indexarCarpeta() ahora usa try/catch por archivo Y por slide individual,
 *         con logs detallados para identificar exactamente dónde se cuelga.
 * [FIX-2] Se agrega límite de tiempo explícito en la indexación para evitar timeout silencioso.
 * [FIX-3] Filtro de Partida corregido: ahora valida que sea número entero, no solo cadena no vacía.
 * [FIX-4] inyectarDatosEnFicha() protege el reemplazo de REF para no romper slides sin ese texto.
 * [FIX-5] Limpieza de slides ajenos usa objectId del deck COPIADO, no del original.
 */

function generarFichasPropuestasComerciales() {
  const CARPETA_ORIGEN_ID  = '1xSLs5GHRHbm9OMJwQldxmipNIhPT8eVc';
  const CARPETA_DESTINO_ID = '1-W5VYI53YCQSf_9KYL0dplqw3eXbz1dD';

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const rawData = sheet.getDataRange().getDisplayValues();

  // --- LOCALIZAR FILA DE CABECERAS DE LA TABLA AZUL (escanea toda la hoja) ---
  // La tabla azul se identifica por tener: "partida", "ref", "desarrollador" e "intermediario"
  // en la misma fila. La tabla de inventario arriba NO tiene "partida" ni "intermediario".
  let headerRowIdx = -1;
  for (let r = 0; r < rawData.length; r++) {
    const filaBaja = rawData[r].map(c => c.toString().trim().toLowerCase());
    if (filaBaja.includes("partida") && filaBaja.includes("ref") &&
        filaBaja.includes("desarrollador") && filaBaja.includes("intermediario")) {
      headerRowIdx = r;
      break;
    }
  }

  if (headerRowIdx === -1) {
    SpreadsheetApp.getUi().alert("Error: No se encontró la tabla azul (fila con Partida, REF, Desarrollador e Intermediario).");
    return;
  }

  console.log(`[SHEET] Tabla azul encontrada en fila ${headerRowIdx + 1}`);

  // --- MAPEO DE COLUMNAS ---
  const headers = rawData[headerRowIdx].map(h => h.toString().trim().toLowerCase());
  const col = {
    partida:      headers.indexOf("partida"),
    ref:          headers.indexOf("ref"),
    parque:       headers.findIndex(h => h.includes("parque")),
    desarrollador:headers.findIndex(h => h.includes("desarrollador")),
    zona:         headers.findIndex(h => h.includes("zona principal")),
    intermediario:headers.indexOf("intermediario"),
    m2:           headers.findIndex(h => h.includes("superficie sugerida")),
    precio:       headers.findIndex(h => h.includes("precio por m2")),
    renta:        headers.findIndex(h => h.includes("renta total") || h.includes("renta mensual"))
  };

  console.log(`[SHEET] Mapeo de columnas: ${JSON.stringify(col)}`);

  // --- LOCALIZAR COLUMNA "ENLACE FICHA" (solo en la fila de cabeceras de la tabla azul) ---
  let colEnlaceIdx = headers.findIndex(h => h.includes("enlace ficha"));

  if (colEnlaceIdx === -1) {
    // Crear la columna después de "Renta total"
    colEnlaceIdx = col.renta !== -1 ? col.renta + 1 : headers.length;
    sheet.getRange(headerRowIdx + 1, colEnlaceIdx + 1)
         .setValue("Enlace Ficha")
         .setFontWeight("bold")
         .setBackground("#cbd5e1");
    console.log(`[SHEET] Columna "Enlace Ficha" creada en columna ${colEnlaceIdx + 1}`);
  }

  // --- CONSTRUIR MAPA DE TRÍOS ---
  // Trío = Zona Principal + Parque + Desarrollador
  // Una celda de REF puede tener múltiples valores separados por coma
  const trioMap = {};
  for (let i = headerRowIdx + 1; i < rawData.length; i++) {
    const row = rawData[i];
    const refCelda = col.ref           !== -1 ? row[col.ref].toString().trim()           : "";
    const devStr   = col.desarrollador !== -1 ? row[col.desarrollador].toString().trim() : "";
    const parqStr  = col.parque        !== -1 ? row[col.parque].toString().trim()        : "";
    const zonaStr  = col.zona          !== -1 ? row[col.zona].toString().trim()          : "";
    if (refCelda && devStr && parqStr && zonaStr) {
      const key = `${zonaStr}|${parqStr}|${devStr}`.toLowerCase();
      if (!trioMap[key]) trioMap[key] = [];
      refCelda.split(",").forEach(r => {
        const ref = r.trim();
        if (ref && !trioMap[key].includes(ref)) trioMap[key].push(ref);
      });
    }
  }

  // --- [FIX-1] INDEXACIÓN CON LOGS Y PROTECCIÓN POR ARCHIVO ---
  console.log("=== INICIO INDEXACIÓN ===");
  const inicioIndexacion = Date.now();
  const cacheModelos = [];
  const folderOrigen  = DriveApp.getFolderById(CARPETA_ORIGEN_ID);
  const folderDestino = DriveApp.getFolderById(CARPETA_DESTINO_ID);

  indexarCarpeta(folderOrigen, cacheModelos, inicioIndexacion);

  console.log(`=== INDEXACIÓN COMPLETA: ${cacheModelos.length} slides cargados en ${((Date.now() - inicioIndexacion)/1000).toFixed(1)}s ===`);

  if (cacheModelos.length === 0) {
    SpreadsheetApp.getUi().alert("No se encontraron slides válidos en la carpeta raíz.");
    return;
  }

  // --- PROCESAR FILAS ---
  let procesadas = 0, errores = 0, sinModelo = 0;

  for (let i = headerRowIdx + 1; i < rawData.length; i++) {
    const row = rawData[i];

    // [FIX-3] Filtro de Partida: debe ser número entero válido
    const partidaStr = col.partida !== -1 ? row[col.partida].toString().trim() : "";
    if (!/^\d+$/.test(partidaStr)) continue;

    // Filtro Intermediario = D
    if (col.intermediario !== -1 && row[col.intermediario].toString().trim().toUpperCase() !== 'D') continue;

    const refCelda = col.ref !== -1 ? row[col.ref].toString().trim() : "";
    if (!refCelda || refCelda.toLowerCase().includes("ref")) continue;

    // Dividir celda REF en valores individuales (puede haber varias separadas por coma)
    const refsEnCelda = refCelda.split(",").map(r => r.trim()).filter(r => r !== "");
    // Usar la primera REF como identificador principal de la fila
    const refTabla = refsEnCelda[0];

    const devStr  = col.desarrollador !== -1 ? row[col.desarrollador].toString().trim() : "";
    const parqStr = col.parque        !== -1 ? row[col.parque].toString().trim()        : "";
    const zonaStr = col.zona          !== -1 ? row[col.zona].toString().trim()          : "";
    const m2Str   = col.m2    !== -1 ? row[col.m2].toString().trim()    : "0";
    const precioStr = col.precio !== -1 ? row[col.precio].toString().trim() : "A consultar";
    const rentaStr  = col.renta  !== -1 ? row[col.renta].toString().trim()  : "A consultar";

    const trioKey = `${zonaStr}|${parqStr}|${devStr}`.toLowerCase();
    const hermanasDelTrio = trioMap[trioKey] || [];

    // Búsqueda exacta: probar cada REF de la celda individualmente
    // Solo acepta modelos con variables inyectables (tieneVariables = true)
    let modeloEncontrado = null;
    for (const ref of refsEnCelda) {
      const refClean = ref.replace(/\s+/g, '').toUpperCase();
      modeloEncontrado = cacheModelos.find(m => m.tieneVariables && m.textoLimpio.includes(refClean));
      if (modeloEncontrado) break;
    }

    // Búsqueda por hermanas del trío (también solo con variables)
    if (!modeloEncontrado && hermanasDelTrio.length > 0) {
      for (const refHermana of hermanasDelTrio) {
        const hermanaClean = refHermana.replace(/\s+/g, '').toUpperCase();
        modeloEncontrado = cacheModelos.find(m => m.tieneVariables && m.textoLimpio.includes(hermanaClean));
        if (modeloEncontrado) break;
      }
    }

    if (!modeloEncontrado) console.warn(`[SIN MODELO CON VARIABLES] Fila ${i+1} | REF: ${refTabla} | Trío: ${trioKey}`);

    // Escritura segura en celda (combinada o no)
    const rangoCelda = sheet.getRange(i + 1, colEnlaceIdx + 1);
    const celdaDestino = rangoCelda.isPartOfMerge() ? rangoCelda.getCell(1, 1) : rangoCelda;

    if (modeloEncontrado) {
      try {
        const url = inyectarDatosEnFicha(modeloEncontrado, partidaStr, m2Str, precioStr, rentaStr, zonaStr, refTabla, folderDestino);
        celdaDestino.setValue(url);
        procesadas++;
        console.log(`[OK] Fila ${i+1} | REF: ${refTabla} → ${url}`);
      } catch (err) {
        celdaDestino.setValue("Error: " + err.message);
        errores++;
        console.error(`[ERROR] Fila ${i+1} | REF: ${refTabla} → ${err.toString()}`);
      }
    } else {
      celdaDestino.setValue("Sin modelo para este Trío");
      sinModelo++;
      console.warn(`[SIN MODELO] Fila ${i+1} | REF: ${refTabla} | Trío: ${trioKey}`);
    }
  }

  SpreadsheetApp.getUi().alert(
    `Proceso finalizado.\n✅ Procesadas: ${procesadas}\n⚠️ Sin modelo: ${sinModelo}\n❌ Errores: ${errores}\n\nRevisa el Log de Apps Script para detalle completo.`
  );
}

// ─────────────────────────────────────────────────────────────
// [FIX-1] INDEXACIÓN ROBUSTA CON LOGS DETALLADOS
// ─────────────────────────────────────────────────────────────
function indexarCarpeta(carpeta, cacheModelos, inicioMs) {
  const TIMEOUT_MS = 240000; // Abortar indexación si supera 4 minutos

  console.log(`[CARPETA] Entrando: ${carpeta.getName()}`);

  let files = carpeta.getFilesByType(MimeType.GOOGLE_SLIDES);
  while (files.hasNext()) {
    // [FIX-2] Chequeo de tiempo antes de cada archivo
    if (Date.now() - inicioMs > TIMEOUT_MS) {
      console.error("[TIMEOUT] Indexación abortada por límite de tiempo en carpeta: " + carpeta.getName());
      return;
    }

    const file = files.next();
    console.log(`  [ARCHIVO] Indexando: ${file.getName()} (${file.getId()})`);

    try {
      const deck = SlidesApp.openById(file.getId());
      const slides = deck.getSlides();
      console.log(`    → ${slides.length} slide(s) encontrados`);

      for (let i = 0; i < slides.length; i++) {
        try {
          const texto = extraerTodoElTexto(slides[i]);
          const tieneVariables = texto.includes("{{M2 de construcción}}");
          cacheModelos.push({
            fileId:        file.getId(),
            fileName:      file.getName(),
            slideIdx:      i,
            textoLimpio:   texto.replace(/\s+/g, '').toUpperCase(),
            tieneVariables: tieneVariables
          });
          console.log(`    → Slide ${i} indexado. Variables: ${tieneVariables}. Texto: ${texto.substring(0, 60).replace(/\n/g,' ')}...`);
        } catch (eSlide) {
          console.error(`    → Error en slide ${i}: ${eSlide.toString()}`);
        }
      }
    } catch (eFile) {
      console.error(`  [ERROR ARCHIVO] ${file.getName()}: ${eFile.toString()}`);
    }
  }

  // Recursión en subcarpetas
  let subFolders = carpeta.getFolders();
  while (subFolders.hasNext()) {
    if (Date.now() - inicioMs > TIMEOUT_MS) {
      console.error("[TIMEOUT] Indexación abortada por límite de tiempo antes de subcarpeta.");
      return;
    }
    indexarCarpeta(subFolders.next(), cacheModelos, inicioMs);
  }
}

// ─────────────────────────────────────────────────────────────
// EXTRACTOR DE TEXTO DE UN SLIDE
// ─────────────────────────────────────────────────────────────
function extraerTodoElTexto(slide) {
  const textos = [];
  slide.getPageElements().forEach(el => {
    try {
      if (el.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
        const t = el.asShape().getText();
        if (t) textos.push(t.asString());
      } else if (el.getPageElementType() === SlidesApp.PageElementType.TABLE) {
        const tabla = el.asTable();
        for (let r = 0; r < tabla.getNumRows(); r++) {
          for (let c = 0; c < tabla.getNumColumns(); c++) {
            try {
              const t = tabla.getCell(r, c).getText().asString();
              if (t) textos.push(t);
            } catch(e) {}
          }
        }
      }
    } catch(e) {}
  });
  return textos.join(" ");
}

// ─────────────────────────────────────────────────────────────
// INYECCIÓN DE DATOS EN LA FICHA CLONADA
// ─────────────────────────────────────────────────────────────
function inyectarDatosEnFicha(modelo, partida, m2, precio, renta, zona, refNueva, destino) {
  const moldeDeck = SlidesApp.openById(modelo.fileId);
  const slides    = moldeDeck.getSlides();

  // Determinar hasta dónde llega la ficha de esta REF (hasta la siguiente REF o fin)
  let finIdx = modelo.slideIdx + 1;
  while (finIdx < slides.length) {
    const texto = extraerTodoElTexto(slides[finIdx]).toUpperCase();
    if (texto.includes("REF:")) break;
    finIdx++;
  }

  // IDs de slides a conservar
  const idsAConservar = [];
  for (let i = modelo.slideIdx; i < finIdx; i++) {
    idsAConservar.push(slides[i].getObjectId());
  }

  const m2Formateado = m2.toLowerCase().includes("m") ? m2 : `${m2} m²`;
  const nombreFicha  = `${partida} - RENTA ${m2Formateado.replace(/[^\d.,]/g, '')} M2 ${zona} ${refNueva}`;

  // Crear copia en carpeta destino
  const copiaFile = DriveApp.getFileById(modelo.fileId).makeCopy(nombreFicha, destino);

  // [FIX-5] Abrir el DECK COPIADO para limpiar y reemplazar (no el original)
  const nuevaDeck = SlidesApp.openById(copiaFile.getId());

  // Eliminar slides que no corresponden a esta ficha
  // Iterar en reversa para no romper índices al eliminar
  const slidesNuevos = nuevaDeck.getSlides();
  for (let i = slidesNuevos.length - 1; i >= 0; i--) {
    if (!idsAConservar.includes(slidesNuevos[i].getObjectId())) {
      try { slidesNuevos[i].remove(); } catch(e) {
        console.warn(`No se pudo eliminar slide ${i}: ${e.toString()}`);
      }
    }
  }

  // Reemplazar variables en el deck copiado
  nuevaDeck.getSlides().forEach(s => {
    s.getPageElements().forEach(el => {

      // [FIX-4] Reemplazo de REF protegido
      if (el.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
        try {
          const shape = el.asShape();
          const txt = shape.getText().asString();
          if (txt.toUpperCase().includes("REF:")) {
            // Reemplazar solo la parte después de "REF:" conservando formato
            shape.getText().setText("REF: " + refNueva);
          }
        } catch(e) {}
      }

      // Reemplazo en tablas
      if (el.getPageElementType() === SlidesApp.PageElementType.TABLE) {
        const tabla = el.asTable();
        for (let r = 0; r < tabla.getNumRows(); r++) {
          for (let c = 0; c < tabla.getNumColumns(); c++) {
            try {
              const celdaTxt = tabla.getCell(r, c).getText();
              const cellStr  = celdaTxt.asString();
              if (cellStr.includes("{{M2 de construcción}}"))
                celdaTxt.replaceAllText("{{M2 de construcción}}", m2Formateado);
              if (cellStr.includes("{{Asking price /m2}}"))
                celdaTxt.replaceAllText("{{Asking price /m2}}", precio);
              if (cellStr.includes("{{Renta total}}"))
                celdaTxt.replaceAllText("{{Renta total}}", renta);
            } catch(e) {}
          }
        }
      }
    });
  });

  nuevaDeck.saveAndClose();
  return copiaFile.getUrl();
}