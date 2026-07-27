/**
 * GENERADOR AUTOMÁTICO DE FICHAS TÉCNICAS (SLIDES & PDF)
 */

const CARPETA_ORIGEN_ID  = '1xSLs5GHRHbm9OMJwQldxmipNIhPT8eVc';
const FOLDER_MAESTRA_INV_ID = '1E9291Gm9a2wdYRqL9uglDUTLNxojfelb';
const FOLDER_MAESTRA_CLI_ID = '1CNxLz5Xj4P3qwMSTa4uHgbcb_Wy5pwue';

// Mapeo de abreviaturas para Estados
const ABREVIATURAS_ESTADO = {
  "BAJA CALIFORNIA": "BC",
  "GUANAJUATO": "GTO",
  "NUEVO LEÓN": "NL",
  "NUEVO LEON": "NL",
  "QUERÉTARO": "QRO",
  "QUERETARO": "QRO",
  "QUINTANA ROO": "Q. ROO",
  "SAN LUIS POTOSÍ": "SLP",
  "SAN LUIS POTOSI": "SLP"
};

function generarFichasPropuestasComerciales() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    SpreadsheetApp.getUi().alert("⚠️ El proceso ya se está ejecutando. Por favor, espera a que termine.");
    return;
  }

  try {
    const ui = SpreadsheetApp.getUi();
    const respuesta = ui.alert(
      'Confirmación de generación',
      '¿Estás seguro de que deseas generar las fichas y actualizar el inventario?',
      ui.ButtonSet.YES_NO
    );

    if (respuesta !== ui.Button.YES) {
      return;
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetPropCom = ss.getActiveSheet();
    const nombreClienteFolder = sheetPropCom.getName().trim().toUpperCase();
    const rawData = sheetPropCom.getDataRange().getDisplayValues();

    // 1. LOCALIZAR TABLA PROPUESTAS COMERCIALES (TABLA AZUL)
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

    // 2. MAPEO DE COLUMNAS EXACTAS
    const headers = rawData[headerRowIdx].map(h => h.toString().trim().toLowerCase());
    const col = {
      partida:       headers.indexOf("partida"),
      ref:           headers.indexOf("ref"),
      parque:        headers.findIndex(h => h.includes("parque")),
      desarrollador: headers.findIndex(h => h.includes("desarrollador")),
      zona:          headers.findIndex(h => h.includes("zona principal")),
      subzona:       headers.findIndex(h => h.includes("sub zona") || h.includes("subzona")),
      estado:        headers.findIndex(h => h.includes("estado")),
      intermediario: headers.indexOf("intermediario"),
      m2Sugerida:    headers.findIndex(h => h.includes("superficie sugerida")),
      precioM2:      headers.findIndex(h => h.includes("precio por m2") || h.includes("asking")),
      precioTotal:   headers.findIndex(h => h.includes("precio total")),
      operacion:     headers.findIndex(h => h.includes("operación") || h.includes("operacion"))
    };

    if (col.ref === -1 || col.m2Sugerida === -1 || col.precioM2 === -1 || col.precioTotal === -1) {
      SpreadsheetApp.getUi().alert("Error: No se encontraron todas las columnas necesarias en la tabla.");
      return;
    }

    // 3. CONSTRUIR TRÍO DE HERMANAS PARA MODELOS EN CACHÉ
    const trioMap = {};
    for (let i = headerRowIdx + 1; i < rawData.length; i++) {
      const row = rawData[i];
      const refCelda = row[col.ref] ? row[col.ref].toString().trim() : "";
      const devStr   = col.desarrollador !== -1 ? row[col.desarrollador].toString().trim() : "";
      const parqStr  = col.parque !== -1 ? row[col.parque].toString().trim() : "";
      const zonaStr  = col.zona !== -1 ? row[col.zona].toString().trim() : "";
      if (refCelda && devStr && parqStr && zonaStr) {
        const key = `${zonaStr}|${parqStr}|${devStr}`.toLowerCase();
        if (!trioMap[key]) trioMap[key] = [];
        refCelda.split(",").forEach(r => {
          const ref = r.trim();
          if (ref && !trioMap[key].includes(ref)) trioMap[key].push(ref);
        });
      }
    }

    // 4. INDEXAR PLANTILLAS SLIDES
    console.log("=== INICIO INDEXACIÓN DE PLANTILLAS ===");
    const inicioIndexacion = Date.now();
    const cacheModelos = [];
    const folderOrigen = DriveApp.getFolderById(CARPETA_ORIGEN_ID);

    indexarCarpeta(folderOrigen, cacheModelos, inicioIndexacion);

    if (cacheModelos.length === 0) {
      SpreadsheetApp.getUi().alert("No se encontraron plantillas de slides válidas en la carpeta de origen.");
      return;
    }

    // 5. OBTENER CARPETA CONTENEDORA DE CLIENTE
    const folderCliente = obtenerOCrearSubcarpeta(DriveApp.getFolderById(FOLDER_MAESTRA_CLI_ID), nombreClienteFolder);

    // MAPEO DE PESTAÑA NAVES PARA ACTUALIZACIÓN EN INVENTARIO
    const sheetNaves = ss.getSheetByName("Naves");
    let mapNaves = {};
    if (sheetNaves) {
      const dataNaves = sheetNaves.getDataRange().getDisplayValues();
      if (dataNaves.length > 0) {
        const headersNaves = dataNaves[0].map(h => h.toString().trim().toLowerCase());
        const colRefNaves = headersNaves.indexOf("ref");
        const colFichaNaves = headersNaves.indexOf("ficha");
        if (colRefNaves !== -1 && colFichaNaves !== -1) {
          for (let r = 1; r < dataNaves.length; r++) {
            const refVal = dataNaves[r][colRefNaves].toString().trim().toUpperCase();
            if (refVal) {
              mapNaves[refVal] = {
                rowIdx: r + 1,
                colFichaIdx: colFichaNaves + 1
              };
            }
          }
        }
      }
    }

    // 6. PROCESAR CADA FILA DE LA TABLA PROPUESTAS COMERCIALES
    let procesadas = 0, errores = 0, sinModelo = 0;

    for (let i = headerRowIdx + 1; i < rawData.length; i++) {
      const row = rawData[i];
      const partidaStr = col.partida !== -1 ? row[col.partida].toString().trim() : "";
      if (!/^\d+$/.test(partidaStr)) continue; // Solo filas con partida válida

      const celdaRefRango = sheetPropCom.getRange(i + 1, col.ref + 1);
      const richTextRef = celdaRefRango.getRichTextValue();
      let refTexto = celdaRefRango.getValue().toString().trim();
      let urlExistentePropCom = richTextRef ? (richTextRef.getLinkUrl() || "") : "";

      const tieneRef = refTexto && refTexto.toUpperCase() !== "CREAR" && refTexto.toLowerCase() !== "seleccionar...";
      const tieneFichaPropCom = Boolean(urlExistentePropCom);

      // EXTRAER DATOS
      const devStr       = col.desarrollador !== -1 ? row[col.desarrollador].toString().trim() : "";
      const parqStr      = col.parque !== -1 ? row[col.parque].toString().trim() : "";
      const zonaStr      = col.zona !== -1 ? row[col.zona].toString().trim() : "";
      const subzonaStr   = col.subzona !== -1 ? row[col.subzona].toString().trim() : "";
      const estadoStr    = col.estado !== -1 ? row[col.estado].toString().trim() : "";
      const m2Sugerida   = row[col.m2Sugerida] ? row[col.m2Sugerida].toString().trim() : "";
      const precioM2     = row[col.precioM2] ? row[col.precioM2].toString().trim() : "";
      const precioTotal  = row[col.precioTotal] ? row[col.precioTotal].toString().trim() : "";
      const operacionStr = col.operacion !== -1 ? row[col.operacion].toString().trim() : "RENTA";

      if (!m2Sugerida || m2Sugerida === "0" || !precioM2 || !precioTotal) {
        errores++;
        console.warn(`[DATOS INCOMPLETOS] Fila ${i + 1} omitida por falta de Superficie sugerida, Precio por m2 o Precio total.`);
        continue;
      }

      const trioKey = `${zonaStr}|${parqStr}|${devStr}`.toLowerCase();
      const hermanasDelTrio = trioMap[trioKey] || [];

      // DETERMINAR BÚSQUEDA DE MODELO
      let modeloEncontrado = null;
      if (tieneRef) {
        const refClean = refTexto.replace(/\s+/g, '').toUpperCase();
        modeloEncontrado = cacheModelos.find(m => m.tieneVariables && m.textoLimpio.includes(refClean));
      }

      if (!modeloEncontrado && hermanasDelTrio.length > 0) {
        for (const refHermana of hermanasDelTrio) {
          const hermanaClean = refHermana.replace(/\s+/g, '').toUpperCase();
          modeloEncontrado = cacheModelos.find(m => m.tieneVariables && m.textoLimpio.includes(hermanaClean));
          if (modeloEncontrado) break;
        }
      }

      // EVALUACIÓN DE CASOS SEGÚN EXISTENCIA DE MODELO
      if (modeloEncontrado) {
        // EVALUAR CÓDIGO DE REF (Usar existente o crear consecutivo)
        let refFinal = "";
        if (tieneRef) {
          refFinal = refTexto;
        } else {
          refFinal = obtenerSiguienteRefConsecutiva(sheetPropCom);
        }

        try {
          // PREPARAR NOMBRES Y UBICACIONES
          const zonaSubzonaLimpia = (subzonaStr || zonaStr).toUpperCase();
          const estadoAbreviado = obtenerEstadoFormateado(estadoStr);
          const m2Limpio = m2Sugerida.replace(/[^\d.,]/g, '');

          // Nomenclatura 1: Inventario
          const nombreInventario = `${operacionStr.toUpperCase()} ${m2Limpio} M2 ${zonaSubzonaLimpia} ${estadoAbreviado} ${refFinal}`.toUpperCase();

          // Nomenclatura 2: Cliente (PropCom)
          const nombreCliente = `${partidaStr} - ${refFinal} - ${operacionStr.toUpperCase()} ${m2Limpio} M2 ${zonaSubzonaLimpia} ${estadoAbreviado}`.toUpperCase();

          // Ubicación Inventario (Carpetas Maestra -> Operación -> Estado)
          const folderOperacionInv = obtenerCarpetaOperacionInventario(operacionStr);
          const folderEstadoInv = obtenerOCrearSubcarpeta(folderOperacionInv, estadoStr.toUpperCase() || "SIN ESTADO");

          // GENERACIÓN Y DUALIDAD DE PDF
          const pdfBlob = generarPdfDesdeModelo(
            modeloEncontrado, partidaStr, m2Sugerida, precioM2, precioTotal,
            refFinal, cacheModelos
          );

          // Guardar PDF en Inventario
          pdfBlob.setName(nombreInventario + ".pdf");
          const pdfFileInv = folderEstadoInv.createFile(pdfBlob);
          const urlPdfInv = pdfFileInv.getUrl();

          // Guardar PDF en Cliente
          const pdfBlobCli = pdfBlob.copyBlob();
          pdfBlobCli.setName(nombreCliente + ".pdf");
          const pdfFileCli = folderCliente.createFile(pdfBlobCli);
          const urlPdfCli = pdfFileCli.getUrl();

          // ACTUALIZAR PROCESO EN TABLA PROPUESTAS COMERCIALES
          const richValueRef = SpreadsheetApp.newRichTextValue()
            .setText(refFinal)
            .setLinkUrl(0, refFinal.length, urlPdfCli)
            .build();

          celdaRefRango.clearDataValidations();
          celdaRefRango.setRichTextValue(richValueRef);
          celdaRefRango.setFontLine("underline").setFontColor("#1155cc");

          // ACTUALIZAR INVENTARIO (Pestaña "Naves" -> Columna "Ficha" -> Escribir "CLIC")
          if (sheetNaves && mapNaves[refFinal.toUpperCase()]) {
            const navObj = mapNaves[refFinal.toUpperCase()];
            const celdaFichaNaves = sheetNaves.getRange(navObj.rowIdx, navObj.colFichaIdx);
            const richValueClic = SpreadsheetApp.newRichTextValue()
              .setText("CLIC")
              .setLinkUrl(0, 4, urlPdfInv)
              .build();
            celdaFichaNaves.setRichTextValue(richValueClic);
            celdaFichaNaves.setFontLine("underline").setFontColor("#1155cc");
          }

          procesadas++;
          console.log(`[MODELO OK] Fila ${i + 1} | REF: ${refFinal} → PDFs creados e inventariados.`);

        } catch (err) {
          errores++;
          console.error(`[ERROR] Fila ${i + 1} | REF: ${refTexto} → ${err.toString()}`);
        }

      } else {
        // CASO: "SIN MODELO"
        if (!tieneRef) {
          // Sin REF y Sin Modelo: Se crea REF nueva sin hipervínculo en PropCom, no se hace nada en inventario
          const refNueva = obtenerSiguienteRefConsecutiva(sheetPropCom);
          celdaRefRango.clearDataValidations();
          celdaRefRango.setValue(refNueva);
          celdaRefRango.setFontLine("none").setFontColor("#000000");
          sinModelo++;
          console.log(`[SIN MODELO] Fila ${i + 1} | Se generó la REF nueva ${refNueva} sin hipervínculo.`);
        } else {
          // Tiene REF y no hay modelo: No se hace nada
          sinModelo++;
          console.warn(`[SIN MODELO] Fila ${i + 1} | REF: ${refTexto} | Trío: ${trioKey} - Sin cambios.`);
        }
      }
    }

    SpreadsheetApp.getUi().alert(
      `Proceso finalizado.\n\n✅ Fichas PDF Creadas/Actualizadas: ${procesadas}\n⚠️ Sin Modelo Encontrado / Procesados Sin Ficha: ${sinModelo}\n❌ Errores: ${errores}`
    );
  } finally {
    lock.releaseLock();
  }
}

/**
 * FUNCION AUXILIAR: Obtiene el siguiente consecutivo de REF en la tabla de inventario (Ej. N1155)
 */
function obtenerSiguienteRefConsecutiva(sheet) {
  const data = sheet.getDataRange().getValues();
  let maxNum = 0;
  let prefijo = "N";

  for (let r = 0; r < data.length; r++) {
    for (let c = 0; c < data[r].length; c++) {
      let val = String(data[r][c]).trim();
      let match = val.match(/^([A-Za-z]+)(\d+)$/);
      if (match) {
        let pFix = match[1].toUpperCase();
        let num = parseInt(match[2], 10);
        if (pFix === "N" && num > maxNum) {
          maxNum = num;
        }
      }
    }
  }

  if (maxNum === 0) maxNum = 1000;
  return prefijo + (maxNum + 1);
}

/**
 * ABREVIACIÓN / FORMATEO DE ESTADOS
 */
function obtenerEstadoFormateado(estado) {
  if (!estado) return "";
  const estUpper = estado.trim().toUpperCase();
  return ABREVIATURAS_ESTADO[estUpper] || estUpper;
}

/**
 * OBTENER O CREAR SUBCARPETAS
 */
function obtenerOCrearSubcarpeta(parentFolder, nombreSubcarpeta) {
  const folders = parentFolder.getFoldersByName(nombreSubcarpeta);
  if (folders.hasNext()) {
    return folders.next();
  }
  return parentFolder.createFolder(nombreSubcarpeta);
}

/**
 * RESOLVER CARPETA DE OPERACIÓN EN MAESTRA INVENTARIO
 */
function obtenerCarpetaOperacionInventario(operacion) {
  const folderMaestra = DriveApp.getFolderById(FOLDER_MAESTRA_INV_ID);
  const opUpper = operacion.toUpperCase();
  let nombreSub = "1. RENTA NAVES";
  if (opUpper.includes("VENTA")) {
    nombreSub = "2. VENTA NAVES";
  }
  return obtenerOCrearSubcarpeta(folderMaestra, nombreSub);
}

/**
 * INDEXA LA CARPETA Y SUS SUB-CARPETAS
 */
function indexarCarpeta(carpeta, cacheModelos, inicioMs) {
  const TIMEOUT_MS = 240000;
  let files = carpeta.getFilesByType(MimeType.GOOGLE_SLIDES);

  while (files.hasNext()) {
    if (Date.now() - inicioMs > TIMEOUT_MS) return;

    const file = files.next();
    try {
      const deck = SlidesApp.openById(file.getId());
      const slides = deck.getSlides();

      for (let i = 0; i < slides.length; i++) {
        try {
          const texto = extraerTodoElTexto(slides[i]);
          const tieneVariables = texto.includes("{{M2 de construcción}}") || texto.includes("{{Precio total}}") || texto.includes("{{Asking price /m2}}");
          cacheModelos.push({
            fileId: file.getId(),
            fileName: file.getName(),
            slideIdx: i,
            textoLimpio: texto.replace(/\s+/g, '').toUpperCase(),
            tieneVariables: tieneVariables
          });
        } catch (eSlide) {}
      }
    } catch (eFile) {}
  }

  let subFolders = carpeta.getFolders();
  while (subFolders.hasNext()) {
    if (Date.now() - inicioMs > TIMEOUT_MS) return;
    indexarCarpeta(subFolders.next(), cacheModelos, inicioMs);
  }
}

/**
 * EXTRAE TODO EL TEXTO DE UN SLIDE
 */
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
            const t = tabla.getCell(r, c).getText().asString();
            if (t) textos.push(t);
          }
        }
      }
    } catch (e) {}
  });
  return textos.join(" ");
}

/**
 * GENERA EL BLOB DEL PDF A PARTIR DE UN MODELO DE SLIDE
 */
function generarPdfDesdeModelo(modelo, partida, m2, precio, precioTotal, refNueva, cacheModelos) {
  const moldeDeck = SlidesApp.openById(modelo.fileId);
  const slides = moldeDeck.getSlides();

  let finIdx = modelo.slideIdx + 1;
  while (finIdx < slides.length) {
    const texto = extraerTodoElTexto(slides[finIdx]).toUpperCase();
    const esInicioPropiedadCache = cacheModelos && cacheModelos.some(m => m.fileId === modelo.fileId && m.slideIdx === finIdx && m.tieneVariables);
    const tieneVariablesNueva = texto.includes("{{M2 DE CONSTRUCCIÓN}}") || texto.includes("{{PRECIO TOTAL}}") || texto.includes("{{ASKING PRICE /M2}}");
    if (texto.includes("REF:") || esInicioPropiedadCache || tieneVariablesNueva) break;
    finIdx++;
  }

  const idsAConservar = [];
  for (let i = modelo.slideIdx; i < finIdx; i++) {
    idsAConservar.push(slides[i].getObjectId());
  }

  const m2Formateado = String(m2).toLowerCase().includes("m") ? m2 : `${m2} m²`;
  const nombreTemp = `TEMP_${partida}_${refNueva}`;

  // Copia temporal de la presentación
  const copiaFile = DriveApp.getFileById(modelo.fileId).makeCopy(nombreTemp);
  const nuevaDeck = SlidesApp.openById(copiaFile.getId());

  // Limpiar slides sobrantes
  const slidesNuevos = nuevaDeck.getSlides();
  for (let i = slidesNuevos.length - 1; i >= 0; i--) {
    if (!idsAConservar.includes(slidesNuevos[i].getObjectId())) {
      try { slidesNuevos[i].remove(); } catch (e) {}
    }
  }

  // Reemplazar etiquetas solicitadas
  nuevaDeck.getSlides().forEach(s => {
    s.getPageElements().forEach(el => {
      if (el.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
        try {
          const shape = el.asShape();
          const txt = shape.getText().asString();
          if (txt.toUpperCase().includes("REF:")) {
            shape.getText().setText("REF: " + refNueva);
          }
        } catch (e) {}
      }

      if (el.getPageElementType() === SlidesApp.PageElementType.TABLE) {
        const tabla = el.asTable();
        for (let r = 0; r < tabla.getNumRows(); r++) {
          for (let c = 0; c < tabla.getNumColumns(); c++) {
            try {
              const celdaTxt = tabla.getCell(r, c).getText();
              const cellStr = celdaTxt.asString();
              
              if (cellStr.includes("{{M2 de construcción}}"))
                celdaTxt.replaceAllText("{{M2 de construcción}}", m2Formateado);
              
              if (cellStr.includes("{{Asking price /m2}}"))
                celdaTxt.replaceAllText("{{Asking price /m2}}", precio);
              
              if (cellStr.includes("{{Precio total}}"))
                celdaTxt.replaceAllText("{{Precio total}}", precioTotal);
                
            } catch (e) {}
          }
        }
      }
    });
  });

  nuevaDeck.saveAndClose();

  // OBTENER PDF BLOB Y ELIMINAR TEMPORAL DE SLIDES
  const pdfBlob = copiaFile.getAs(MimeType.PDF);
  copiaFile.setTrashed(true);

  return pdfBlob;
}