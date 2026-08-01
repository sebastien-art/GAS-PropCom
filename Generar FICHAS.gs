/**
 * GENERADOR AUTOMÁTICO DE FICHAS TÉCNICAS (SLIDES & PDF) VÍA WEBAPP DE INVENTARIO
 */

// IDs de Carpetas
const CARPETA_ORIGEN_ID      = '1xSLs5GHRHbm9OMJwQldxmipNIhPT8eVc';
const FOLDER_MAESTRA_INV_ID = '1E9291Gm9a2wdYRqL9uglDUTLNxojfelb';
const FOLDER_MAESTRA_CLI_ID = '1CNxLz5Xj4P3qwMSTa4uHgbcb_Wy5pwue';

// URL de la WebApp de Inventario
const INVENTARIO_WEB_APP_URL = "https://script.google.com/macros/s/AKfycbzpUroVXu87JyY05EZ9MvF9gc1vI5ljsQ-gPgDgANIMkVwMoVxe88L7EghjFTdrn3pUxA/exec";

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

    if (respuesta !== ui.Button.YES) return;

    const ssActive = SpreadsheetApp.getActiveSpreadsheet();
    const sheetPropCom = ssActive.getActiveSheet();
    const nombreClienteFolder = sheetPropCom.getName().trim().toUpperCase();
    const rawData = sheetPropCom.getDataRange().getDisplayValues();

    console.log("=== INICIO DE PROCESO DE GENERACIÓN DE FICHAS ===");

    // 1. LOCALIZAR TABLA PROPUESTAS COMERCIALES
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
      SpreadsheetApp.getUi().alert("Error: No se encontró la tabla principal.");
      return;
    }

    // 2. MAPEO DE COLUMNAS
    const headers = rawData[headerRowIdx].map(h => h.toString().trim().toLowerCase());
    const col = {
      partida:       headers.indexOf("partida"),
      ref:           headers.indexOf("ref"),
      ficha:         headers.indexOf("ficha"),
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
      SpreadsheetApp.getUi().alert("Error: Faltan columnas en la tabla.");
      return;
    }

    // La columna para escribir la explicación será la celda a la izquierda de 'Ficha' (si 'Ficha' existe)
    // De lo contrario, a la izquierda de 'REF'
    const colTargetExplicacion = col.ficha > 0 ? (col.ficha - 1) : (col.ref > 0 ? col.ref - 1 : -1);

    // 3. CONSULTAR MAPA DE HERMANAS DE INVENTARIO VÍA WEBAPP
    console.log("Solicitando mapa de naves hermanas a la WebApp...");
    const datosHermanas = obtenerMapaHermanasWebAPP();
    const mapaDuo = datosHermanas.mapaDuo || {};
    const mapaParqueSolo = datosHermanas.mapaParqueSolo || {};

    // 4. INDEXAR PLANTILLAS SLIDES
    const inicioIndexacion = Date.now();
    const cacheModelos = [];
    const folderOrigen = DriveApp.getFolderById(CARPETA_ORIGEN_ID);

    indexarCarpeta(folderOrigen, cacheModelos, inicioIndexacion);

    if (cacheModelos.length === 0) {
      SpreadsheetApp.getUi().alert("No se encontraron plantillas de Slides válidas.");
      return;
    }

    const folderCliente = obtenerOCrearSubcarpeta(DriveApp.getFolderById(FOLDER_MAESTRA_CLI_ID), nombreClienteFolder);

    // 5. PROCESAMIENTO DE FILAS
    let procesadasConFicha = 0;
    let refsNuevasConFicha = 0;
    let refsNuevasSinFicha = 0;
    let existSinFicha = 0;
    let errores = 0;

    for (let i = headerRowIdx + 1; i < rawData.length; i++) {
      const row = rawData[i];
      const partidaStr = col.partida !== -1 ? row[col.partida].toString().trim() : "";
      if (!/^\d+$/.test(partidaStr)) continue;

      const numFilaReal = i + 1;
      const celdaRefRango = sheetPropCom.getRange(numFilaReal, col.ref + 1);
      const celdaStatusRango = colTargetExplicacion !== -1 ? sheetPropCom.getRange(numFilaReal, colTargetExplicacion + 1) : null;

      let refTexto = celdaRefRango.getValue().toString().trim();
      const tieneRefOriginal = refTexto && refTexto.toUpperCase() !== "CREAR" && refTexto.toLowerCase() !== "seleccionar...";

      const devStr       = col.desarrollador !== -1 ? row[col.desarrollador].toString().trim() : "";
      const parqStr      = col.parque !== -1 ? row[col.parque].toString().trim() : "";
      const zonaStr      = col.zona !== -1 ? row[col.zona].toString().trim() : "";
      const subzonaStr   = col.subzona !== -1 ? row[col.subzona].toString().trim() : "";
      const estadoStr    = col.estado !== -1 ? row[col.estado].toString().trim() : "";
      const m2Sugerida   = row[col.m2Sugerida] ? row[col.m2Sugerida].toString().trim() : "";
      const precioM2     = row[col.precioM2] ? row[col.precioM2].toString().trim() : "";
      const precioTotal  = row[col.precioTotal] ? row[col.precioTotal].toString().trim() : "";
      const operacionStr = col.operacion !== -1 ? row[col.operacion].toString().trim() : "RENTA";

      let logFila = `[Fila ${numFilaReal} | Partida ${partidaStr}] Parque: "${parqStr}", Dev: "${devStr}" -> `;

      if (!m2Sugerida || m2Sugerida === "0" || !precioM2 || !precioTotal) {
        errores++;
        logFila += "⚠️ Omitida (faltan M2 o Precios).";
        console.warn(logFila);
        if (celdaStatusRango) celdaStatusRango.setValue("⚠️ Faltan datos de M2/Precio");
        continue;
      }

      let modeloEncontrado = null;
      let detalleMatch = "";

      // A) Búsqueda directa por REF original
      if (tieneRefOriginal) {
        const refLimpia = extraerCodigoRefLimpio(refTexto);
        modeloEncontrado = cacheModelos.find(m => m.tieneVariables && m.refsEncontradas.includes(refLimpia));
        if (modeloEncontrado) {
          detalleMatch = `Modelo hallado directo por REF ${refLimpia} (Slide: ${modeloEncontrado.fileName})`;
        }
      }

      // B) Búsqueda por Hermanas
      if (!modeloEncontrado && parqStr) {
        const kParque = limpiarTextoClave(parqStr);
        const kDuo = devStr ? `${kParque}|${limpiarTextoClave(devStr)}` : kParque;

        let hermanasInv = mapaDuo[kDuo] || [];
        let origenBusqueda = "Dúo";

        if (hermanasInv.length === 0) {
          hermanasInv = mapaParqueSolo[kParque] || [];
          origenBusqueda = "Parque solo";
        }

        for (const refHermana of hermanasInv) {
          modeloEncontrado = cacheModelos.find(m => m.tieneVariables && m.refsEncontradas.includes(refHermana));
          if (modeloEncontrado) {
            detalleMatch = `Modelo hallado usando hermana ${refHermana} (${origenBusqueda}) en Slide: ${modeloEncontrado.fileName}`;
            break;
          }
        }

        if (!modeloEncontrado) {
          detalleMatch = `Sin modelo en Slides. Hermanas probadas (${hermanasInv.length}): [${hermanasInv.join(", ")}]`;
        }
      }

      // C) Generación o asignación de REF
      let refFinal = "";
      let esRefRecienCreada = false;

      if (tieneRefOriginal) {
        refFinal = refTexto;
      } else {
        const numSugM2 = parseFloat(String(m2Sugerida).replace(/[^\d.]/g, "")) || 0;
        refFinal = solicitarNuevaRefWebAPP(parqStr, numSugM2, operacionStr);
        
        if (!refFinal) {
          errores++;
          logFila += "❌ Error al generar REF vía WebApp.";
          console.error(logFila);
          if (celdaStatusRango) celdaStatusRango.setValue("❌ Error WebApp al crear REF");
          continue;
        }
        esRefRecienCreada = true;
      }

      // D) Generar PDF
      if (modeloEncontrado) {
        try {
          const zonaSubzonaLimpia = (subzonaStr || zonaStr).toUpperCase();
          const estadoAbreviado = obtenerEstadoFormateado(estadoStr);
          const m2Limpio = m2Sugerida.replace(/[^\d.,]/g, '');

          const nombreInventario = `${operacionStr.toUpperCase()} ${m2Limpio} M2 ${zonaSubzonaLimpia} ${estadoAbreviado} ${refFinal}`.toUpperCase();
          const nombreCliente = `${partidaStr} - ${refFinal} - ${operacionStr.toUpperCase()} ${m2Limpio} M2 ${zonaSubzonaLimpia} ${estadoAbreviado}`.toUpperCase();

          const folderOperacionInv = obtenerCarpetaOperacionInventario(operacionStr);
          const folderEstadoInv = obtenerOCrearSubcarpeta(folderOperacionInv, estadoStr.toUpperCase() || "SIN ESTADO");

          const pdfBlob = generarPdfDesdeModelo(
            modeloEncontrado, partidaStr, m2Sugerida, precioM2, precioTotal,
            refFinal, cacheModelos
          );

          pdfBlob.setName(nombreInventario + ".pdf");
          folderEstadoInv.createFile(pdfBlob);

          const pdfBlobCli = pdfBlob.copyBlob();
          pdfBlobCli.setName(nombreCliente + ".pdf");
          const pdfFileCli = folderCliente.createFile(pdfBlobCli);
          const urlPdfCli = pdfFileCli.getUrl();

          const richValueRef = SpreadsheetApp.newRichTextValue()
            .setText(refFinal)
            .setLinkUrl(0, refFinal.length, urlPdfCli)
            .build();

          celdaRefRango.clearDataValidations();
          celdaRefRango.setRichTextValue(richValueRef);
          celdaRefRango.setFontLine("underline").setFontColor("#1155cc");

          procesadasConFicha++;
          if (esRefRecienCreada) refsNuevasConFicha++;

          logFila += `✅ Creada (${refFinal}) | ${detalleMatch}`;
          console.log(logFila);
          if (celdaStatusRango) celdaStatusRango.setValue(`✅ Ficha Creada (${refFinal}) - ${detalleMatch}`);

        } catch (err) {
          errores++;
          logFila += `❌ Error PDF: ${err.toString()}`;
          console.error(logFila);
          if (celdaStatusRango) celdaStatusRango.setValue(`❌ Error creando PDF: ${err.toString()}`);
        }
      } else {
        if (esRefRecienCreada) {
          celdaRefRango.clearDataValidations();
          celdaRefRango.setValue(refFinal);
          celdaRefRango.setFontLine("none").setFontColor("#000000");
          refsNuevasSinFicha++;
        } else {
          existSinFicha++;
        }

        logFila += `⚠️ ${detalleMatch}`;
        console.warn(logFila);
        if (celdaStatusRango) celdaStatusRango.setValue(`⚠️ ${detalleMatch}`);
      }
    }

    // RESUMEN FINAL
    const totalRefsCreadas = refsNuevasConFicha + refsNuevasSinFicha;
    SpreadsheetApp.getUi().alert(
      `🎉 Proceso finalizado.\n\n` +
      `🆔 REFs creadas por WebApp: ${totalRefsCreadas}\n` +
      `   ├─ CON ficha PDF: ${refsNuevasConFicha}\n` +
      `   └─ SIN ficha PDF (sin plantilla): ${refsNuevasSinFicha}\n\n` +
      `📄 Total fichas PDF generadas: ${procesadasConFicha}\n` +
      `⚠️ Filas con REF existente pero sin plantilla: ${existSinFicha}\n` +
      `❌ Errores / Filas incompletas: ${errores}\n\n` +
      `📌 Revisa los detalles impresos a la izquierda de la columna Ficha.`
    );

  } finally {
    lock.releaseLock();
  }
}

// AUXILIARES CONEXIÓN WEBAPP
function obtenerMapaHermanasWebAPP() {
  try {
    const url = INVENTARIO_WEB_APP_URL + "?action=obtenerMapaHermanas";
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const json = JSON.parse(response.getContentText());
    return json.status === "success" ? json : { mapaDuo: {}, mapaParqueSolo: {} };
  } catch (e) {
    console.error("Error consultando mapa de hermanas a la WebApp: " + e.toString());
    return { mapaDuo: {}, mapaParqueSolo: {} };
  }
}

function solicitarNuevaRefWebAPP(parque, m2, operacion) {
  try {
    const url = INVENTARIO_WEB_APP_URL 
      + "?action=crearNuevaRef"
      + "&parque=" + encodeURIComponent(parque)
      + "&m2=" + encodeURIComponent(m2)
      + "&operacion=" + encodeURIComponent(operacion);

    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const json = JSON.parse(response.getContentText());
    return (json.status === "success" && json.nuevaRef) ? json.nuevaRef : null;
  } catch (e) {
    return null;
  }
}

function extraerCodigoRefLimpio(texto) {
  if (!texto) return "";
  const match = String(texto).toUpperCase().match(/N\d+/);
  return match ? match[0] : String(texto).replace(/[^A-Z0-9]/gi, '').toUpperCase();
}

function limpiarTextoClave(txt) {
  if (!txt) return "";
  return String(txt).toLowerCase().replace(/[^a-z0-9]/g, '');
}

function obtenerEstadoFormateado(estado) {
  if (!estado) return "";
  const estUpper = estado.trim().toUpperCase();
  return ABREVIATURAS_ESTADO[estUpper] || estUpper;
}

function obtenerOCrearSubcarpeta(parentFolder, nombreSubcarpeta) {
  const folders = parentFolder.getFoldersByName(nombreSubcarpeta);
  return folders.hasNext() ? folders.next() : parentFolder.createFolder(nombreSubcarpeta);
}

function obtenerCarpetaOperacionInventario(operacion) {
  const folderMaestra = DriveApp.getFolderById(FOLDER_MAESTRA_INV_ID);
  const nombreSub = operacion.toUpperCase().includes("VENTA") ? "2. VENTA NAVES" : "1. RENTA NAVES";
  return obtenerOCrearSubcarpeta(folderMaestra, nombreSub);
}

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
          const textoCompleto = extraerTodoElTexto(slides[i]);
          const tieneVariables = textoCompleto.includes("{{M2 de construcción}}") || 
                                 textoCompleto.includes("{{Precio total}}") || 
                                 textoCompleto.includes("{{Asking price /m2}}");

          const matchesRef = textoCompleto.match(/N\d+/g) || [];
          const refsEncontradas = matchesRef.map(r => r.toUpperCase());

          cacheModelos.push({
            fileId: file.getId(),
            fileName: file.getName(),
            slideIdx: i,
            textoRaw: textoCompleto,
            refsEncontradas: refsEncontradas,
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

  const copiaFile = DriveApp.getFileById(modelo.fileId).makeCopy(nombreTemp);
  const nuevaDeck = SlidesApp.openById(copiaFile.getId());

  const slidesNuevos = nuevaDeck.getSlides();
  for (let i = slidesNuevos.length - 1; i >= 0; i--) {
    if (!idsAConservar.includes(slidesNuevos[i].getObjectId())) {
      try { slidesNuevos[i].remove(); } catch (e) {}
    }
  }

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

  const pdfBlob = copiaFile.getAs(MimeType.PDF);
  copiaFile.setTrashed(true);

  return pdfBlob;
}