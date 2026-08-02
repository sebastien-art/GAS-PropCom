/**
 * GENERADOR AUTOMÁTICO DE FICHAS TÉCNICAS Y AUDITORÍA DE PLANTILLAS
 */

// IDs DE CARPETAS DE DRIVE
const CARPETA_ORIGEN_ID          = '1xSLs5GHRHbm9OMJwQldxmipNIhPT8eVc'; // 5. MODELOS FICHAS POR ESTADO
const FOLDER_MAESTRA_INV_ID     = '1E9291Gm9a2wdYRqL9uglDUTLNxojfelb'; // 0. FICHAS (Activas)
const FOLDER_NO_DISPONIBLES_ID  = '18xPP1HXcOoMUv_CtVIzYKRqZwxTAvmXv'; // 7. NO DISPONIBLES FICHAS PASADAS (Resguardo)

// URL DE LA WEBAPP DE INVENTARIO
const INVENTARIO_WEB_APP_URL = "https://script.google.com/macros/s/AKfycbzpUroVXu87JyY05EZ9MvF9gc1vI5ljsQ-gPgDgANIMkVwMoVxe88L7EghjFTdrn3pUxA/exec";

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
    let rawData = sheetPropCom.getDataRange().getDisplayValues();

    console.log("=== INICIO DE PROCESO DE GENERACIÓN DE FICHAS ===");

    let headerRowIdx = -1;
    for (let r = 0; r < rawData.length; r++) {
      const filaBaja = rawData[r].map(c => c.toString().trim().toLowerCase());
      if (filaBaja.includes("partida") && filaBaja.includes("ref") && filaBaja.includes("desarrollador")) {
        headerRowIdx = r;
        break;
      }
    }

    if (headerRowIdx === -1) {
      SpreadsheetApp.getUi().alert("Error: No se encontró la tabla principal.");
      return;
    }

    let headers = rawData[headerRowIdx].map(h => h.toString().trim().toLowerCase());

    let colMantenimiento = headers.findIndex(h => h.includes("mantenimiento"));
    let colPlantillaIdx = headers.indexOf("plantilla");

    if (colPlantillaIdx === -1) {
      if (colMantenimiento !== -1) {
        colPlantillaIdx = colMantenimiento + 1;
        sheetPropCom.insertColumnAfter(colMantenimiento + 1);
      } else {
        colPlantillaIdx = headers.length;
        sheetPropCom.insertColumnAfter(headers.length);
      }
      sheetPropCom.getRange(headerRowIdx + 1, colPlantillaIdx + 1)
        .setValue("Plantilla")
        .setFontWeight("bold")
        .setHorizontalAlignment("center");

      rawData = sheetPropCom.getDataRange().getDisplayValues();
      headers = rawData[headerRowIdx].map(h => h.toString().trim().toLowerCase());
      colPlantillaIdx = headers.indexOf("plantilla");
    }

    const col = {
      partida:       headers.indexOf("partida"),
      ref:           headers.indexOf("ref"),
      ficha:         headers.indexOf("ficha"),
      plantilla:     colPlantillaIdx,
      parque:        headers.findIndex(h => h.includes("parque")),
      desarrollador: headers.findIndex(h => h.includes("desarrollador")),
      zona:          headers.findIndex(h => h.includes("zona principal")),
      subzona:       headers.findIndex(h => h.includes("sub zona") || h.includes("subzona")),
      estado:        headers.findIndex(h => h.includes("estado")),
      m2Sugerida:    headers.findIndex(h => h.includes("superficie sugerida")),
      precioM2:      headers.findIndex(h => h.includes("precio por m2") || h.includes("asking")),
      precioTotal:   headers.findIndex(h => h.includes("precio total")),
      operacion:     headers.findIndex(h => h.includes("operación") || h.includes("operacion"))
    };

    if (col.ref === -1 || col.m2Sugerida === -1 || col.precioM2 === -1) {
      SpreadsheetApp.getUi().alert("Error: Faltan columnas esenciales en la tabla.");
      return;
    }

    const colTargetExplicacion = col.ficha > 0 ? (col.ficha - 1) : (col.ref > 0 ? col.ref - 1 : -1);

    const datosHermanas = obtenerMapaHermanasWebAPP();
    const mapaDuo = datosHermanas.mapaDuo || {};
    const mapaParqueSolo = datosHermanas.mapaParqueSolo || {};

    const cacheCarpetasEstado = {};
    const cacheModelosPorEstado = {};

    let folderOrigenMaestra = null;
    try {
      if (CARPETA_ORIGEN_ID) folderOrigenMaestra = DriveApp.getFolderById(CARPETA_ORIGEN_ID);
    } catch (eDrive) {
      console.warn("No se pudo acceder a CARPETA_ORIGEN_ID.");
    }

    let procesadasConFicha = 0, refsNuevasConFicha = 0, refsNuevasSinFicha = 0, existSinFicha = 0, errores = 0;

    for (let i = headerRowIdx + 1; i < rawData.length; i++) {
      const row = rawData[i];
      const partidaStr = col.partida !== -1 ? row[col.partida].toString().trim() : "";
      if (!/^\d+$/.test(partidaStr)) continue;

      const numFilaReal = i + 1;
      const celdaRefRango = sheetPropCom.getRange(numFilaReal, col.ref + 1);
      const celdaPlantillaRango = sheetPropCom.getRange(numFilaReal, col.plantilla + 1);
      const celdaStatusRango = colTargetExplicacion !== -1 ? sheetPropCom.getRange(numFilaReal, colTargetExplicacion + 1) : null;

      let refTexto = celdaRefRango.getValue().toString().trim();
      const esSolicitudCrear = refTexto.toUpperCase() === "CREAR";
      const tieneRefExistente = refTexto && !esSolicitudCrear && refTexto.toLowerCase() !== "seleccionar...";

      if (!esSolicitudCrear && !tieneRefExistente) continue;

      const devStr       = col.desarrollador !== -1 ? row[col.desarrollador].toString().trim() : "";
      const parqStr      = col.parque !== -1 ? row[col.parque].toString().trim() : "";
      const zonaStr      = col.zona !== -1 ? row[col.zona].toString().trim() : "";
      const subzonaStr   = col.subzona !== -1 ? row[col.subzona].toString().trim() : "";
      const estadoValorRaw = col.estado !== -1 ? row[col.estado] : "";
      const estadoStr    = (estadoValorRaw && estadoValorRaw.toString().trim() !== "") ? estadoValorRaw.toString().trim() : "EdoMex";
      const m2Sugerida   = row[col.m2Sugerida] ? row[col.m2Sugerida].toString().trim() : "";
      const precioM2     = row[col.precioM2] ? row[col.precioM2].toString().trim() : "";
      const precioTotal  = col.precioTotal !== -1 && row[col.precioTotal] ? row[col.precioTotal].toString().trim() : "";
      const operacionStr = (col.operacion !== -1 && row[col.operacion]) ? row[col.operacion].toString().trim() : "RENTA";

      const numSugM2 = parseFloat(String(m2Sugerida).replace(/[^\d.]/g, "")) || 0;

      if (!m2Sugerida || numSugM2 === 0 || !precioM2) {
        errores++;
        if (celdaStatusRango) celdaStatusRango.setValue("⚠️ Faltan datos M2/Precio");
        continue;
      }

      // CANDADO C: VERIFICAR M2 EXACTOS EN REF EXISTENTE
      let urlFichaAnteriorParaResguardo = "";
      if (tieneRefExistente) {
        const refLimpia = extraerCodigoRefLimpio(refTexto);
        const datosRef = consultarDatosRefWebAPP(refLimpia);

        if (datosRef.status === "success" && datosRef.encontrado) {
          urlFichaAnteriorParaResguardo = datosRef.urlFichaAnterior || "";
          
          if (datosRef.m2Original > 0 && datosRef.m2Original !== numSugM2) {
            errores++;
            const msgError = `❌ Error: La superficie ingresada (${numSugM2} m²) no coincide con la superficie original de la REF ${refLimpia} (${datosRef.m2Original} m²)`;
            if (celdaStatusRango) celdaStatusRango.setValue(msgError);
            console.warn(msgError);
            continue;
          }
        }
      }

      // INDEXAR PLANTILLAS DEL ESTADO
      const claveEstado = estadoStr.toLowerCase().replace(/[^a-z]/g, '');
      if (!cacheModelosPorEstado[claveEstado]) {
        cacheModelosPorEstado[claveEstado] = [];
        let folderEstadoModelos = cacheCarpetasEstado[claveEstado];

        if (!folderEstadoModelos && folderOrigenMaestra) {
          const subfolders = folderOrigenMaestra.getFolders();
          while (subfolders.hasNext()) {
            const sub = subfolders.next();
            const nombreLimpio = sub.getName().toLowerCase().replace(/[^a-z]/g, '');
            if (nombreLimpio.includes(claveEstado) || claveEstado.includes(nombreLimpio.replace("fichas", ""))) {
              folderEstadoModelos = sub;
              cacheCarpetasEstado[claveEstado] = sub;
              break;
            }
          }
        }

        const carpetaABuscar = folderEstadoModelos || folderOrigenMaestra;
        if (carpetaABuscar) {
          indexarCarpetaDirecta(carpetaABuscar, cacheModelosPorEstado[claveEstado]);
        }
      }

      const cacheModelos = cacheModelosPorEstado[claveEstado] || [];

      // BUSCAR MODELO
      let modeloEncontrado = null;
      let refCoincidente = "";
      let esPorHermana = false;

      if (tieneRefExistente) {
        const refLimpia = extraerCodigoRefLimpio(refTexto);
        modeloEncontrado = cacheModelos.find(m => m.tieneVariables && m.refsEncontradas.includes(refLimpia));
        if (modeloEncontrado) refCoincidente = refLimpia;
      }

      if (!modeloEncontrado && parqStr) {
        const kParque = limpiarTextoClave(parqStr);
        const kDuo = devStr ? `${kParque}|${limpiarTextoClave(devStr)}` : kParque;
        let hermanasInv = mapaDuo[kDuo] || mapaParqueSolo[kParque] || [];

        for (const refHermana of hermanasInv) {
          modeloEncontrado = cacheModelos.find(m => m.tieneVariables && m.refsEncontradas.includes(refHermana));
          if (modeloEncontrado) {
            refCoincidente = refHermana;
            esPorHermana = true;
            break;
          }
        }
      }

      // COLUMNA PLANTILLA
      if (modeloEncontrado) {
        const slideUrl = `https://docs.google.com/presentation/d/${modeloEncontrado.fileId}/edit#slide=id.${modeloEncontrado.slideObjectId || ""}`;
        const labelPlantilla = esPorHermana ? `Sí (Hermana: ${refCoincidente})` : `Sí (${refCoincidente})`;
        
        const richValuePlantilla = SpreadsheetApp.newRichTextValue()
          .setText(labelPlantilla)
          .setLinkUrl(0, labelPlantilla.length, slideUrl)
          .build();

        celdaPlantillaRango
          .setRichTextValue(richValuePlantilla)
          .setFontColor("#1155cc")
          .setFontLine("underline")
          .setFontWeight("bold")
          .setHorizontalAlignment("center");
      } else {
        celdaPlantillaRango
          .setValue("No")
          .setFontColor("#cc0000")
          .setFontLine("none")
          .setFontWeight("normal")
          .setHorizontalAlignment("center");
      }

      // CREAR O ASIGNAR REF
      let refFinal = "";
      let esRefRecienCreada = false;

      if (tieneRefExistente) {
        refFinal = refTexto;
      } else if (esSolicitudCrear) {
        const resCrear = solicitarNuevaRefWebAPP(parqStr, numSugM2, operacionStr, precioM2);
        
        if (resCrear.status !== "success" || !resCrear.nuevaRef) {
          errores++;
          const msgErr = resCrear.message ? `❌ ${resCrear.message}` : "❌ Error WebApp al crear REF";
          if (celdaStatusRango) celdaStatusRango.setValue(msgErr);
          continue;
        }
        refFinal = resCrear.nuevaRef;
        esRefRecienCreada = true;
      }

      // GENERACIÓN DE PDF
      if (modeloEncontrado) {
        let pdfGeneradoExito = false;
        let detalleError = "";

        try {
          const zonaSubzonaLimpia = (subzonaStr || zonaStr).toUpperCase();
          const estadoAbreviado = obtenerEstadoFormateado(estadoStr);
          
          // FORMATO SIN SEPARADOR DE MILES EN EL TÍTULO (ej: 60000 M2)
          const m2EnteroSinSeparador = Math.round(numSugM2).toString();
          const nombreInventario = `${operacionStr.toUpperCase()} ${m2EnteroSinSeparador} M2 ${zonaSubzonaLimpia} ${estadoAbreviado} ${refFinal}`.toUpperCase();

          const folderOperacionInv = obtenerCarpetaOperacionInventario(operacionStr);
          const folderEstadoInv = obtenerOCrearSubcarpeta(folderOperacionInv, estadoStr.toUpperCase());

          // LIMPIEZA DE DUPLICADOS EN LA CARPETA Y RESGUARDO OBLIGATORIO PREVIO
          limpiarDuplicadosYResguardar(folderEstadoInv, refFinal, urlFichaAnteriorParaResguardo, operacionStr, estadoStr);

          // Generar nuevo PDF
          const pdfBlob = generarPdfDesdeModelo(
            modeloEncontrado, partidaStr, m2Sugerida, precioM2, precioTotal,
            refFinal, cacheModelos
          );

          pdfBlob.setName(nombreInventario + ".pdf");
          const pdfFileInv = folderEstadoInv.createFile(pdfBlob);
          const urlPdfInv = pdfFileInv.getUrl();

          if (urlPdfInv) {
            // A) Actualizar hipervínculo REF en hoja local de Trabajo
            const richValueRef = SpreadsheetApp.newRichTextValue()
              .setText(refFinal)
              .setLinkUrl(0, refFinal.length, urlPdfInv)
              .build();

            celdaRefRango.clearDataValidations();
            celdaRefRango.setRichTextValue(richValueRef);
            celdaRefRango.setFontLine("underline").setFontColor("#1155cc");

            // B) Enviar a la WebApp para actualizar columna FICHA en Inventario Naves
            actualizarFichaEnInventarioWebAPP(refFinal, urlPdfInv);

            pdfGeneradoExito = true;
          }

        } catch (err) {
          detalleError = err.toString();
          console.error(`Error PDF Fila ${numFilaReal}: ` + detalleError);
        }

        if (pdfGeneradoExito) {
          procesadasConFicha++;
          if (esRefRecienCreada) {
            refsNuevasConFicha++;
            if (celdaStatusRango) celdaStatusRango.setValue("✅ REF creada en Inventario con FICHA");
          } else {
            if (celdaStatusRango) celdaStatusRango.setValue("✅ Ficha actualizada en inventario");
          }
        } else {
          errores++;
          const msjCorto = detalleError.length > 35 ? detalleError.substring(0, 35) + "..." : detalleError;
          if (celdaStatusRango) celdaStatusRango.setValue(`❌ Error PDF: ${msjCorto}`);
        }

      } else {
        if (esRefRecienCreada) {
          celdaRefRango.clearDataValidations();
          celdaRefRango.setValue(refFinal);
          celdaRefRango.setFontLine("none").setFontColor("#000000");
          refsNuevasSinFicha++;
          if (celdaStatusRango) celdaStatusRango.setValue("ℹ️ REF creada en Inventario (Sin Ficha PDF)");
        } else {
          existSinFicha++;
          if (celdaStatusRango) celdaStatusRango.setValue("⚠️ Sin modelo, Ficha NO ACTUALIZADA");
        }
      }
    }

    // RESUMEN FINAL
    const totalRefsCreadas = refsNuevasConFicha + refsNuevasSinFicha;
    SpreadsheetApp.getUi().alert(
      `🎉 Proceso finalizado.\n\n` +
      `🆔 REFs creadas en WebApp: ${totalRefsCreadas}\n` +
      `   ├─ CON ficha PDF: ${refsNuevasConFicha}\n` +
      `   └─ SIN ficha PDF: ${refsNuevasSinFicha}\n\n` +
      `📄 Total fichas PDF generadas: ${procesadasConFicha}\n` +
      `⚠️ Sin plantilla encontrada: ${existSinFicha}\n` +
      `❌ Errores / Bloqueados por Candado: ${errores}\n\n` +
      `📌 Proceso completado. Revisa la columna Ficha.`
    );

  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// CONEXIONES CON WEBAPP DE INVENTARIO
// ==========================================

function obtenerMapaHermanasWebAPP() {
  try {
    const url = INVENTARIO_WEB_APP_URL + "?action=obtenerMapaHermanas";
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const json = JSON.parse(response.getContentText());
    return json.status === "success" ? json : { mapaDuo: {}, mapaParqueSolo: {} };
  } catch (e) {
    console.error("Error obteniendo hermanas: " + e.toString());
    return { mapaDuo: {}, mapaParqueSolo: {} };
  }
}

function consultarDatosRefWebAPP(ref) {
  try {
    const url = INVENTARIO_WEB_APP_URL + "?action=obtenerDatosRef&ref=" + encodeURIComponent(ref);
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    return JSON.parse(response.getContentText());
  } catch (e) {
    console.error("Error consultando REF en WebApp: " + e.toString());
    return { status: "error" };
  }
}

function solicitarNuevaRefWebAPP(parque, m2, operacion, precioM2) {
  try {
    const m2Limpio = String(m2).replace(/[^\d.]/g, "");
    const precioLimpio = String(precioM2).replace(/[^\d.]/g, "");

    const url = INVENTARIO_WEB_APP_URL + "?action=crearNuevaRef" +
      "&parque=" + encodeURIComponent(parque || "") +
      "&m2=" + encodeURIComponent(m2Limpio) +
      "&operacion=" + encodeURIComponent(operacion || "") +
      "&precioM2=" + encodeURIComponent(precioLimpio);

    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    return JSON.parse(response.getContentText());
  } catch (e) {
    console.error("Error en solicitarNuevaRefWebAPP: " + e.toString());
    return { status: "error", message: e.toString() };
  }
}

function actualizarFichaEnInventarioWebAPP(ref, urlPdf) {
  try {
    const url = INVENTARIO_WEB_APP_URL 
      + "?action=actualizarFicha"
      + "&ref=" + encodeURIComponent(ref)
      + "&urlPdf=" + encodeURIComponent(urlPdf);

    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    return JSON.parse(response.getContentText());
  } catch (e) {
    console.error("Error enviando Ficha a Inventario: " + e.toString());
    return null;
  }
}

// ==========================================
// LIMPIEZA Y RESGUARDO DE DUPLICADOS EN DRIVE
// ==========================================

function limpiarDuplicadosYResguardar(folderActiva, refBuscada, urlPrevia, operacion, estado) {
  if (!folderActiva || !refBuscada) return;

  const refUpper = refBuscada.toUpperCase();
  const files = folderActiva.getFiles();

  while (files.hasNext()) {
    const file = files.next();
    const nameUpper = file.getName().toUpperCase();

    // Si el archivo existente en la carpeta activa contiene la REF (ej. N3242)
    if (nameUpper.includes(refUpper) && nameUpper.endsWith(".PDF")) {
      moverAResguardo(file, operacion, estado);
    }
  }

  // Si había una URL registrada previamente que estaba en otra carpeta
  if (urlPrevia) {
    try {
      const matchId = urlPrevia.match(/[-\w]{25,}/);
      if (matchId) {
        const filePrevio = DriveApp.getFileById(matchId[0]);
        if (filePrevio && !filePrevio.getName().startsWith("No disponible -")) {
          moverAResguardo(filePrevio, operacion, estado);
        }
      }
    } catch (e) {}
  }
}

function moverAResguardo(file, operacion, estado) {
  if (!FOLDER_NO_DISPONIBLES_ID || !file) return;

  try {
    if (file.getName().startsWith("No disponible -")) return;

    const folderResguardoRaiz = DriveApp.getFolderById(FOLDER_NO_DISPONIBLES_ID);
    const opUpper = (operacion || "").toUpperCase();
    const nombreSubOp = opUpper.includes("VENTA") ? "2. NO DISPONIBLES - VENTA NAVES" : "1. NO DISPONIBLES - RENTA NAVES";
    const folderOpResguardo = obtenerOCrearSubcarpeta(folderResguardoRaiz, nombreSubOp);
    const folderEstadoResguardo = obtenerOCrearSubcarpeta(folderOpResguardo, (estado || "EDOMEX").toUpperCase());

    const nombreOriginal = file.getName();
    file.setName(`No disponible - ${nombreOriginal}`);
    file.moveTo(folderEstadoResguardo);

    console.log(`Archivo movido a Resguardo: ${file.getName()}`);
  } catch (e) {
    console.warn("No se pudo mover archivo a resguardo: " + e.toString());
  }
}

// ==========================================
// AUXILIARES DRIVE Y SLIDES
// ==========================================

function obtenerOCrearSubcarpeta(parentFolder, nombreSubcarpeta) {
  if (!parentFolder) return null;
  const nombreLimpio = (nombreSubcarpeta && nombreSubcarpeta.trim() !== "") ? nombreSubcarpeta.trim() : "EDOMEX";
  const folders = parentFolder.getFoldersByName(nombreLimpio);
  return folders.hasNext() ? folders.next() : parentFolder.createFolder(nombreLimpio);
}

function obtenerCarpetaOperacionInventario(operacion) {
  if (!FOLDER_MAESTRA_INV_ID) throw new Error("FOLDER_MAESTRA_INV_ID no configurado");
  const folderMaestra = DriveApp.getFolderById(FOLDER_MAESTRA_INV_ID);
  const opLimpia = (operacion || "").toUpperCase();
  const nombreSub = opLimpia.includes("VENTA") ? "2. VENTA NAVES" : "1. RENTA NAVES";
  return obtenerOCrearSubcarpeta(folderMaestra, nombreSub);
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

function indexarCarpetaDirecta(carpeta, cacheModelos) {
  if (!carpeta) return;

  let files = carpeta.getFilesByType(MimeType.GOOGLE_SLIDES);
  while (files.hasNext()) {
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

          if (refsEncontradas.length > 0) {
            cacheModelos.push({
              fileId: file.getId(),
              fileName: file.getName(),
              slideIdx: i,
              slideObjectId: slides[i].getObjectId(),
              textoRaw: textoCompleto,
              refsEncontradas: refsEncontradas,
              tieneVariables: tieneVariables
            });
          }
        } catch (eSlide) {}
      }
    } catch (eFile) {}
  }

  let subFolders = carpeta.getFolders();
  while (subFolders.hasNext()) {
    let sub = subFolders.next();
    let subFiles = sub.getFilesByType(MimeType.GOOGLE_SLIDES);
    while (subFiles.hasNext()) {
      const file = subFiles.next();
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

            if (refsEncontradas.length > 0) {
              cacheModelos.push({
                fileId: file.getId(),
                fileName: file.getName(),
                slideIdx: i,
                slideObjectId: slides[i].getObjectId(),
                textoRaw: textoCompleto,
                refsEncontradas: refsEncontradas,
                tieneVariables: tieneVariables
              });
            }
          } catch (eSlide) {}
        }
      } catch (eFile) {}
    }
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
    if (texto.includes("REF:") || esInicioPropiedadCache) break;
    finIdx++;
  }

  const idsAConservar = [];
  for (let i = modelo.slideIdx; i < finIdx; i++) {
    idsAConservar.push(slides[i].getObjectId());
  }

  const m2Formateado = String(m2).toLowerCase().includes("m") ? m2 : `${m2} m²`;
  const nombreTemp = `TEMP_${partida}_${refNueva}_${Date.now()}`;

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
  Utilities.sleep(1000);

  const pdfBlob = copiaFile.getAs(MimeType.PDF);
  copiaFile.setTrashed(true);

  return pdfBlob;
}