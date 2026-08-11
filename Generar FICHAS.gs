/**
 * GENERADOR AUTOMÁTICO DE FICHAS PROPUESTAS COMERCIALES (ULTRA RÁPIDO)
 * Sin recorridos de Google Drive / Slides. Consume las plantillas directo del Inventario.
 */

const FOLDER_MAESTRA_INV_ID     = '1E9291Gm9a2wdYRqL9uglDUTLNxojfelb'; // FICHAS (Activas)
const INVENTARIO_WEB_APP_URL    = "https://script.google.com/macros/s/AKfycbzpUroVXu87JyY05EZ9MvF9gc1vI5ljsQ-gPgDgANIMkVwMoVxe88L7EghjFTdrn3pUxA/exec";

const ABREVIATURAS_ESTADO = {
  "BAJA CALIFORNIA": "BC", "GUANAJUATO": "GTO", "NUEVO LEÓN": "NL", "NUEVO LEON": "NL",
  "QUERÉTARO": "QRO", "QUERETARO": "QRO", "QUINTANA ROO": "Q. ROO", "SAN LUIS POTOSÍ": "SLP", "SAN LUIS POTOSI": "SLP"
};

function generarFichasPropuestasComerciales() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(3000)) {
    SpreadsheetApp.getUi().alert("⚠️ El proceso ya se está ejecutando. Por favor, espera a que termine.");
    return;
  }

  try {
    const ui = SpreadsheetApp.getUi();
    const respuesta = ui.alert(
      'Confirmación de generación',
      '¿Deseas procesar la tabla? Las REFs con ficha vinculada serán ignoradas por completo.',
      ui.ButtonSet.YES_NO
    );

    if (respuesta !== ui.Button.YES) return;

    const ssActive = SpreadsheetApp.getActiveSpreadsheet();
    const sheetPropCom = ssActive.getActiveSheet();
    let rawData = sheetPropCom.getDataRange().getDisplayValues();

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

    // Columna Plantilla
    let colPlantillaIdx = headers.indexOf("plantilla");
    if (colPlantillaIdx === -1) {
      let colMantenimiento = headers.findIndex(h => h.includes("mantenimiento"));
      let targetCol = colMantenimiento !== -1 ? colMantenimiento + 1 : headers.length;
      sheetPropCom.insertColumnAfter(targetCol);
      sheetPropCom.getRange(headerRowIdx + 1, targetCol + 1)
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

    const colTargetExplicacion = col.ficha > 0 ? (col.ficha - 1) : (col.partida > 0 ? col.partida - 1 : -1);

    // FILTRAR FILAS A PROCESAR
    const filasAProcesar = [];
    let ignoradasConLink = 0;

    for (let i = headerRowIdx + 1; i < rawData.length; i++) {
      const row = rawData[i];
      const partidaStr = col.partida !== -1 ? row[col.partida].toString().trim() : "";
      if (!/^\d+$/.test(partidaStr)) continue;

      const numFilaReal = i + 1;
      const celdaRefRango = sheetPropCom.getRange(numFilaReal, col.ref + 1);
      const richTextRef = celdaRefRango.getRichTextValue();
      const urlLink = richTextRef ? richTextRef.getLinkUrl() : null;
      const tieneLinkActivo = urlLink && urlLink.toString().trim().toLowerCase().startsWith("http");

      // IGNORAR SI YA TIENE LINK
      if (tieneLinkActivo) {
        ignoradasConLink++;
        continue;
      }

      let refTexto = celdaRefRango.getValue().toString().trim();
      const esSolicitudCrear = refTexto.toUpperCase() === "CREAR";

      if (!esSolicitudCrear && (!refTexto || refTexto.toLowerCase() === "seleccionar...")) continue;

      filasAProcesar.push({
        numFilaReal: numFilaReal,
        row: row,
        refTexto: refTexto,
        esSolicitudCrear: esSolicitudCrear,
        celdaRefRango: celdaRefRango,
        celdaPlantillaRango: sheetPropCom.getRange(numFilaReal, col.plantilla + 1),
        celdaStatusRango: colTargetExplicacion !== -1 ? sheetPropCom.getRange(numFilaReal, colTargetExplicacion + 1) : null
      });
    }

    if (filasAProcesar.length === 0) {
      SpreadsheetApp.getUi().alert(`🎉 Proceso finalizado.\n\n🔒 REFs con ficha previa omitidas: ${ignoradasConLink}\nNo hay filas pendientes por procesar.`);
      return;
    }

    // 1. OBTENER DATOS DEL INVENTARIO (HERMANAS E ÍNDICE DE PLANTILLAS PRE-CALCULADO)
    const datosHermanas = obtenerMapaHermanasWebAPP();
    const mapaDuo = datosHermanas.mapaDuo || {};
    const mapaParqueSolo = datosHermanas.mapaParqueSolo || {};
    
    // Solicitamos los modelos ya indexados desde el Inventario (sin recorrer Drive)
    const cacheModelos = obtenerIndicePlantillasWebAPP();

    let procesadasConFicha = 0, refsNuevasConFicha = 0, refsNuevasSinFicha = 0, existSinFicha = 0, errores = 0;

    for (const item of filasAProcesar) {
      const row = item.row;
      const partidaStr = col.partida !== -1 ? row[col.partida].toString().trim() : "";
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
      const operacionLimpia = operacionStr.toUpperCase().includes("VENTA") ? "VENTA" : "RENTA";

      const numSugM2 = parseFloat(String(m2Sugerida).replace(/[^\d.]/g, "")) || 0;

      if (!m2Sugerida || numSugM2 === 0 || !precioM2) {
        errores++;
        if (item.celdaStatusRango) item.celdaStatusRango.setValue("⚠️ Faltan datos M2/Precio");
        continue;
      }

      // 2. BUSCAR MODELO EN MEMORIA (INSTANTÁNEO)
      let modeloEncontrado = null;
      let refCoincidente = "";
      let esPorHermana = false;

      const tieneRefExistente = item.refTexto && !item.esSolicitudCrear;

      if (tieneRefExistente) {
        const refLimpia = extraerCodigoRefLimpio(item.refTexto);
        modeloEncontrado = cacheModelos.find(m => {
          if (!m.refsEncontradas) return false;
          if (Array.isArray(m.refsEncontradas)) {
            return m.refsEncontradas.some(r => String(r).trim().toUpperCase() === refLimpia);
          } else {
            return String(m.refsEncontradas).trim().toUpperCase().split(",").map(s => s.trim()).includes(refLimpia);
          }
        });
        if (modeloEncontrado) refCoincidente = refLimpia;
      }

      if (!modeloEncontrado && parqStr) {
        const kParque = `${limpiarTextoClave(parqStr)}|${operacionLimpia}`;
        const kDuo = devStr ? `${limpiarTextoClave(parqStr)}|${limpiarTextoClave(devStr)}|${operacionLimpia}` : kParque;
        let hermanasInv = mapaDuo[kDuo] || mapaParqueSolo[kParque] || [];

        for (const refHermana of hermanasInv) {
          const refHermanaLimpia = extraerCodigoRefLimpio(refHermana);
          modeloEncontrado = cacheModelos.find(m => {
            if (!m.refsEncontradas) return false;
            if (Array.isArray(m.refsEncontradas)) {
              return m.refsEncontradas.some(r => String(r).trim().toUpperCase() === refHermanaLimpia);
            } else {
              return String(m.refsEncontradas).trim().toUpperCase().split(",").map(s => s.trim()).includes(refHermanaLimpia);
            }
          });
          if (modeloEncontrado) {
            refCoincidente = refHermanaLimpia;
            esPorHermana = true;
            break;
          }
        }
      }

      // Escribir celda Plantilla
      if (modeloEncontrado) {
        const slideUrl = `https://docs.google.com/presentation/d/${modeloEncontrado.fileId}/edit#slide=id.${modeloEncontrado.slideObjectId || ""}`;
        const labelPlantilla = esPorHermana ? `Sí (Hermana: ${refCoincidente})` : `Sí (${refCoincidente})`;
        
        const richValuePlantilla = SpreadsheetApp.newRichTextValue()
          .setText(labelPlantilla)
          .setLinkUrl(0, labelPlantilla.length, slideUrl)
          .build();

        item.celdaPlantillaRango
          .setRichTextValue(richValuePlantilla)
          .setFontColor("#1155cc")
          .setFontLine("underline")
          .setFontWeight("bold")
          .setHorizontalAlignment("center");
      } else {
        item.celdaPlantillaRango
          .setValue("No")
          .setFontColor("#cc0000")
          .setFontLine("none")
          .setFontWeight("normal")
          .setHorizontalAlignment("center");
      }

      // 3. VALIDAR O CREAR REF
      let refFinal = "";
      let esRefRecienCreada = false;
      let urlPdfExistente = "";

      if (tieneRefExistente) {
        refFinal = item.refTexto;
        const refLimpia = extraerCodigoRefLimpio(item.refTexto);
        const datosRef = consultarDatosRefWebAPP(refLimpia);

        if (datosRef.status === "success" && datosRef.encontrado) {
          if (datosRef.m2Original > 0 && datosRef.m2Original !== numSugM2) {
            errores++;
            const msgError = `❌ Discrepancia M2: Hoja (${numSugM2} m²) vs Inventario (${datosRef.m2Original} m²)`;
            if (item.celdaStatusRango) item.celdaStatusRango.setValue(msgError);
            continue;
          }
          
          // 1. Intentar obtener PDF desde la WebApp
          const posibleUrl = datosRef.ficha || datosRef.urlPdf || datosRef.linkFicha || datosRef.link || "";
          if (posibleUrl && posibleUrl.toString().trim().startsWith("http")) {
            urlPdfExistente = posibleUrl.toString().trim();
          }
        }

        // 2. RESPALDO DIRECTO EN DRIVE: Si la WebApp no dio URL, buscamos si ya existe el PDF en la carpeta del Inventario
        if (!urlPdfExistente) {
          try {
            const folderOperacionInv = obtenerCarpetaOperacionInventario(operacionStr);
            const folderEstadoInv = obtenerOCrearSubcarpeta(folderOperacionInv, estadoStr.toUpperCase());
            const archivos = folderEstadoInv.getFiles();

            while (archivos.hasNext()) {
              const file = archivos.next();
              const nombreArch = file.getName().toUpperCase();
              // Verifica que sea el PDF de esa REF exacta y que no sea un "No disponible"
              if (nombreArch.endsWith(`${refLimpia}.PDF`) && !nombreArch.includes("NO DISPONIBLE")) {
                urlPdfExistente = file.getUrl();
                // Aprovechamos y actualizamos la WebApp para que ya lo sepa
                actualizarFichaEnInventarioWebAPP(refLimpia, urlPdfExistente);
                break;
              }
            }
          } catch (errDrive) {
            // Si falla la búsqueda en Drive, continuará con la generación
          }
        }

      } else if (item.esSolicitudCrear) {
        const resCrear = solicitarNuevaRefWebAPP(parqStr, numSugM2, operacionStr, precioM2);
        
        if (resCrear.status !== "success" || !resCrear.nuevaRef) {
          errores++;
          const msgErr = resCrear.message ? `❌ ${resCrear.message}` : "❌ Error WebApp al crear REF";
          if (item.celdaStatusRango) item.celdaStatusRango.setValue(msgErr);
          continue;
        }
        refFinal = resCrear.nuevaRef;
        esRefRecienCreada = true;
      }

      // 4. VINCULAR O GENERAR PDF DE LA FICHA
      
      // CASO A: REUTILIZAR FICHA EXISTENTE (WEBAPP O ENCONTRADA EN DRIVE)
      if (urlPdfExistente !== "") {
        const richValueRef = SpreadsheetApp.newRichTextValue()
          .setText(refFinal)
          .setLinkUrl(0, refFinal.length, urlPdfExistente)
          .build();

        item.celdaRefRango.clearDataValidations();
        item.celdaRefRango.setRichTextValue(richValueRef);
        item.celdaRefRango.setFontLine("underline").setFontColor("#1155cc");

        procesadasConFicha++;
        if (item.celdaStatusRango) item.celdaStatusRango.setValue("✅ Ficha recuperada del Inventario");
        continue;
      }

      // CASO B: NO EXISTE PDF, GENERAR UNO NUEVO
      if (modeloEncontrado) {
        let pdfGeneradoExito = false;
        let detalleError = "";

        try {
          const zonaSubzonaLimpia = (subzonaStr || zonaStr).toUpperCase();
          const estadoAbreviado = obtenerEstadoFormateado(estadoStr);
          const m2EnteroSinSeparador = Math.round(numSugM2).toString();
          const nombreInventario = `${operacionStr.toUpperCase()} ${m2EnteroSinSeparador} M2 ${zonaSubzonaLimpia} ${estadoAbreviado} ${refFinal}`.toUpperCase();

          const folderOperacionInv = obtenerCarpetaOperacionInventario(operacionStr);
          const folderEstadoInv = obtenerOCrearSubcarpeta(folderOperacionInv, estadoStr.toUpperCase());

          const pdfBlob = generarPdfDesdeModelo(
            modeloEncontrado, partidaStr, m2Sugerida, precioM2, precioTotal,
            refFinal, cacheModelos
          );

          pdfBlob.setName(nombreInventario + ".pdf");
          const pdfFileInv = folderEstadoInv.createFile(pdfBlob);
          const urlPdfInv = pdfFileInv.getUrl();

          if (urlPdfInv) {
            const richValueRef = SpreadsheetApp.newRichTextValue()
              .setText(refFinal)
              .setLinkUrl(0, refFinal.length, urlPdfInv)
              .build();

            item.celdaRefRango.clearDataValidations();
            item.celdaRefRango.setRichTextValue(richValueRef);
            item.celdaRefRango.setFontLine("underline").setFontColor("#1155cc");

            actualizarFichaEnInventarioWebAPP(refFinal, urlPdfInv);
            pdfGeneradoExito = true;
          }

        } catch (err) {
          detalleError = err.toString();
        }

        if (pdfGeneradoExito) {
          procesadasConFicha++;
          if (esRefRecienCreada) {
            refsNuevasConFicha++;
            if (item.celdaStatusRango) item.celdaStatusRango.setValue("✅ REF creada en Inventario con FICHA");
          } else {
            if (item.celdaStatusRango) item.celdaStatusRango.setValue("✅ Ficha asignada correctamente");
          }
        } else {
          errores++;
          if (item.celdaStatusRango) item.celdaStatusRango.setValue(`❌ Error PDF: ${detalleError.substring(0, 30)}`);
        }

      } else {
        if (esRefRecienCreada) {
          item.celdaRefRango.clearDataValidations();
          item.celdaRefRango.setValue(refFinal);
          item.celdaRefRango.setFontLine("none").setFontColor("#000000");
          refsNuevasSinFicha++;
          if (item.celdaStatusRango) item.celdaStatusRango.setValue("ℹ️ REF creada en Inventario (Sin Ficha PDF)");
        } else {
          existSinFicha++;
          if (item.celdaStatusRango) item.celdaStatusRango.setValue("⚠️ Sin modelo de plantilla encontrado");
        }
      }
    }

    // RESUMEN
    SpreadsheetApp.getUi().alert(
      `🎉 Proceso finalizado.\n\n` +
      `🆔 Nuevas REFs creadas: ${refsNuevasConFicha + refsNuevasSinFicha}\n` +
      `📄 Fichas PDF generadas: ${procesadasConFicha}\n` +
      `🔒 REFs con ficha previa omitidas: ${ignoradasConLink}\n` +
      `❌ Errores / Discrepancias: ${errores}`
    );

  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// PETICIONES A LA WEBAPP (SIN RECORRER DRIVE)
// ==========================================

function obtenerMapaHermanasWebAPP() {
  try {
    const url = INVENTARIO_WEB_APP_URL + "?action=obtenerMapaHermanas";
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const json = JSON.parse(response.getContentText());
    return json.status === "success" ? json : { mapaDuo: {}, mapaParqueSolo: {} };
  } catch (e) {
    return { mapaDuo: {}, mapaParqueSolo: {} };
  }
}

function obtenerIndicePlantillasWebAPP() {
  try {
    const url = INVENTARIO_WEB_APP_URL + "?action=obtenerIndicePlantillas";
    const response = UrlFetchApp.fetch(url, { 
      muteHttpExceptions: true,
      followRedirects: true 
    });
    
    const responseText = response.getContentText();
    const json = JSON.parse(responseText);

    if (Array.isArray(json)) return json;
    if (json.modelos && Array.isArray(json.modelos)) return json.modelos;
    if (json.data && Array.isArray(json.data)) return json.data;

    return [];
  } catch (e) {
    return [];
  }
}

function consultarDatosRefWebAPP(ref) {
  try {
    const url = INVENTARIO_WEB_APP_URL + "?action=obtenerDatosRef&ref=" + encodeURIComponent(ref);
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    return JSON.parse(response.getContentText());
  } catch (e) {
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
    return { status: "error", message: e.toString() };
  }
}

function actualizarFichaEnInventarioWebAPP(ref, urlPdf) {
  try {
    const url = INVENTARIO_WEB_APP_URL 
      + "?action=actualizarFicha"
      + "&ref=" + encodeURIComponent(ref)
      + "&urlPdf=" + encodeURIComponent(urlPdf);

    UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  } catch (e) {}
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

function obtenerCarpetaOperacionInventario(operacion) {
  const folderMaestra = DriveApp.getFolderById(FOLDER_MAESTRA_INV_ID);
  const opLimpia = (operacion || "").toUpperCase();
  const nombreSub = opLimpia.includes("VENTA") ? "2. VENTA NAVES" : "1. RENTA NAVES";
  return obtenerOCrearSubcarpeta(folderMaestra, nombreSub);
}

function obtenerOCrearSubcarpeta(parentFolder, nombreSubcarpeta) {
  if (!parentFolder) return null;
  const nombreLimpio = (nombreSubcarpeta && nombreSubcarpeta.trim() !== "") ? nombreSubcarpeta.trim() : "EDOMEX";
  const folders = parentFolder.getFoldersByName(nombreLimpio);
  return folders.hasNext() ? folders.next() : parentFolder.createFolder(nombreLimpio);
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
  Utilities.sleep(300);

  const pdfBlob = copiaFile.getAs(MimeType.PDF);
  DriveApp.getFileById(copiaFile.getId()).setTrashed(true);

  return pdfBlob;
}
