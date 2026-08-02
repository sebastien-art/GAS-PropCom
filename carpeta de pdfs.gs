/**
 * PREPARAR CARPETA Y ZIP DE CLIENTE
 * - Blindado: Solo lee filas dentro de la tabla que tengan un número de PARTIDA válido.
 * - Ignora notas, mapas o enlaces ubicados debajo de la tabla.
 * - Formato final: "XX - REF - Resto del título.pdf"
 * - Coloca hipervínculos en la COLUMNA D.
 */

const FOLDER_ID_PROPUESTAS_CLIENTES = "1YG1LMk8D0zYWib7YqnH9q3XudZZY4_cY".trim();

function prepararCliente() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getActiveSheet();
  const nombreCliente = hoja.getName().trim();

  try {
    const data = hoja.getDataRange().getValues();
    const richTextValues = hoja.getDataRange().getRichTextValues();
    
    // 1. LOCALIZAR LA FILA DE ENCABEZADOS ("Partida" Y "REF")
    let filaEncabezado = -1;
    let idxPartida = -1;
    let idxRef = -1;

    for (let r = 0; r < data.length; r++) {
      for (let c = 0; c < data[r].length; c++) {
        const val = String(data[r][c]).trim().toLowerCase();
        if (val === "partida") idxPartida = c;
        if (val === "ref") idxRef = c;
      }
      if (idxPartida !== -1 && idxRef !== -1) {
        filaEncabezado = r;
        break;
      }
    }

    if (idxRef === -1 || idxPartida === -1) {
      ui.alert("⚠️ No se encontraron los encabezados 'Partida' o 'REF' en la hoja.");
      return;
    }

    // 2. EXTRAER ÚNICAMENTE LAS FILAS DE LA TABLA QUE TENGAN NÚMERO DE PARTIDA
    const elementosAProcesar = [];

    for (let r = filaEncabezado + 1; r < data.length; r++) {
      const valPartidaRaw = data[r][idxPartida];
      const strPartida = String(valPartidaRaw).trim();

      // REGLE DE BLINDAJE: Si la columna Partida no es un número entero (ej. está vacía, o es nota/mapa), 
      // asumimos que terminó la tabla principal o es una fila no válida.
      if (!/^\d+$/.test(strPartida)) {
        // Si encontramos texto que no es número después de la tabla, detenemos la lectura
        if (strPartida !== "" && elementosAProcesar.length > 0) {
          break; 
        }
        continue;
      }

      // Validar si la celda REF tiene hipervínculo activo
      const celdaRichText = richTextValues[r][idxRef];
      if (!celdaRichText) continue;

      const textoRef = celdaRichText.getText().trim();
      if (!textoRef) continue;

      let linkUrl = celdaRichText.getLinkUrl();
      if (!linkUrl && celdaRichText.getRuns) {
        const runs = celdaRichText.getRuns();
        for (let run of runs) {
          if (run.getLinkUrl()) {
            linkUrl = run.getLinkUrl();
            break;
          }
        }
      }

      // SOLO AGREGAR SI LA REF TIENE UN LINK ACTIVO
      if (linkUrl) {
        const numPartidaFormatted = strPartida.padStart(2, '0'); // Convierte "1" en "01"

        elementosAProcesar.push({
          numPartida: numPartidaFormatted,
          refTexto: textoRef,
          linkUrl: linkUrl
        });
      }
    }

    if (elementosAProcesar.length === 0) {
      ui.alert("⚠️ No se encontraron filas con número de Partida y enlace activo en REF.");
      return;
    }

    // 3. CREAR O LIMPIAR SUBCARPETA DEL CLIENTE EN DRIVE
    const fechaHoy = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "dd-MM-yy");
    const carpetaCliente = obtenerOCrearCarpetaClienteSameDay_(nombreCliente, fechaHoy);

    let contador = 0;
    const archivosGuardados = [];

    // 4. DESCARGAR Y RENOMBRAR CON EL FORMATO "XX - REF - RESTO DEL TÍTULO"
    elementosAProcesar.forEach((item) => {
      const itemId = extraerIdFromUrl_(item.linkUrl);
      if (!itemId) return;

      try {
        const resultado = obtenerPdfYNombreOriginal_(itemId, item.linkUrl);

        if (resultado && resultado.blob) {
          const nombreNuevo = armarNombreFinalLimpio_(item.numPartida, item.refTexto, resultado.nombreOriginal);
          
          const blobFinal = resultado.blob.setName(nombreNuevo);
          const archivoCreado = carpetaCliente.createFile(blobFinal);
          
          archivosGuardados.push(archivoCreado);
          contador++;
        }
      } catch (eFile) {
        console.warn(`No se pudo procesar REF ${item.refTexto}: ${eFile.message}`);
      }
    });

    if (archivosGuardados.length === 0) {
      ui.alert("⚠️ No se pudieron obtener los archivos de las fichas.");
      return;
    }

    // 5. CREAR EL ARCHIVO .ZIP
    const blobsParaZip = archivosGuardados.map(file => file.getBlob());
    const nombreZip = `Fichas_${nombreCliente}_${fechaHoy}.zip`;
    const archivoZipBlob = Utilities.zip(blobsParaZip, nombreZip);
    const archivoZipDrive = carpetaCliente.createFile(archivoZipBlob);

    // 6. ESCRIBIR HIPERVÍNCULOS EN LA COLUMNA D (COLUMNA 4)
    const ultimaFila = hoja.getLastRow();
    const filaBase = ultimaFila + 2;

    const urlCarpeta = carpetaCliente.getUrl();
    const urlZip = archivoZipDrive.getDownloadUrl();

    // Link a la Carpeta de Drive
    const celdaCarpeta = hoja.getRange(filaBase, 4);
    celdaCarpeta.setFormula(`=HYPERLINK("${urlCarpeta}", "📁 Carpeta de Fichas (${nombreCliente} - ${fechaHoy})")`);
    celdaCarpeta.setFontWeight("bold").setFontColor("#1155cc");

    // Link al Archivo ZIP
    const celdaZip = hoja.getRange(filaBase + 1, 4);
    celdaZip.setFormula(`=HYPERLINK("${urlZip}", "📦 Descargar Fichas (.ZIP)")`);
    celdaZip.setFontWeight("bold").setFontColor("#28a745");

    // 7. VENTANA EMERGENTE DE CONFIRMACIÓN
    mostrarVentanaDescarga_(urlCarpeta, urlZip, nombreCliente, contador);

  } catch (error) {
    ui.alert("❌ Error procesando carpeta:\n" + String(error && error.message ? error.message : error));
  }
}

/**
 * LIMPIEZA PROFUNDA Y FORMATO: "XX - REF - Resto del título sin basura.pdf"
 */
function armarNombreFinalLimpio_(numPartida, refTexto, nombreOriginal) {
  let limpio = nombreOriginal.replace(/\.pdf$/i, '').trim();

  // Quitar la palabra "Ficha" si viene al inicio
  limpio = limpio.replace(/^Ficha\s*-\s*/i, '').replace(/^Ficha\s*/i, '').trim();

  // Quitar la REF para no duplicarla
  const regexRef = new RegExp(refTexto, "gi");
  limpio = limpio.replace(regexRef, '').trim();

  // Eliminar guiones dobles, números raros pegados o basuras intermedios
  limpio = limpio
    .replace(/^[\s\-_0-9]+/, '')  // Quita números o guiones basuras al inicio
    .replace(/[\s\-_]+/g, ' ')    // Convierte secuencias de guiones/espacios en un solo espacio
    .trim();

  const prefijoPartida = numPartida ? `${numPartida} - ` : '';

  if (!limpio) {
    return `${prefijoPartida}${refTexto}.pdf`;
  }

  return `${prefijoPartida}${refTexto} - ${limpio}.pdf`;
}

// FUNCIONES AUXILIARES DRIVE Y VENTANA EMERGENTE
function obtenerPdfYNombreOriginal_(itemId, linkUrl) {
  if (linkUrl.includes("/folders/")) {
    const query = `'${itemId}' in parents and mimeType = 'application/pdf' and trashed = false`;
    const busqueda = DriveApp.searchFiles(query);
    if (busqueda.hasNext()) {
      const file = busqueda.next();
      return { blob: file.getBlob(), nombreOriginal: file.getName() };
    }
  }

  try {
    const file = DriveApp.getFileById(itemId);
    return { blob: file.getBlob(), nombreOriginal: file.getName() };
  } catch (e) {}

  try {
    const downloadUrl = `https://drive.google.com/uc?export=download&id=${itemId}`;
    const response = UrlFetchApp.fetch(downloadUrl, {
      headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    });

    if (response.getResponseCode() === 200) {
      return { blob: response.getBlob(), nombreOriginal: `Ficha ${itemId}` };
    }
  } catch (e) {}

  return null;
}

function obtenerOCrearCarpetaClienteSameDay_(nombreCliente, fechaHoy) {
  let carpetaPadre = DriveApp.getFolderById(FOLDER_ID_PROPUESTAS_CLIENTES);
  const nombreCarpeta = `${nombreCliente} - ${fechaHoy}`;
  const subcarpetas = carpetaPadre.getFoldersByName(nombreCarpeta);

  if (subcarpetas.hasNext()) {
    const carpetaExistente = subcarpetas.next();
    const archivos = carpetaExistente.getFiles();
    while (archivos.hasNext()) {
      archivos.next().setTrashed(true);
    }
    return carpetaExistente;
  } else {
    return carpetaPadre.createFolder(nombreCarpeta);
  }
}

function extraerIdFromUrl_(url) {
  if (!url) return null;
  const match = url.match(/[-\w]{25,}/);
  return match ? match[0] : null;
}

function mostrarVentanaDescarga_(urlCarpeta, urlZip, nombreCliente, total) {
  const htmlContent = `
    <div style="font-family: Arial, sans-serif; padding: 10px; text-align: center;">
      <h3 style="color: #2b579a; margin-top:0;">🎉 ¡Proceso Finalizado!</h3>
      <p style="font-size: 14px; color: #333;">
        Se empaquetaron <b>${total} fichas</b> para <b>${nombreCliente}</b>.
      </p>
      <hr style="border: 0; border-top: 1px solid #ddd; margin: 15px 0;">
      
      <div style="margin-bottom: 12px;">
        <a href="${urlZip}" target="_blank" style="
          background-color: #28a745;
          color: white;
          padding: 10px 18px;
          text-decoration: none;
          font-weight: bold;
          border-radius: 5px;
          display: inline-block;
          width: 80%;">
          📦 Descargar Archivo .ZIP
        </a>
      </div>

      <div>
        <a href="${urlCarpeta}" target="_blank" style="
          background-color: #007bff;
          color: white;
          padding: 10px 18px;
          text-decoration: none;
          font-weight: bold;
          border-radius: 5px;
          display: inline-block;
          width: 80%;">
          📂 Abrir Carpeta en Google Drive
        </a>
      </div>
    </div>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(360)
    .setHeight(230);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Descargar Propuesta Comercial');
}