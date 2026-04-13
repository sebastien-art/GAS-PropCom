/**
 * Lanza el Sidebar usando el archivo HTML externo (SidebarHTML.html).
 */
function showNaveSidebar() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('SidebarHTML')
      .setTitle('Buscador de Naves IEM')
      .setWidth(350);
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (e) {
    SpreadsheetApp.getUi().alert("Error al abrir el panel: " + e.toString());
  }
}

/**
 * Limpia el caché de la hoja Naves.
 */
function limpiarCacheNaves() {
  const cache = CacheService.getScriptCache();
  cache.remove("nave_headers");
  cache.remove("nave_refcol");
  cache.remove("nave_refs");
}

/**
 * Se dispara automáticamente al editar la hoja.
 * Si el cambio fue en "Naves", invalida el caché.
 */
function onEdit(e) {
  const hoja = e.source.getActiveSheet();
  if (hoja.getName() === "Naves") {
    limpiarCacheNaves();
  }
}

function buscarNaveEnServidor(ref) {
  if (!ref) return "Error: No escribiste nada";
  try {
    const id = "1jHh3SUkVrQtOZPQ2T2iOhcXF_FbjdXivt3uuBYfYHz0";
    const ss = SpreadsheetApp.openById(id);
    const sh = ss.getSheetByName("Naves");
    const dataRange = sh.getDataRange().getValues();
    const richTextValues = sh.getDataRange().getRichTextValues(); // Para el link
    const headers = dataRange[0];
    const colRefIdx = headers.indexOf("REF");
    const colFichaIdx = headers.indexOf("Ficha");

    let rowIndex = -1;
    for (let i = 1; i < dataRange.length; i++) {
      if (dataRange[i][colRefIdx].toString().trim().toUpperCase() === ref.toString().trim().toUpperCase()) {
        rowIndex = i; break;
      }
    }
    if (rowIndex === -1) return "No se encontró la REF: " + ref;

    const rowData = dataRange[rowIndex];
    const campos = ["Intermediario","Operación","Ficha","REF","Estado","Zona Principal","Sub Zona","Desarrollador","Parque","Nave","M2 de construcción","M2 de terreno","M2 mínimos rentables","Asking price /m2","Mantenimiento / m2","Energía (kVAs)","Disponibilidad","Comentarios","Renta total","Mantenimiento total","Coordenadas","Ubicación","Andenes de carga","Rampas","A piso", "Resistencia de piso (espesor, resistencia tonelada por m2)","Altura libre","Altura máxima","Tipo de construcción","Tipo de techo","% Skylight","Seguridad 24/7","Oficinas (m2 o %)","Moneda del contrato","Año de construcción","Protección contra incendios","Plazo mínimo de contrato","Gas natural","Caseta de seguridad privada","ID de carpeta de fotos"];

    // IMPORTANTE: Devolvemos un objeto con "datos" y "link"
    let respuestaFinal = { "datos": {}, "linkFicha": "" };
    
    if (colFichaIdx !== -1) {
      respuestaFinal.linkFicha = richTextValues[rowIndex][colFichaIdx].getLinkUrl() || "";
    }

    campos.forEach(c => {
      let idx = headers.indexOf(c);
      respuestaFinal.datos[c] = (idx !== -1 && rowData[idx] !== "" && rowData[idx] !== null) ? rowData[idx] : "---";
    });

    return JSON.stringify(respuestaFinal);
  } catch (err) { return "Error: " + err.message; }
}


/**
 * NUEVA FUNCIÓN: Extrae el link real de la celda Ficha y lo devuelve al Sidebar.
 */
function abrirFichaDesdeSidebar(ref) {
  try {
    const id = "1jHh3SUkVrQtOZPQ2T2iOhcXF_FbjdXivt3uuBYfYHz0";
    const sh = SpreadsheetApp.openById(id).getSheetByName("Naves");
    const data = sh.getDataRange().getValues();
    const richText = sh.getDataRange().getRichTextValues();
    const headers = data[0];
    const colRefIdx = headers.indexOf("REF");
    const colFichaIdx = headers.indexOf("Ficha");

    const target = ref.toString().trim().toUpperCase();

    for (let i = 1; i < data.length; i++) {
      if (data[i][colRefIdx].toString().trim().toUpperCase() === target) {
        // Obtenemos el link real del RichText para que no importe si dice "OK" o "VER"
        let url = richText[i][colFichaIdx].getLinkUrl();
        if (url) return url;
        break;
      }
    }
    return null;
  } catch (e) {
    return null;
  }
}