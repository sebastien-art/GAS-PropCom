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
    const cache = CacheService.getScriptCache();
    const target = ref.toString().trim().toUpperCase();

    // 1. Conexión rápida
    const ss = SpreadsheetApp.openById(id);
    const sh = ss.getSheetByName("Naves");
    
    // 2. Traemos TODA la columna REF y los Encabezados de un solo golpe (Esto es lo que tarda)
    // Usamos getValues() una sola vez para ser más eficientes
    const dataRange = sh.getDataRange().getValues();
    const headers = dataRange[0];
    const colRefIdx = headers.indexOf("REF");

    if (colRefIdx === -1) return "Error: No se encontró columna REF";

    // 3. Buscamos la fila en el array (esto es instantáneo en memoria)
    // Empezamos desde i=1 para saltar encabezados
    let rowIndex = -1;
    for (let i = 1; i < dataRange.length; i++) {
      if (dataRange[i][colRefIdx].toString().trim().toUpperCase() === target) {
        rowIndex = i;
        break;
      }
    }

    if (rowIndex === -1) return "No se encontró la REF: " + target;

    // 4. Extraemos los datos de esa fila
    const rowData = dataRange[rowIndex];

    // 5. Mapeo de campos
    const campos = [
      "Intermediario","Operación","Ficha","REF","Teléfono o link","Nombre","Estado","Zona Principal","Sub Zona","Desarrollador",
      "Parque","Nave","M2 de construcción","M2 de terreno","M2 mínimos rentables","Asking price /m2","Mantenimiento / m2",
      "Energía (kVAs)","Disponibilidad","Comentarios","Renta total","Mantenimiento total",
      "Coordenadas","Ubicación","Andenes de carga","Rampas","A piso", "Resistencia de piso (espesor, resistencia tonelada por m2)","Altura libre",
      "Altura máxima","Tipo de construcción","Tipo de techo","% Skylight",
      "Seguridad 24/7","Oficinas (m2 o %)","Moneda del contrato","Año de construcción",
      "Protección contra incendios","Plazo mínimo de contrato","Gas natural","Caseta de seguridad privada","ID de carpeta de fotos"
    ];

    let res = {};
    campos.forEach(c => {
      let idx = headers.indexOf(c);
      if (idx !== -1) {
        let valor = rowData[idx];
        res[c] = (valor !== "" && valor !== null) ? valor : "---";
      } else {
        res[c] = "---";
      }
    });

    return JSON.stringify(res);
    
  } catch (err) {
    return "Fallo en servidor: " + err.message;
  }
}