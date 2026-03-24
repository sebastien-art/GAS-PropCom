/***************************************************************
 * MAPA CLIENTE (IMAGEN + PANEL LATERAL) — FIX NO SOBRESCRIBE
 *
 * ✅ Qué hace:
 *  - Lee la TABLA FINAL detectando encabezados AZULES (#9fc5e8)
 *  - Usa TODOS los renglones por "Partida" (ES) o "Item / 项目" (ENG/CHI) o "Item"
 *  - Usa la columna "Coordenadas" (ES) o "Coordinates" (ENG/CHI)
 *  - Genera imagen (Static Map) usando Maps Service
 *  - Guarda PNG SIEMPRE en: OUTPUT_ROOT_FOLDER_ID / MAPS_FOLDER_NAME
 *  - Nombre: MAP_FILE_PREFIX + {NombrePestaña} + ".png"
 *  - Debajo deja link al PNG, SIN SOBRESCRIBIR texto existente
 *  - Panel lateral con mapa interactivo (Google Maps JS)
 *
 * Requiere en 00_Config.gs:
 *  const OUTPUT_ROOT_FOLDER_ID = "...";
 *  const MAPS_FOLDER_NAME = "Mapas Propuestas";
 *  const MAP_FILE_PREFIX = "MapaCliente_";
 ***************************************************************/

/*************** 0) CONFIG (EJECUTA 1 VEZ) *********************/
function ConfigurarGoogleMapsKey() {
  // ⚠️ Si ya la configuraste antes, NO necesitas volver a ejecutar esta función.
  const API_KEY = "REEMPLAZA_CON_TU_API_KEY";
  PropertiesService.getScriptProperties().setProperty("GMAPS_KEY", API_KEY);
}

/*************** 1) GENERAR MAPA (IMAGEN) **********************/
function GenerarMapaCliente() {
  const HEADER_BG = "#9fc5e8"; // light blue 3
  const MAPTYPE = "roadmap";

  const API_KEY = PropertiesService.getScriptProperties().getProperty("GMAPS_KEY");
  if (!API_KEY) throw new Error("No encontré GMAPS_KEY. Ejecuta ConfigurarGoogleMapsKey() primero.");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  // 1) Lee tabla final por encabezados azules
  const table = getFinalTableMeta_(sheet, HEADER_BG);
  const points = readFinalTablePoints_(sheet, table);
  if (!points.valid.length) throw new Error("No hay coordenadas válidas en la tabla final.");

  // 2) Construye Static Map con Maps Service
  const bounds = getBounds_(points.valid);
  const center = { lat: (bounds.minLat + bounds.maxLat) / 2, lng: (bounds.minLng + bounds.maxLng) / 2 };
  const zoom = (points.valid.length === 1) ? 14 : computeZoomForBounds_(bounds, 640, 640);

  const map = Maps.newStaticMap()
    .setSize(640, 640)
    .setCenter(center.lat, center.lng)
    .setZoom(zoom)
    .setMapType(
      MAPTYPE === "satellite" ? Maps.StaticMap.Type.SATELLITE :
      MAPTYPE === "terrain"   ? Maps.StaticMap.Type.TERRAIN :
      MAPTYPE === "hybrid"    ? Maps.StaticMap.Type.HYBRID :
                                Maps.StaticMap.Type.ROADMAP
    )
    .setFormat(Maps.StaticMap.Format.PNG);

  // Primero agrega los >9 sin label. Luego 1-9 con label (Static Maps: label 1 caracter)
  const labeled = [];
  const unlabeled = [];

  points.valid.forEach(p => {
    if (p.partida >= 1 && p.partida <= 9) labeled.push(p);
    else unlabeled.push(p);
  });

  unlabeled.forEach(p => map.addMarker(p.lat, p.lng));

  labeled.sort((a, b) => a.partida - b.partida).forEach(p => {
    map.setMarkerStyle(
      Maps.StaticMap.MarkerSize.MID,
      Maps.StaticMap.Color.RED,
      String(p.partida)
    );
    map.addMarker(p.lat, p.lng);
  });

  // 3) Blob PNG
  const blob = map.getBlob().setName("MapaCliente.png");

  // 4) Guardar PNG SIEMPRE en ROOT / MAPS_FOLDER_NAME
  const folder = getOrCreateMapsFolderInRoot_();
  const fileName = `${MAP_FILE_PREFIX}${sheet.getName()}.png`;

  trashFilesByName_(folder, fileName);
  const file = folder.createFile(blob).setName(fileName);

  // 5) Salida debajo de la tabla (SIN SOBRESCRIBIR)
  const outCol = table.colPartida;

  // ✅ Limpia salidas previas "Mapa (Drive):" (para no duplicar) pero SIN tocar otras notas
  clearPreviousMapOutput_(sheet, outCol, table.headerRow);

  // ✅ Encuentra el último renglón con contenido (en un bloque ancho) y escribe abajo
  const lastContentRow = findLastNonEmptyRowInBand_(sheet, table.headerRow, outCol, 12);
  const outRow = lastContentRow + 2;

  // Asegura filas suficientes
  const needEndRow = outRow + 3;
  const maxRows = sheet.getMaxRows();
  if (needEndRow > maxRows) sheet.insertRowsAfter(maxRows, needEndRow - maxRows);

  // Limpia SOLO el área exacta donde escribiremos (3 filas x 6 cols)
  sheet.getRange(outRow, outCol, 3, 6).clearContent();

  sheet.getRange(outRow, outCol).setValue("Mapa (Drive):").setFontWeight("bold");
  sheet.getRange(outRow, outCol + 1)
    .setFormula(`=HYPERLINK("https://drive.google.com/file/d/${file.getId()}/view","Abrir imagen del mapa")`);

  sheet.getRange(outRow + 1, outCol).setValue('Para ajustar imagen, usar menú "Abrir mapa en panel lateral"');

  sheet.getRange(outRow + 2, outCol).setValue("Actualizado:").setFontWeight("bold");
  const tz = ss.getSpreadsheetTimeZone();
  sheet.getRange(outRow + 2, outCol + 1)
    .setValue(Utilities.formatDate(new Date(), tz, "dd/MM/yy"));

  // Reporte inválidos (si aplica)
  if (points.invalid.length) {
    const warnStart = outRow + 3 + 42;
    const needWarnEnd = warnStart + 2 + points.invalid.length;
    const maxRows2 = sheet.getMaxRows();
    if (needWarnEnd > maxRows2) sheet.insertRowsAfter(maxRows2, needWarnEnd - maxRows2);

    sheet.getRange(warnStart, outCol).setValue("Coordenadas inválidas (no se mapearon):").setFontWeight("bold");
    sheet.getRange(warnStart + 1, outCol, 1, 3).setValues([["Partida/Item", "REF", "Coordenadas/Coordinates"]]).setFontWeight("bold");
    sheet.getRange(warnStart + 2, outCol, points.invalid.length, 3).setValues(points.invalid);
  }

  ss.toast(`Mapa guardado en ROOT/${MAPS_FOLDER_NAME}: ${fileName}`, "GenerarMapaCliente", 5);
}

/*************** 2) ABRIR MAPA EN PANEL LATERAL ****************/
function AbrirMapaClienteSidebar() {
  const API_KEY = PropertiesService.getScriptProperties().getProperty("GMAPS_KEY");
  if (!API_KEY) throw new Error("No encontré GMAPS_KEY. Ejecuta ConfigurarGoogleMapsKey() primero.");

  const sheet = SpreadsheetApp.getActiveSheet();
  const table = getFinalTableMeta_(sheet, "#9fc5e8");
  const points = readFinalTablePoints_(sheet, table);
  if (!points.valid.length) throw new Error("No hay coordenadas válidas en la tabla final.");

  const data = points.valid.sort((a, b) => a.partida - b.partida);
  const dataJson = JSON.stringify(data);

  const html =
    '<!doctype html><html><head><meta charset="utf-8"/>' +
    '<meta name="viewport" content="width=device-width, initial-scale=1"/>' +
    '<style>html,body{height:100%;margin:0;font-family:Arial,sans-serif;}#map{height:100vh;width:100%;}</style>' +
    '</head><body><div id="map"></div>' +
    '<script>' +
    'var points=' + dataJson + ';' +
    'function initMap(){' +
    '  var map=new google.maps.Map(document.getElementById("map"),{' +
    '    center:{lat:points[0].lat,lng:points[0].lng},zoom:12,mapTypeId:"roadmap"' +
    '  });' +
    '  var info=new google.maps.InfoWindow();' +
    '  var bounds=new google.maps.LatLngBounds();' +
    '  points.forEach(function(p){' +
    '    var pos={lat:p.lat,lng:p.lng}; bounds.extend(pos);' +
    '    var m=new google.maps.Marker({position:pos,map:map,label:String(p.partida)});' +
    '    m.addListener("click",function(){' +
    '      var c="<div style=\\"font-size:13px\\">"+' +
    '        "<div><b>Item/Partida:</b> "+p.partida+"</div>"+' +
    '        (p.ref?("<div><b>REF:</b> "+p.ref+"</div>"):"")+' +
    '        (p.estado?("<div><b>Estado:</b> "+p.estado+"</div>"):"")+' +
    '        (p.zona?("<div><b>Zona:</b> "+p.zona+"</div>"):"")+' +
    '        "<div style=\\"color:#666;font-size:12px;\\">"+p.lat.toFixed(6)+", "+p.lng.toFixed(6)+"</div>"+' +
    '        "</div>";' +
    '      info.setContent(c); info.open(map,m);' +
    '    });' +
    '  });' +
    '  map.fitBounds(bounds);' +
    '}' +
    '</script>' +
    '<script async defer src="https://maps.googleapis.com/maps/api/js?key=' + encodeURIComponent(API_KEY) + '&callback=initMap"></script>' +
    '</body></html>';

  SpreadsheetApp.getUi().showSidebar(HtmlService.createHtmlOutput(html).setTitle("Mapa Cliente"));
}

/*************** 3) HELPERS TABLA ******************************/
function getFinalTableMeta_(sheet, headerBgHex) {
  const lastCol = sheet.getLastColumn();

  // ✅ acepta ES y ENG/CHI
  const partidaCell =
    findBlueHeaderCellAny_(sheet, ["Partida", "Item", "Item / 项目"], headerBgHex);
  if (!partidaCell) {
    throw new Error('No encontré el encabezado "Partida" ni "Item / 项目" con fondo azul.');
  }

  const headerRow = partidaCell.getRow();
  const headerVals = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0].map(v => String(v || "").trim());
  const headerBgs  = sheet.getRange(headerRow, 1, 1, lastCol).getBackgrounds()[0].map(c => (c || "").toLowerCase());

  const colPartida =
    findHeaderColAny_(headerVals, headerBgs, ["Partida", "Item", "Item / 项目"], headerBgHex);

  const colCoord =
    findHeaderColAny_(headerVals, headerBgs, ["Coordenadas", "Coordinates"], headerBgHex);

  const colREF = findHeaderColAny_(headerVals, headerBgs, ["REF"], headerBgHex) || null;
  const colEstado = findHeaderColAny_(headerVals, headerBgs, ["Estado", "State / 州", "State"], headerBgHex) || null;

  // En ENG/CHI no hay "Zona Principal"; normalmente no existe en tabla parks/final.
  // Dejamos fallback por si existe.
  const colZona = findHeaderColAny_(headerVals, headerBgs, ["Zona Principal", "Main zone / 主区域", "Main zone"], headerBgHex) || null;

  if (!colPartida) throw new Error('No encontré "Partida" o "Item / 项目" con azul en la fila de encabezados.');
  if (!colCoord)   throw new Error('No encontré "Coordenadas" o "Coordinates" con azul en la fila de encabezados.');

  return { headerRow, colPartida, colCoord, colREF, colEstado, colZona };
}

function readFinalTablePoints_(sheet, table) {
  const maxRows = sheet.getLastRow();
  const valid = [];
  const invalid = [];

  let r = table.headerRow + 1;
  while (r <= maxRows) {
    const partidaStr = String(sheet.getRange(r, table.colPartida).getValue() || "").trim();
    if (!partidaStr) break;

    const partidaNum = parseInt(partidaStr, 10);
    if (!isFinite(partidaNum)) break;

    const coordRaw = sheet.getRange(r, table.colCoord).getValue();
    const parsed = parseCoordToLatLng_(coordRaw);

    const refVal = table.colREF ? sheet.getRange(r, table.colREF).getValue() : "";
    const estadoVal = table.colEstado ? sheet.getRange(r, table.colEstado).getValue() : "";
    const zonaVal = table.colZona ? sheet.getRange(r, table.colZona).getValue() : "";

    if (!parsed) {
      invalid.push([partidaNum, String(refVal || ""), String(coordRaw || "")]);
    } else {
      valid.push({
        partida: partidaNum,
        lat: parsed.lat,
        lng: parsed.lng,
        ref: String(refVal || ""),
        estado: String(estadoVal || ""),
        zona: String(zonaVal || "")
      });
    }
    r++;
  }

  return { valid, invalid, lastRow: r - 1 };
}

// ---- match helper: exact cell text + blue bg
function findBlueHeaderCellAny_(sheet, texts, bgHex) {
  const bg = String(bgHex || "").toLowerCase();
  for (const t of texts) {
    const target = String(t).trim();
    if (!target) continue;
    const matches = sheet.createTextFinder(target).matchEntireCell(true).findAll() || [];
    for (const cell of matches) {
      if ((cell.getBackground() || "").toLowerCase() === bg) return cell;
    }
  }
  return null;
}

// ---- match helper: headers array + bg array, exact text + blue bg
function findHeaderColAny_(vals, bgs, headerTexts, bgHex) {
  const bg = String(bgHex || "").toLowerCase();
  const wanted = headerTexts.map(h => String(h || "").trim().toLowerCase()).filter(Boolean);

  for (let i = 0; i < vals.length; i++) {
    const v = String(vals[i] || "").trim().toLowerCase();
    const c = String(bgs[i] || "").toLowerCase();
    if (c === bg && wanted.includes(v)) return i + 1;
  }
  return null;
}

/*************** ✅ FIX: NO SOBRESCRIBIR ************************/
function clearPreviousMapOutput_(sheet, col, startRow) {
  const finder = sheet.createTextFinder("Mapa (Drive):").matchEntireCell(true);
  const all = finder.findAll() || [];
  for (const cell of all) {
    const r = cell.getRow();
    const c = cell.getColumn();
    if (c === col && r >= startRow) {
      // Limpia SOLO 3 filas x 6 columnas (lo que escribe el bloque de mapa)
      sheet.getRange(r, c, 3, 6).clearContent();
    }
  }
}

function findLastNonEmptyRowInBand_(sheet, startRow, startCol, widthCols) {
  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) return startRow;

  const numRows = lastRow - startRow + 1;
  const values = sheet.getRange(startRow, startCol, numRows, widthCols).getDisplayValues();

  for (let i = values.length - 1; i >= 0; i--) {
    const row = values[i];
    let any = false;
    for (let j = 0; j < row.length; j++) {
      if (String(row[j] || "").trim() !== "") { any = true; break; }
    }
    if (any) return startRow + i;
  }
  return startRow;
}

/*************** 4) PARSE COORDS *******************************/
function parseCoordToLatLng_(value) {
  if (value == null) return null;
  const s = String(value).trim();
  if (!s) return null;

  // decimal: lat,lng (o lng,lat)
  const mDec = s.match(/(-?\d{1,3}(?:\.\d+)?)\s*,\s*(-?\d{1,3}(?:\.\d+)?)/);
  if (mDec) {
    let a = parseFloat(mDec[1]);
    let b = parseFloat(mDec[2]);
    if (!isFinite(a) || !isFinite(b)) return null;
    if (Math.abs(a) > 90 && Math.abs(b) <= 90) { const t = a; a = b; b = t; }
    if (Math.abs(a) > 90 || Math.abs(b) > 180) return null;
    return { lat: a, lng: b };
  }

  // DMS: 32°27'17.7"N 116°59'04.0"W
  const re = /(\d{1,3})\s*[°º]\s*(\d{1,2})\s*['’′]\s*([\d.]+)\s*(?:"|”|″)?\s*([NSEW])/gi;
  const parts = [];
  let mm;
  while ((mm = re.exec(s)) !== null) {
    parts.push({ deg: +mm[1], min: +mm[2], sec: +mm[3], hem: String(mm[4]).toUpperCase() });
  }
  if (parts.length < 2) return null;

  const latPart = parts.find(p => p.hem === "N" || p.hem === "S");
  const lngPart = parts.find(p => p.hem === "E" || p.hem === "W");
  if (!latPart || !lngPart) return null;

  const lat = dmsToDec_(latPart.deg, latPart.min, latPart.sec, latPart.hem);
  const lng = dmsToDec_(lngPart.deg, lngPart.min, lngPart.sec, lngPart.hem);
  if (Math.abs(lat) > 90 || Math.abs(lng) > 180) return null;

  return { lat, lng };
}

function dmsToDec_(deg, min, sec, hem) {
  let dec = Math.abs(deg) + (min / 60) + (sec / 3600);
  if (hem === "S" || hem === "W") dec *= -1;
  return dec;
}

/*************** 5) BOUNDS + ZOOM ******************************/
function getBounds_(points) {
  let minLat = 90, maxLat = -90, minLng = 180, maxLng = -180;
  points.forEach(p => {
    minLat = Math.min(minLat, p.lat); maxLat = Math.max(maxLat, p.lat);
    minLng = Math.min(minLng, p.lng); maxLng = Math.max(maxLng, p.lng);
  });
  if (minLat === maxLat) { minLat -= 0.01; maxLat += 0.01; }
  if (minLng === maxLng) { minLng -= 0.01; maxLng += 0.01; }
  return { minLat, maxLat, minLng, maxLng };
}

function computeZoomForBounds_(b, mapW, mapH) {
  const ZOOM_MAX = 21;
  const WORLD_DIM = { height: 256, width: 256 };

  function latRad(lat) {
    const sin = Math.sin((lat * Math.PI) / 180);
    const rad = Math.log((1 + sin) / (1 - sin)) / 2;
    return Math.max(Math.min(rad, Math.PI), -Math.PI) / 2;
  }
  function zoom(mapPx, worldPx, fraction) {
    return Math.floor(Math.log(mapPx / worldPx / fraction) / Math.LN2);
  }

  const ne = { lat: b.maxLat, lng: b.maxLng };
  const sw = { lat: b.minLat, lng: b.minLng };

  const latFraction = (latRad(ne.lat) - latRad(sw.lat)) / Math.PI;
  let lngDiff = ne.lng - sw.lng; if (lngDiff < 0) lngDiff += 360;
  const lngFraction = lngDiff / 360;

  const latZoom = zoom(mapH, WORLD_DIM.height, latFraction);
  const lngZoom = zoom(mapW, WORLD_DIM.width, lngFraction);

  return Math.min(latZoom, lngZoom, ZOOM_MAX);
}

/*************** 6) DRIVE HELPERS (ROOT) ************************/
function getOrCreateMapsFolderInRoot_() {
  const root = DriveApp.getFolderById(OUTPUT_ROOT_FOLDER_ID);
  const it = root.getFoldersByName(MAPS_FOLDER_NAME);
  return it.hasNext() ? it.next() : root.createFolder(MAPS_FOLDER_NAME);
}

function trashFilesByName_(folder, fileName) {
  const it = folder.getFilesByName(fileName);
  while (it.hasNext()) {
    try { it.next().setTrashed(true); } catch (e) {}
  }
}
