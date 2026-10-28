/***************************************************************
 * MAPA CLIENTE (IMAGEN + PANEL LATERAL) — FIX NO SOBRESCRIBE
 *
 * ✅ Qué hace:
 *  - Lee la TABLA FINAL detectando encabezados AZULES (#9fc5e8)
 *  - Usa TODOS los renglones por "Partida" (ES) o "Item / 项目" (ENG/CHI) o "Item"
 *  - Usa la columna "Coordenadas" (ES) o "Coordinates" (ENG/CHI)
 *  - Genera imagen (Static Map) usando Maps Service
 *  - Dibuja un HEXÁGONO por cada propiedad, con su propio número de
 *    "Partida", encajado en una cuadrícula hexagonal fija (contiguos
 *    entre sí, nunca flotando ni superpuestos — ver
 *    renderMapaHexagonosIndividuales_ en "MapaHexagonal.js", AJUSTE 17).
 *    El mapa de círculos/abanico sigue disponible aparte en
 *    GenerarMapaClienteCirculos(), y el mapa de cuadrícula hexagonal CON
 *    CONTEO (AJUSTE 11) en GenerarMapaClienteHexConteo(), ambos más abajo
 *    en este archivo.
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

// Centro y zoom con los que encuadrar un conjunto de puntos (sin descargar
// el mapa todavía). zoomMax (opcional) = tope de acercamiento.
function calcularEncuadreHexParaPuntos_(puntos, mapPxW, mapPxH, zoomMax) {
  const bounds = getBounds_(puntos);
  const center = { lat: (bounds.minLat + bounds.maxLat) / 2, lng: (bounds.minLng + bounds.maxLng) / 2 };

  // Deja margen para que el hexágono de un punto en el borde del recuadro
  // no quede cortado por el encuadre del mapa.
  const margenPx = radioCirculoEnPxDelMapaBase_(mapPxH);
  const zoomAjustado = computeZoomForBounds_(bounds, mapPxW - margenPx * 2, mapPxH - margenPx * 2);
  const zoomInicial = (puntos.length === 1) ? 14 : Math.min(zoomAjustado, zoomMax || 21);

  // AJUSTE 18 (ver CAMBIO-2-CIRCULOS-MAPA-UBICACION.txt): computeZoomForBounds_
  // de arriba encuadra según la posición REAL de las propiedades — pero en
  // el mapa de hexágonos individuales un hexágono puede DESPLAZARSE de su
  // posición real si se solaparía con otro (ver separarHexagonosSolapados_
  // en "MapaHexagonal.js"), y ese desplazamiento puede empujar un hexágono más
  // allá del encuadre calculado, cortándolo en el borde. Se ajusta el zoom
  // (alejándolo si hace falta) contemplando dónde termina cada hexágono
  // DESPUÉS del desplazamiento, no solo la posición real de la propiedad.
  const zoom = (puntos.length === 1) ? zoomInicial :
    encontrarZoomHexagonosContiguos_(puntos, center, mapPxW, mapPxH, zoomInicial);

  return { center: center, zoom: zoom };
}

// Zoom ÚNICO para todas las secciones de un mapa seccionado. Cada sección
// necesita un zoom distinto según qué tan dispersas estén SUS propiedades
// (ej. en CTS LOGISTICS: 11, 12 y 14) — con cada una a su propio zoom el
// usuario vio que la escala del mapa y el tamaño del hexágono cambiaban de
// una captura a otra ("en todos el mismo"). Se usa el zoom MÁS ALEJADO que
// necesite alguna sección: así en todas caben sus hexágonos sin cortarse, y
// como el radio del hexágono depende del zoom (ver radioHexEnPuntosTemp_),
// también queda del mismo tamaño en todas. Se vuelve a comprobar el encaje
// de cada sección en ese zoom común por si el desplazamiento por colisión de
// alguna obligara a alejarlo un nivel más.
function zoomComunSecciones_(gruposDePuntos, mapPxW, mapPxH, zoomMax) {
  const encuadres = gruposDePuntos.map(p => calcularEncuadreHexParaPuntos_(p, mapPxW, mapPxH, zoomMax));
  let zoomComun = Math.min.apply(null, encuadres.map(e => e.zoom));
  for (let i = 0; i < 25; i++) {
    const zs = gruposDePuntos.map((p, k) => encontrarZoomHexagonosContiguos_(p, encuadres[k].center, mapPxW, mapPxH, zoomComun));
    const nuevo = Math.min.apply(null, zs);
    if (nuevo >= zoomComun) break;
    zoomComun = nuevo;
  }
  return zoomComun;
}

// Genera UN mapa hexagonal (Static Map + hexágonos individuales) para un
// conjunto de puntos ya acotado — extraído de GenerarMapaCliente() para
// poder llamarlo tanto con TODAS las propiedades (caso normal) como con
// las de UN SOLO cuadrante (caso seccionado, ver dividirEnCuadrantes_()).
// Cada sección se CENTRA en sus propias propiedades; el zoom es el que se
// pase en zoomFijo (mismo para todas las secciones) o, si no, el que
// necesiten estos puntos.
// zoomMax (opcional): tope de acercamiento al calcular el zoom de estos
// puntos: un cuadrante con pocas propiedades pegadas (ej. 2 a 100 m una de
// otra) acercaría el mapa a nivel de calle y quedaría sin contexto; 14 es el
// mismo zoom que ya se usa para una propiedad sola.
//
// zoomFijo (opcional): si se pasa, se usa ESE zoom tal cual en vez de
// calcular el de estos puntos — así todas las secciones de un mapa
// seccionado quedan con la MISMA escala (ver zoomComunSecciones_).
function prepararMapaHexParaPuntos_(puntos, mapPxW, mapPxH, zoomMax, zoomFijo) {
  const encuadre = calcularEncuadreHexParaPuntos_(puntos, mapPxW, mapPxH, zoomMax);
  const center = encuadre.center;
  const zoom = zoomFijo || encuadre.zoom;

  const map = Maps.newStaticMap()
    .setSize(mapPxW, mapPxH)
    .setCenter(center.lat, center.lng)
    .setZoom(zoom)
    .setMapType(Maps.StaticMap.Type.ROADMAP)
    .setFormat(Maps.StaticMap.Format.PNG);

  const baseMapBlob = map.getBlob().setName("MapaBase.png");

  // El dibujo de los hexágonos NO se hace acá: se separa para poder dibujar
  // varios mapas (los cuadrantes) en una sola tanda — ver
  // renderMapasHexagonosIndividualesLote_() en "MapaHexagonal.js".
  //
  // AJUSTE 17 (ver CAMBIO-2-CIRCULOS-MAPA-UBICACION.txt): el AJUSTE 16
  // había reutilizado el pipeline de abanico de los círculos (grupos
  // flotando separados, con halo y líneas, y un tope que OMITÍA grupos
  // muy grandes) — el usuario lo rechazó: no quería más omisiones ni
  // reportes de "no se dibujaron", y quería los hexágonos CONTIGUOS entre
  // sí (pegados) en vez de flotando. Desde el AJUSTE 26 el render deja cada
  // hexágono en su posición REAL y solo separa (lo mínimo, hasta quedar
  // pegados borde con borde) a los que se solaparían — así ninguna
  // propiedad se pierde, sin abanico ni halo ni "no se dibujaron".
  return { baseMapBlob: baseMapBlob, points: puntos, center: center, zoom: zoom };
}

// Versión de UN solo mapa (caso normal, sin seccionar): prepara y dibuja.
function generarMapaHexParaPuntos_(puntos, mapPxW, mapPxH) {
  const item = prepararMapaHexParaPuntos_(puntos, mapPxW, mapPxH);
  return renderMapaHexagonosIndividuales_(item.baseMapBlob, item.points, item.center, item.zoom, mapPxW, mapPxH);
}

// Sube VARIOS PNG a una carpeta de Drive a la vez (peticiones HTTP en
// paralelo con UrlFetchApp.fetchAll). Antes se hacía folder.createFile() de
// a uno — ~1.9 s por archivo, en fila; con 3-4 cuadrantes eran ~6-8 s solo
// guardando. Devuelve los IDs de los archivos creados, en el mismo orden.
// archivos: [{ name, blob }]
function subirPngsEnParalelo_(folder, archivos) {
  if (!archivos.length) return [];
  const token = ScriptApp.getOAuthToken();

  // Fase 1: sube el contenido de cada PNG (el blob se manda tal cual, sin
  // armar un cuerpo multipart a mano: convertirlo a arreglos de bytes de
  // JavaScript resultó MÁS LENTO que folder.createFile()).
  const resps1 = UrlFetchApp.fetchAll(archivos.map(a => ({
    url: "https://www.googleapis.com/upload/drive/v3/files?uploadType=media&fields=id",
    method: "post",
    contentType: "image/png",
    headers: { Authorization: "Bearer " + token },
    payload: a.blob,
    muteHttpExceptions: true
  })));
  const ids = resps1.map((resp, i) => {
    if (resp.getResponseCode() !== 200) {
      throw new Error('No se pudo subir "' + archivos[i].name + '" a Drive: ' + resp.getContentText().substring(0, 200));
    }
    return JSON.parse(resp.getContentText()).id;
  });

  // Fase 2: le pone su nombre y lo mueve a la carpeta de mapas.
  const resps2 = UrlFetchApp.fetchAll(ids.map((id, i) => ({
    url: "https://www.googleapis.com/drive/v3/files/" + id + "?addParents=" + folder.getId() + "&removeParents=root&fields=id",
    method: "patch",
    contentType: "application/json",
    headers: { Authorization: "Bearer " + token },
    payload: JSON.stringify({ name: archivos[i].name }),
    muteHttpExceptions: true
  })));
  resps2.forEach((resp, i) => {
    if (resp.getResponseCode() !== 200) {
      throw new Error('No se pudo mover "' + archivos[i].name + '" a la carpeta de mapas: ' + resp.getContentText().substring(0, 200));
    }
  });

  return ids;
}

// Divide los puntos en 4 cuadrantes geográficos simples (NE/NO/SE/SO) a
// partir del CENTRO del recuadro que los contiene a todos — no por
// cantidad pareja de propiedades, sino por área geográfica desde el
// centro (ver pregunta al usuario, "Cuadrantes geográficos simples").
// Puede dar cuadrantes con cantidades muy distintas entre sí (incluso
// alguno vacío) si las propiedades no están repartidas parejo alrededor
// del centro — eso es intencional y se maneja en GenerarMapaCliente()
// (un cuadrante vacío simplemente no genera una captura).
function dividirEnCuadrantes_(puntos, center) {
  const cuadrantes = { NE: [], NO: [], SE: [], SO: [] };
  puntos.forEach(p => {
    const norte = p.lat >= center.lat;
    const este = p.lng >= center.lng;
    if (norte && este) cuadrantes.NE.push(p);
    else if (norte && !este) cuadrantes.NO.push(p);
    else if (!norte && este) cuadrantes.SE.push(p);
    else cuadrantes.SO.push(p);
  });
  return cuadrantes;
}

// Respaldo de dividirEnCuadrantes_(): cuando los cuadrantes simples dejan
// alguno vacío (propiedades amontonadas de un solo lado del centro) no
// habría 4 secciones. Se reparte por MEDIANA: primero se cortan las
// propiedades en mitad oeste/este por longitud y luego cada mitad en
// norte/sur por latitud — siempre da 4 grupos no vacíos (hacen falta al
// menos 4 propiedades para que ninguna mitad quede vacía; en la práctica el
// seccionamiento nunca se activa con tan pocas). Mantiene las mismas
// etiquetas geográficas.
function dividirEnCuadrantesPorMediana_(puntos) {
  const porLng = puntos.slice().sort((a, b) => a.lng - b.lng);
  const mitad = Math.floor(porLng.length / 2);
  const oeste = porLng.slice(0, mitad);
  const este = porLng.slice(mitad);
  const cortarNorteSur = (grupo) => {
    const porLat = grupo.slice().sort((a, b) => a.lat - b.lat);
    const m = Math.floor(porLat.length / 2);
    return { sur: porLat.slice(0, m), norte: porLat.slice(m) };
  };
  const o = cortarNorteSur(oeste), e = cortarNorteSur(este);
  return { NE: e.norte, NO: o.norte, SE: e.sur, SO: o.sur };
}

// AJUSTE 31 (ver CAMBIO-2-CIRCULOS-MAPA-UBICACION.txt): decide en CUÁNTAS
// secciones partir cuando el mapa único no cabe. ANTES se armaban siempre 4
// cuadrantes en cuanto hacía falta seccionar — el usuario detectó con "CTS
// LOGISTICS" que la sección SE quedaba con solo 2 de las 56 propiedades
// (casi vacía) mientras que partiendo el mapa en solo 2 MITADES (norte/sur)
// ambas ya quedaban legibles (mancha contigua de 7 y 2, muy por debajo del
// umbral de 10) — la 3ra y 4ta captura no hacían falta.
//
// Ahora se prueban primero 2 formas de partir en 2 mitades (por latitud:
// sur/norte; por longitud: oeste/este) y, SOLO si ninguna de las 2 resuelve
// el apiñamiento en AMBAS mitades, se cae a los 4 cuadrantes de siempre
// (dividirEnCuadrantes_). Si más de una partición en 2 funciona, se elige la
// que deje la PEOR mancha más baja (la más legible de las 2 opciones).
function elegirSecciones_(puntos, mapPxW, mapPxH) {
  const manchaDeGrupo_ = (grupo) => {
    if (grupo.length < 2) return 0;
    const enc = calcularEncuadreHexParaPuntos_(grupo, mapPxW, mapPxH);
    const R = radioHexEnPuntosTemp_(enc.zoom, enc.center.lat, mapPxH);
    return tamanoMayorManchaContigua_(grupo, enc.center, enc.zoom, mapPxW, mapPxH, R);
  };

  const partirPorMediana_ = (eje) => {
    const ordenado = puntos.slice().sort((a, b) => eje === "lat" ? a.lat - b.lat : a.lng - b.lng);
    const mitad = Math.floor(ordenado.length / 2);
    return [ordenado.slice(0, mitad), ordenado.slice(mitad)];
  };

  const candidatosDe2 = [
    { etiquetas: ["Sur", "Norte"], grupos: partirPorMediana_("lat") },
    { etiquetas: ["Oeste", "Este"], grupos: partirPorMediana_("lng") }
  ];

  let mejor = null;
  candidatosDe2.forEach(candidato => {
    const manchas = candidato.grupos.map(manchaDeGrupo_);
    if (manchas.every(m => m <= UMBRAL_MANCHA_CONTIGUA)) {
      const peorMancha = Math.max.apply(null, manchas);
      if (!mejor || peorMancha < mejor.peorMancha) mejor = { candidato: candidato, peorMancha: peorMancha };
    }
  });

  if (mejor) {
    return mejor.candidato.etiquetas
      .map((etiqueta, i) => ({ etiqueta: etiqueta, puntos: mejor.candidato.grupos[i] }))
      .filter(g => g.puntos.length > 0);
  }

  // Ninguna partición en 2 alcanzó: 4 cuadrantes, como hasta ahora.
  const centro = calcularEncuadreHexParaPuntos_(puntos, mapPxW, mapPxH).center;
  let cuadrantes = dividirEnCuadrantes_(puntos, centro);
  if (ORDEN_CUADRANTES.some(q => cuadrantes[q].length === 0)) {
    cuadrantes = dividirEnCuadrantesPorMediana_(puntos);
  }
  return ORDEN_CUADRANTES
    .map(q => ({ etiqueta: q, puntos: cuadrantes[q] }))
    .filter(g => g.puntos.length > 0);
}

/*************** 1) GENERAR MAPA (IMAGEN) **********************/
function GenerarMapaCliente() {
  const HEADER_BG = "#9fc5e8"; // light blue 3

  const API_KEY = PropertiesService.getScriptProperties().getProperty("GMAPS_KEY");
  if (!API_KEY) throw new Error("No encontré GMAPS_KEY. Ejecuta ConfigurarGoogleMapsKey() primero.");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  // 1) Lee tabla final por encabezados azules
  const table = getFinalTableMeta_(sheet, HEADER_BG);
  const points = readFinalTablePoints_(sheet, table);
  if (!points.valid.length) throw new Error("No hay coordenadas válidas en la tabla final.");

  // AJUSTE 28 (ver CAMBIO-2-CIRCULOS-MAPA-UBICACION.txt): ANTES aquí se
  // separaba un "núcleo" de propiedades "sueltas" (excluirOutliersDeVista_,
  // AJUSTE 9) y el mapa único se dibujaba SOLO con el núcleo — las sueltas
  // se omitían en silencio (desde el AJUSTE 15 ni siquiera se reportaban en
  // la hoja). El usuario detectó que faltaba un hexágono en "INPLANMEX" (4
  // propiedades: la partida 3, a 3.8 km de su vecina más cercana, se
  // descartó por 0.08 km de diferencia frente a la partida 1) y se comprobó
  // que 15 de las 23 pestañas omitían al menos una propiedad. Regla nueva:
  // NINGUNA propiedad con coordenadas válidas se omite del mapa.
  const MAP_PX_W = 640, MAP_PX_H = 360; // 16:9, misma proporción que la página por defecto de una presentación (720x405)

  // Seccionamiento en 4 cuadrantes: solo cuando el mapa único de verdad
  // quedaría demasiado apiñado — ver tamanoMayorManchaContigua_() en
  // "MapaHexagonal.js" y AJUSTE 30 en CAMBIO-2-CIRCULOS-MAPA-UBICACION.txt.
  // Se mide sobre el mapa único a SU PROPIO zoom (el que se usaría si no se
  // seccionara): si los hexágonos que quedarían pegados unos con otros en un
  // mismo bloque (sin separación) son más de UMBRAL_MANCHA_CONTIGUA, ese
  // bloque se vuelve ilegible (números encimados) y se secciona.
  const encuadreUnico = calcularEncuadreHexParaPuntos_(points.valid, MAP_PX_W, MAP_PX_H);
  const R_unico = radioHexEnPuntosTemp_(encuadreUnico.zoom, encuadreUnico.center.lat, MAP_PX_H);
  const manchaMax = tamanoMayorManchaContigua_(points.valid, encuadreUnico.center, encuadreUnico.zoom, MAP_PX_W, MAP_PX_H, R_unico);
  const necesitaSeccionar = manchaMax > UMBRAL_MANCHA_CONTIGUA;

  const folder = getOrCreateMapsFolderInRoot_();
  const fileNameUnico = `${MAP_FILE_PREFIX}${sheet.getName()}.png`;
  // Todos los nombres de sección POSIBLES (2 mitades o 4 cuadrantes) de una
  // corrida anterior — se limpian los que no correspondan al modo elegido
  // esta vez, sea cual sea (ver TODAS_LAS_ETIQUETAS_SECCION en "00_Config.js").
  const fileNamesSeccionPosibles = TODAS_LAS_ETIQUETAS_SECCION.map(et => `${MAP_FILE_PREFIX}${sheet.getName()}_${et}.png`);

  const archivosGenerados = [];

  if (necesitaSeccionar) {
    // Limpia el archivo único de una corrida ANTERIOR sin seccionar, y
    // CUALQUIER nombre de sección de una corrida anterior (así se haya
    // usado el modo de 2 mitades o el de 4 cuadrantes) — para que
    // getMapFilesForSheetStrict_INT() no mezcle archivos viejos de un modo
    // con los nuevos del otro.
    trashFilesByName_(folder, fileNameUnico);
    fileNamesSeccionPosibles.forEach(fn => trashFilesByName_(folder, fn));

    // AJUSTE 31: elige la MENOR cantidad de secciones que alcance (2 mitades
    // si resuelven el apiñamiento; si no, los 4 cuadrantes de siempre) — ver
    // elegirSecciones_() más arriba.
    const grupos = elegirSecciones_(points.valid, MAP_PX_W, MAP_PX_H)
      .map(g => ({ etiqueta: g.etiqueta, fileName: `${MAP_FILE_PREFIX}${sheet.getName()}_${g.etiqueta}.png`, puntos: g.puntos }));

    // Velocidad: las secciones se dibujan TODAS juntas en una sola tanda
    // (una presentación temporal, un batchUpdate, miniaturas en paralelo) y
    // se suben a Drive en paralelo — hacerlo de a una multiplicaba por 4 el
    // trabajo fijo de cada mapa. Ver renderMapasHexagonosIndividualesLote_()
    // en "MapaHexagonal.js" y subirPngsEnParalelo_() arriba.
    // MISMO zoom (escala) en todas las secciones — ver zoomComunSecciones_.
    const zoomComun = zoomComunSecciones_(grupos.map(g => g.puntos), MAP_PX_W, MAP_PX_H, ZOOM_MAX_SECCION);
    const items = grupos.map(g => prepararMapaHexParaPuntos_(g.puntos, MAP_PX_W, MAP_PX_H, ZOOM_MAX_SECCION, zoomComun));
    const blobs = renderMapasHexagonosIndividualesLote_(items, MAP_PX_W, MAP_PX_H);
    const ids = subirPngsEnParalelo_(folder, blobs.map((blob, i) => ({ name: grupos[i].fileName, blob: blob })));
    grupos.forEach((g, i) => {
      archivosGenerados.push({ quadrante: g.etiqueta, fileId: ids[i], cantidad: g.puntos.length });
    });
  } else {
    // Limpia archivos de secciones de una corrida ANTERIOR seccionada de
    // esta misma pestaña (ej. si se agregaron/quitaron filas y ahora ya no
    // hace falta seccionar), para no dejar PNGs viejos sueltos en Drive.
    fileNamesSeccionPosibles.forEach(fn => trashFilesByName_(folder, fn));
    trashFilesByName_(folder, fileNameUnico);
    const blob = generarMapaHexParaPuntos_(points.valid, MAP_PX_W, MAP_PX_H);
    const ids = subirPngsEnParalelo_(folder, [{ name: fileNameUnico, blob: blob }]);
    archivosGenerados.push({ quadrante: null, fileId: ids[0], cantidad: points.valid.length });
  }

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

  // El texto de esta celda ("Mapa (Drive):") tiene que quedar SIEMPRE
  // igual, sea 1 mapa o seccionado — clearPreviousMapOutput_() lo busca
  // por texto EXACTO para poder limpiarlo en la próxima corrida; el detalle
  // de "seccionado en N" va en el texto del link, no en la etiqueta.
  sheet.getRange(outRow, outCol).setValue("Mapa (Drive):").setFontWeight("bold");
  if (archivosGenerados.length <= 1) {
    sheet.getRange(outRow, outCol + 1)
      .setFormula(`=HYPERLINK("https://drive.google.com/file/d/${archivosGenerados[0].fileId}/view","Abrir imagen del mapa")`);
  } else {
    sheet.getRange(outRow, outCol + 1)
      .setFormula(`=HYPERLINK("https://drive.google.com/drive/folders/${folder.getId()}","Abrir carpeta (mapa en ${archivosGenerados.length} secciones)")`);
  }

  sheet.getRange(outRow + 1, outCol).setValue('Para ajustar imagen, usar menú "Abrir mapa en panel lateral"');

  sheet.getRange(outRow + 2, outCol).setValue("Actualizado:").setFontWeight("bold");
  const tz = ss.getSpreadsheetTimeZone();
  sheet.getRange(outRow + 2, outCol + 1)
    .setValue(Utilities.formatDate(new Date(), tz, "dd/MM/yy"));

  // Reporte inválidos (si aplica)
  let cursorRow = outRow + 3 + 1;
  if (points.invalid.length) {
    const needWarnEnd = cursorRow + 2 + points.invalid.length;
    const maxRows2 = sheet.getMaxRows();
    if (needWarnEnd > maxRows2) sheet.insertRowsAfter(maxRows2, needWarnEnd - maxRows2);

    sheet.getRange(cursorRow, outCol).setValue("Coordenadas inválidas (no se mapearon):").setFontWeight("bold");
    sheet.getRange(cursorRow + 1, outCol, 1, 3).setValues([["Partida/Item", "REF", "Coordenadas/Coordinates"]]).setFontWeight("bold");
    sheet.getRange(cursorRow + 2, outCol, points.invalid.length, 3).setValues(points.invalid);
    cursorRow += 2 + points.invalid.length + 1;
  }

  // AJUSTE 15/17 (ver CAMBIO-2-CIRCULOS-MAPA-UBICACION.txt): a pedido del
  // usuario, la hoja ya NO reporta propiedades excluidas del mapa —
  // exceptuando "Coordenadas inválidas" arriba (esas de verdad no se
  // pueden ubicar). Las "muy alejadas del zoom" (AJUSTE 9) se siguen
  // excluyendo del encuadre sin dejar constancia (AJUSTE 15), y desde el
  // AJUSTE 17 el mapa de hexágonos individuales ya no omite NINGUNA
  // propiedad por proximidad (se desplaza a la celda libre más cercana en
  // vez de omitirse), así que tampoco hace falta ese reporte.
  if (archivosGenerados.length <= 1) {
    ss.toast(`Mapa guardado en ROOT/${MAPS_FOLDER_NAME}: ${fileNameUnico}`, "GenerarMapaCliente", 5);
  } else {
    const detalle = archivosGenerados.map(a => `${a.quadrante} (${a.cantidad})`).join(", ");
    ss.toast(`Mapa seccionado en ${archivosGenerados.length} partes (${detalle}) — guardado en ROOT/${MAPS_FOLDER_NAME}`, "GenerarMapaCliente", 6);
  }
}

/*************** 1b) GENERAR MAPA DE CÍRCULOS/ABANICO (ALTERNATIVA) ****/
// AJUSTE 11: hasta este ajuste, ESTA lógica (círculos + abanico + halo +
// tope de tamaño de abanico) era la que generaba GenerarMapaCliente() y
// el mapa definitivo insertado en las Slides. El usuario pidió reemplazar
// el mapa definitivo por la cuadrícula hexagonal (ver arriba); esta
// función preserva el comportamiento anterior tal cual, ahora como una
// alternativa de comparación (igual que ya se hizo con el mapa de calor y
// el de zonas en el AJUSTE 5) — guarda su PNG bajo un nombre distinto
// ("MapaCirculos_<hoja>.png") para no pisar el archivo que usa
// getMapFilesForSheetStrict_INT().
function GenerarMapaClienteCirculos() {
  const HEADER_BG = "#9fc5e8";
  const MAPTYPE = "roadmap";

  const API_KEY = PropertiesService.getScriptProperties().getProperty("GMAPS_KEY");
  if (!API_KEY) throw new Error("No encontré GMAPS_KEY. Ejecuta ConfigurarGoogleMapsKey() primero.");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  const table = getFinalTableMeta_(sheet, HEADER_BG);
  const points = readFinalTablePoints_(sheet, table);
  if (!points.valid.length) throw new Error("No hay coordenadas válidas en la tabla final.");

  const { core, outliers } = excluirOutliersDeVista_(points.valid);

  const bounds = getBounds_(core);
  const center = { lat: (bounds.minLat + bounds.maxLat) / 2, lng: (bounds.minLng + bounds.maxLng) / 2 };
  const MAP_PX_W = 640, MAP_PX_H = 360;

  const margenPx = radioCirculoEnPxDelMapaBase_(MAP_PX_H);
  const zoom = (core.length === 1) ? 14 :
    computeZoomForBounds_(bounds, MAP_PX_W - margenPx * 2, MAP_PX_H - margenPx * 2);

  const map = Maps.newStaticMap()
    .setSize(MAP_PX_W, MAP_PX_H)
    .setCenter(center.lat, center.lng)
    .setZoom(zoom)
    .setMapType(
      MAPTYPE === "satellite" ? Maps.StaticMap.Type.SATELLITE :
      MAPTYPE === "terrain"   ? Maps.StaticMap.Type.TERRAIN :
      MAPTYPE === "hybrid"    ? Maps.StaticMap.Type.HYBRID :
                                Maps.StaticMap.Type.ROADMAP
    )
    .setFormat(Maps.StaticMap.Format.PNG);

  const baseMapBlob = map.getBlob().setName("MapaBase.png");

  const renderResult = renderMapaConCirculos_(baseMapBlob, core, center, zoom, MAP_PX_W, MAP_PX_H);
  const blob = renderResult.blob.setName("MapaCirculos.png");
  const noMostrados = renderResult.noMostrados;

  const folder = getOrCreateMapsFolderInRoot_();
  const fileName = `MapaCirculos_${sheet.getName()}.png`;

  trashFilesByName_(folder, fileName);
  const file = folder.createFile(blob).setName(fileName);

  const avisos = [];
  if (outliers.length) avisos.push(`${outliers.length} muy alejada(s) fuera del zoom`);
  if (noMostrados.length) avisos.push(`${noMostrados.length} no dibujada(s) por proximidad excesiva`);
  const aviso = avisos.length ? ` (${avisos.join("; ")})` : "";
  ss.toast(`Mapa (círculos) guardado en ROOT/${MAPS_FOLDER_NAME}: ${fileName}${aviso}`, "GenerarMapaClienteCirculos", 5);
}

/*************** 1c) GENERAR MAPA DE CUADRÍCULA HEXAGONAL CON CONTEO
 * (ALTERNATIVA) *************************************************/
// AJUSTE 16: hasta este ajuste, ESTA lógica (cuadrícula hexagonal fija con
// el CONTEO de propiedades por celda, ver "MapaHexagonal.js" y AJUSTE 11)
// era la que generaba GenerarMapaCliente() y el mapa definitivo insertado
// en las Slides. El usuario probó varios ajustes de tamaño (AJUSTE 12/13/14)
// pero al final pidió volver a identificar cada propiedad por su propio
// número de "Partida" en vez de un conteo agregado — ver GenerarMapaCliente()
// más arriba, que ahora usa hexágonos individuales con el mismo pipeline de
// abanico que los círculos. Esta función preserva el comportamiento de
// cuadrícula-con-conteo tal cual, ahora como alternativa de comparación,
// guardando su PNG bajo un nombre distinto ("MapaHexConteo_<hoja>.png").
function GenerarMapaClienteHexConteo() {
  const HEADER_BG = "#9fc5e8";
  const MAPTYPE = "roadmap";

  const API_KEY = PropertiesService.getScriptProperties().getProperty("GMAPS_KEY");
  if (!API_KEY) throw new Error("No encontré GMAPS_KEY. Ejecuta ConfigurarGoogleMapsKey() primero.");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  const table = getFinalTableMeta_(sheet, HEADER_BG);
  const points = readFinalTablePoints_(sheet, table);
  if (!points.valid.length) throw new Error("No hay coordenadas válidas en la tabla final.");

  const { core, outliers } = excluirOutliersDeVista_(points.valid);

  const bounds = getBounds_(core);
  const center = { lat: (bounds.minLat + bounds.maxLat) / 2, lng: (bounds.minLng + bounds.maxLng) / 2 };
  const MAP_PX_W = 640, MAP_PX_H = 360;

  const margenPx = radioCirculoEnPxDelMapaBase_(MAP_PX_H);
  const zoom = (core.length === 1) ? 14 :
    computeZoomForBounds_(bounds, MAP_PX_W - margenPx * 2, MAP_PX_H - margenPx * 2);

  const map = Maps.newStaticMap()
    .setSize(MAP_PX_W, MAP_PX_H)
    .setCenter(center.lat, center.lng)
    .setZoom(zoom)
    .setMapType(
      MAPTYPE === "satellite" ? Maps.StaticMap.Type.SATELLITE :
      MAPTYPE === "terrain"   ? Maps.StaticMap.Type.TERRAIN :
      MAPTYPE === "hybrid"    ? Maps.StaticMap.Type.HYBRID :
                                Maps.StaticMap.Type.ROADMAP
    )
    .setFormat(Maps.StaticMap.Format.PNG);

  const baseMapBlob = map.getBlob().setName("MapaBase.png");

  const blob = renderMapaHexagonal_(baseMapBlob, core, center, zoom, MAP_PX_W, MAP_PX_H).setName("MapaHexConteo.png");

  const folder = getOrCreateMapsFolderInRoot_();
  const fileName = `MapaHexConteo_${sheet.getName()}.png`;

  trashFilesByName_(folder, fileName);
  folder.createFile(blob).setName(fileName);

  const aviso = outliers.length ? ` (${outliers.length} muy alejada(s) fuera del zoom)` : "";
  ss.toast(`Mapa (hexágonos con conteo) guardado en ROOT/${MAPS_FOLDER_NAME}: ${fileName}${aviso}`, "GenerarMapaClienteHexConteo", 5);
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

  // Marcador hexagonal (ver CAMBIO-2-CIRCULOS-MAPA-UBICACION.txt): este
  // panel lateral usaba el pin rojo redondo por defecto de Google Maps, que
  // quedó inconsistente una vez que el mapa definitivo (el que va en
  // Slides, ver MapaHexagonal.js/renderMapaHexagonosIndividuales_) pasó a
  // hexágonos color granate (CIRCULO_COLOR). La API de Maps JS no tiene una
  // forma "hexágono" predefinida, así que se arma con un ícono SVG manual
  // (google.maps.Symbol con un "path" de hexágono regular de punta plana,
  // mismo estilo que el mapa estático) en vez del Marker.icon por defecto.
  const hexPath = "M 11,0 L 5.5,9.53 L -5.5,9.53 L -11,0 L -5.5,-9.53 L 5.5,-9.53 Z";

  const html =
    '<!doctype html><html><head><meta charset="utf-8"/>' +
    '<meta name="viewport" content="width=device-width, initial-scale=1"/>' +
    '<style>html,body{height:100%;margin:0;font-family:Arial,sans-serif;}#map{height:100vh;width:100%;}</style>' +
    '</head><body><div id="map"></div>' +
    '<script>' +
    'var points=' + dataJson + ';' +
    'var HEX_PATH="' + hexPath + '";' +
    'var HEX_COLOR="' + CIRCULO_COLOR + '";' +
    'function initMap(){' +
    '  var map=new google.maps.Map(document.getElementById("map"),{' +
    '    center:{lat:points[0].lat,lng:points[0].lng},zoom:12,mapTypeId:"roadmap"' +
    '  });' +
    '  var info=new google.maps.InfoWindow();' +
    '  var bounds=new google.maps.LatLngBounds();' +
    '  points.forEach(function(p){' +
    '    var pos={lat:p.lat,lng:p.lng}; bounds.extend(pos);' +
    '    var m=new google.maps.Marker({position:pos,map:map,' +
    '      icon:{path:HEX_PATH,fillColor:HEX_COLOR,fillOpacity:1,strokeColor:"#ffffff",strokeWeight:1.5,scale:1.6,labelOrigin:new google.maps.Point(0,0)},' +
    '      label:{text:String(p.partida),color:"#ffffff",fontSize:"12px",fontWeight:"bold"}' +
    '    });' +
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
    findHeaderColAny_(headerVals, headerBgs, ["Coordenadas", "Coordinates", "Coordinates / 坐标"], headerBgHex);

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

  // Velocidad (2026-09-18): antes se hacían hasta 5 lecturas sueltas a la
  // hoja POR FILA (partida, coordenadas, REF, estado, zona) — cientos de
  // llamadas en una tabla grande. Ahora se lee de un solo golpe el bloque
  // de columnas que abarca todas las necesarias y se recorre en memoria.
  const colsUsadas = [table.colPartida, table.colCoord, table.colREF, table.colEstado, table.colZona].filter(c => c);
  const colMin = Math.min.apply(null, colsUsadas);
  const colMax = Math.max.apply(null, colsUsadas);
  const numFilasLeer = maxRows - table.headerRow;
  const bloque = numFilasLeer > 0
    ? sheet.getRange(table.headerRow + 1, colMin, numFilasLeer, colMax - colMin + 1).getValues()
    : [];
  const celda = (fila, col) => fila[col - colMin];

  let r = table.headerRow + 1;
  while (r <= maxRows) {
    const fila = bloque[r - table.headerRow - 1];

    const partidaStr = String(celda(fila, table.colPartida) || "").trim();
    if (!partidaStr) break;

    const partidaNum = parseInt(partidaStr, 10);
    if (!isFinite(partidaNum)) break;

    const coordRaw = celda(fila, table.colCoord);
    const parsed = parseCoordToLatLng_(coordRaw);

    const refVal = table.colREF ? celda(fila, table.colREF) : "";
    const estadoVal = table.colEstado ? celda(fila, table.colEstado) : "";
    const zonaVal = table.colZona ? celda(fila, table.colZona) : "";

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
// BUG encontrado y corregido el 2026-09-16 (misma clase de bug que el ya
// corregido en leerTablaConLinksPorNombre_ para la tabla, ver Ajuste 3 de
// CAMBIO-1-TABLA-NATIVA-SLIDES.txt): esta función devolvía la PRIMERA
// coincidencia de "Partida"/"Item" con fondo azul en TODA la hoja, no la
// ÚLTIMA — en una pestaña con más de una tabla con encabezado azul (p.ej.
// una tabla vieja más arriba y la tabla final vigente más abajo, como
// "CTS LOGISTICS": fila 88 y fila 132), el mapa terminaba leyendo la tabla
// VIEJA en vez de la vigente, con menos filas y una zona geográfica
// distinta a la real. Se corrigió recorriendo los matches de atrás para
// adelante, igual que findFinalTableMeta_ES() ya hacía para la tabla.
function findBlueHeaderCellAny_(sheet, texts, bgHex) {
  const bg = String(bgHex || "").toLowerCase();
  for (const t of texts) {
    const target = String(t).trim();
    if (!target) continue;
    const matches = sheet.createTextFinder(target).matchEntireCell(true).findAll() || [];
    for (let i = matches.length - 1; i >= 0; i--) {
      const cell = matches[i];
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

/*************** 7) CÍRCULOS DE ÁREA SOBRE EL MAPA **************/
// Historial del radio: 4cm (original) -> 4/6≈0.67cm (2026-09-15, primer
// ajuste) -> 0.5cm (2026-09-16, segundo ajuste, a pedido del usuario tras
// ver "CTS LOGISTICS" con 28 propiedades: seguía habiendo demasiado
// solapamiento). Sigue siendo un radio FIJO — con propiedades muy juntas
// puede seguir habiendo superposición (ver CLUSTER_DISTANCIA_CM más abajo
// para el manejo especial de esos casos).
const CIRCULO_RADIO_CM = 0.5;
// Cambiado de azul semitransparente a rojo sólido el 2026-09-16, a pedido
// del usuario tras revisar el resultado real ("color entero rojo").
const CIRCULO_COLOR = "#EA4335"; // rojo Google (Material Design "red 500-ish")
const CIRCULO_OPACIDAD_RELLENO = 1.0; // "entero" = sin transparencia
const PT_POR_CM = 28.3465;
// El mapa (cuadrado o 16:9) se inserta en el slide dentro de una caja de
// 1300x550pt (pageWidth-140 x pageHeight-260 sobre una plantilla de
// 1440x810) y, al ser más angosto que esa caja, siempre termina
// mostrándose con 550pt de alto (limitado por el alto de la caja, no por
// el ancho). Por eso el radio se calcula contra 550pt, sin importar el
// tamaño en píxeles real del PNG exportado.
const MAPA_FINAL_ALTO_PT = 550;
// Alto de página por defecto de una presentación nueva de Slides (no se
// puede fijar un tamaño de página personalizado vía Apps Script).
const TEMP_PRES_ALTO_PT = 405;

// Cuánto ocupa, en píxeles del mapa base (Static Map), el radio del
// círculo ya convertido a su tamaño final en el slide. Se usa para dejarle
// margen al encuadre del mapa y que ningún círculo quede cortado.
function radioCirculoEnPxDelMapaBase_(mapPxH) {
  const escalaTempAFinal = TEMP_PRES_ALTO_PT / MAPA_FINAL_ALTO_PT;
  const radioFinalPt = CIRCULO_RADIO_CM * PT_POR_CM;
  const radioTempPt = radioFinalPt * escalaTempAFinal;
  const escalaPxAPt = TEMP_PRES_ALTO_PT / mapPxH;
  return radioTempPt / escalaPxAPt;
}

// Proyección Web Mercator estándar (la misma que usa Google Maps) para
// ubicar un punto lat/lng dentro de los píxeles de un Static Map ya
// generado con un centro y zoom conocidos.
function proyectarPuntoEnMapa_(lat, lng, centerLat, centerLng, zoom, mapPxW, mapPxH) {
  const rad = Math.PI / 180;
  const proyectar = (la, ln) => {
    const seno = Math.min(Math.max(Math.sin(la * rad), -0.9999), 0.9999);
    return {
      x: 128 + ln * (256 / 360),
      y: 128 - (256 / (4 * Math.PI)) * Math.log((1 + seno) / (1 - seno))
    };
  };
  const escala = Math.pow(2, zoom);
  const c = proyectar(centerLat, centerLng);
  const p = proyectar(lat, lng);
  return {
    x: mapPxW / 2 + (p.x - c.x) * escala,
    y: mapPxH / 2 + (p.y - c.y) * escala
  };
}

// Distancia real entre 2 coordenadas (fórmula de Haversine), en metros.
function distanciaHaversineM_(lat1, lng1, lat2, lng2) {
  const R = 6371000;
  const rad = Math.PI / 180;
  const dLat = (lat2 - lat1) * rad;
  const dLng = (lng2 - lng1) * rad;
  const a = Math.sin(dLat / 2) ** 2 +
    Math.cos(lat1 * rad) * Math.cos(lat2 * rad) * Math.sin(dLng / 2) ** 2;
  return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
}

// AJUSTE 9 (ver CAMBIO-2-CIRCULOS-MAPA-UBICACION.txt): separa las
// propiedades "sueltas" (muy lejos de donde está el grueso de las demás)
// del "núcleo" (la mayoría, concentradas en una zona) — para poder
// encuadrar/hacer zoom sobre el núcleo SIN que una o dos propiedades muy
// alejadas obliguen a alejar tanto la vista que el núcleo denso quede
// ilegible.
//
// Primer intento (descartado): excluir solo por distancia a la mediana.
// Con "CTS LOGISTICS" (57 propiedades) esto solo detectó 1 outlier real
// (la de Toluca, a 80km) — el "núcleo" restante seguía abarcando ~15-20km
// (Tepotzotlán a Tlalnepantla), un área todavía demasiado grande para un
// zoom legible con 56 números.
//
// Versión actual: en vez de una sola pasada, se va sacando ITERATIVAMENTE
// el punto más lejano del centroide del núcleo actual, uno por uno, hasta
// que el núcleo quepa en un área de TARGET_DIAGONAL_KM (la distancia entre
// las 2 esquinas de su recuadro) o hasta topar con el límite de seguridad
// MAX_EXCLUIR_FRACCION (nunca se excluye más de 1 de cada 4 propiedades,
// para no vaciar el mapa en un dataset genuinamente disperso sin núcleo
// claro). Esto sí logra un acercamiento real sobre la zona más densa.
const TARGET_DIAGONAL_KM = 8;
const MAX_EXCLUIR_FRACCION = 0.25;

function excluirOutliersDeVista_(points) {
  if (points.length < 4) return { core: points, outliers: [] };

  function diagonalKm(pts) {
    const b = getBounds_(pts);
    return distanciaHaversineM_(b.minLat, b.minLng, b.maxLat, b.maxLng) / 1000;
  }

  const core = points.slice();
  const outliers = [];
  const maxExcluir = Math.floor(points.length * MAX_EXCLUIR_FRACCION);

  while (core.length > 3 && outliers.length < maxExcluir && diagonalKm(core) > TARGET_DIAGONAL_KM) {
    const centroid = {
      lat: core.reduce((s, p) => s + p.lat, 0) / core.length,
      lng: core.reduce((s, p) => s + p.lng, 0) / core.length
    };
    let peorIdx = 0, peorD = -1;
    core.forEach((p, i) => {
      const d = distanciaHaversineM_(centroid.lat, centroid.lng, p.lat, p.lng);
      if (d > peorD) { peorD = d; peorIdx = i; }
    });
    outliers.push(core[peorIdx]);
    core.splice(peorIdx, 1);
  }

  return { core: core, outliers: outliers };
}

// BUG DE FONDO encontrado y corregido el 2026-09-16, en 2 etapas (ver
// AJUSTE 8 en CAMBIO-2-CIRCULOS-MAPA-UBICACION.txt):
//
// Etapa 1 — el criterio de "están cerca" dependía SOLO de la distancia en
// píxeles ya proyectados en el mapa, que depende del zoom elegido. Con una
// tabla de propiedades muy dispersas geográficamente (más de las 5-30
// concentradas en una zona que es el caso normal — ver
// MAX_FILAS_TABLA_SLIDE_RECOMENDADO en TablaNativaSlides.js), el zoom
// necesario para que quepan todas en un solo mapa termina siendo tan
// alejado que el círculo de 0.5cm representa varios KILÓMETROS reales, y a
// esa escala CASI TODAS las propiedades quedan "cerca" en píxeles aunque
// estén a 10-20km reales una de otra — terminaban las 57 en un solo
// abanico gigante e ilegible. Primer intento de arreglo: agregar un tope
// de distancia REAL (Haversine) además del de píxeles, usando Union-Find
// para agrupar.
//
// Etapa 2 — ese primer intento no alcanzó: Union-Find ENCADENA
// transitivamente (si A está cerca de B, y B cerca de C, A/B/C quedan en
// el mismo grupo aunque A y C estén lejos entre sí). En un corredor denso
// real (muchas propiedades industriales alineadas a pocos cientos de
// metros una de la siguiente, como en "CTS LOGISTICS") esto seguía
// fusionando decenas de propiedades en un solo grupo gigante sin importar
// qué tan chico se pusiera el tope de distancia. Se reemplazó Union-Find
// por una GRILLA geográfica de tamaño fijo: cada celda mide como mucho
// MAX_CLUSTER_DISTANCIA_M de lado, así que ningún grupo puede terminar
// abarcando más que eso sin importar cuán denso sea el corredor — sin
// esta forma de encadenamiento, cada grupo queda acotado por diseño.
const MAX_CLUSTER_DISTANCIA_M = 300;
const METROS_POR_GRADO_LAT = 111320;

// AJUSTE 10 (ver CAMBIO-2-CIRCULOS-MAPA-UBICACION.txt): tamaño máximo de
// grupo que todavía se acomoda en abanico. El usuario pidió puntualmente
// que, cuando hay mucha proximidad, se acomoden a los costados usando el
// centro del grupo (esto ya lo hacía dibujarClusterEnAbanico_, ver más
// abajo) — PERO que si son más de 3 propiedades muy juntas y acomodarlas
// ya no se ve bien, mejor no mostrarlas en vez de forzar un abanico
// apretado e ilegible (el caso real encontrado en el AJUSTE 8/9: grupos
// de 8-12 propiedades en el corredor más denso de "CTS LOGISTICS").
const MAX_PUNTOS_POR_ABANICO = 3;

// Agrupa puntos por celda de una grilla geográfica de lado
// MAX_CLUSTER_DISTANCIA_M (ver historial arriba) — reemplaza al Union-Find
// por distancia que se usaba antes. radioTempPt/margenExtra ya no se usan
// para decidir el agrupado (se dejan en la firma para no tener que tocar
// las llamadas existentes); el ancho de la grilla en longitud se ajusta
// con el coseno de la latitud porque los grados de longitud miden menos
// metros mientras más lejos del ecuador.
function agruparPuntosCercanos_(puntos, radioTempPt, margenExtra) {
  const latCellDeg = MAX_CLUSTER_DISTANCIA_M / METROS_POR_GRADO_LAT;
  const grupos = new Map();
  puntos.forEach(p => {
    const lngCellDeg = MAX_CLUSTER_DISTANCIA_M / (METROS_POR_GRADO_LAT * Math.cos(p.lat * Math.PI / 180));
    const key = Math.round(p.lat / latCellDeg) + "_" + Math.round(p.lng / lngCellDeg);
    if (!grupos.has(key)) grupos.set(key, []);
    grupos.get(key).push(p);
  });
  return Array.from(grupos.values());
}

// AJUSTE 10 — segunda pasada, sobre los GRUPOS que ya arma
// agruparPuntosCercanos_ (no sobre los puntos crudos): se probó con datos
// reales de "CTS LOGISTICS" que la grilla real (arriba) por sí sola no
// alcanzaba — dejaba pasar como grupos "independientes" (de 1-2
// propiedades cada uno) a celdas VECINAS que, aunque no comparten celda,
// quedan tan cerca en PANTALLA a este zoom que sus círculos/abanicos se
// siguen viendo pegados entre sí. Se fusionan acá esos grupos vecinos por
// distancia en píxeles entre sus centroides (Union-Find, igual que el
// mecanismo viejo que se había descartado por encadenar sin límite — pero
// acá encadenar ya no es un problema: si la cadena crece demasiado, el
// grupo fusionado supera MAX_PUNTOS_POR_ABANICO y simplemente se omite del
// mapa en vez de mostrarse ilegible, que es exactamente lo que se pidió).
function fusionarGruposEnPantalla_(grupos, radioTempPt, margenExtra) {
  const n = grupos.length;
  const centroides = grupos.map(g => ({
    x: g.reduce((s, p) => s + p.x, 0) / g.length,
    y: g.reduce((s, p) => s + p.y, 0) / g.length
  }));
  const padre = Array.from({ length: n }, (_, i) => i);
  function encontrar(i) { while (padre[i] !== i) { padre[i] = padre[padre[i]]; i = padre[i]; } return i; }
  function unir(i, j) { const ri = encontrar(i), rj = encontrar(j); if (ri !== rj) padre[ri] = rj; }

  const umbral = radioTempPt * 2 * margenExtra;
  for (let i = 0; i < n; i++) {
    for (let j = i + 1; j < n; j++) {
      const dx = centroides[i].x - centroides[j].x, dy = centroides[i].y - centroides[j].y;
      if (Math.sqrt(dx * dx + dy * dy) < umbral) unir(i, j);
    }
  }

  const fusionados = new Map();
  for (let i = 0; i < n; i++) {
    const raiz = encontrar(i);
    if (!fusionados.has(raiz)) fusionados.set(raiz, []);
    fusionados.get(raiz).push(...grupos[i]);
  }
  return Array.from(fusionados.values());
}

// Círculo individual: caso normal, sin solapamiento con otras propiedades.
//
// BUG encontrado y corregido el 2026-09-16, en 2 intentos:
//  - 1er intento (insuficiente): números de 2 cifras (cualquier propiedad
//    de la 10 en adelante) se envolvían solos a 2 líneas dentro del
//    círculo, viéndose "cortados" (ej. "3" arriba y "8" abajo en vez de
//    "38"). Se probó achicar la letra según la cantidad de dígitos
//    (14pt→11pt para 2 cifras), pero el problema seguía igual — el margen
//    interno que Slides le agrega por defecto al texto DENTRO de una
//    figura (ellipse) deja tan poco espacio útil que ni una letra más
//    chica alcanzaba a entrar en una sola línea.
//  - Solución real: en vez de poner el texto DENTRO del círculo (donde no
//    se puede controlar ese margen interno con el servicio simple
//    SlidesApp), se agrega un CUADRO DE TEXTO aparte, más ancho que el
//    propio círculo, centrado exactamente encima — un cuadro de texto no
//    tiene ese margen tan agresivo y, al ser más ancho que el círculo,
//    el número nunca necesita envolverse sin importar cuántas cifras
//    tenga.
function insertarNumeroSobreForma_(slide, cx, cy, radioForma, texto, fontSize) {
  const anchoCaja = radioForma * 3.4;
  const altoCaja = radioForma * 1.3;
  const caja = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX,
    cx - anchoCaja / 2, cy - altoCaja / 2, anchoCaja, altoCaja);
  const tr = caja.getText();
  tr.setText(texto);
  tr.getTextStyle().setFontFamily("Arial").setFontSize(fontSize).setBold(true).setForegroundColor("#ffffff");
  tr.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
  // BUG encontrado y corregido el 2026-09-16: solo se centraba el texto
  // HORIZONTALMENTE (paragraphAlignment) — sin fijar la alineación
  // vertical del cuadro, Slides lo deja arriba por defecto, así que el
  // número quedaba corrido hacia arriba en vez de centrado en medio de la
  // figura (círculo o hexágono). setContentAlignment(MIDDLE) centra
  // también verticalmente.
  caja.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
  return caja;
}

// Inserta la forma del marcador (círculo o hexágono) — el hexágono usa un
// recuadro ancho x alto distinto al del círculo (2R x √3·R, "de punta
// plana") para que salga un hexágono regular en vez de uno estirado, ya
// que Slides arma la figura a partir del recuadro que se le pasa.
function insertarFormaMarcador_(slide, cx, cy, radioTempPt, tipoForma) {
  if (tipoForma === SlidesApp.ShapeType.HEXAGON) {
    const w = 2 * radioTempPt, h = Math.sqrt(3) * radioTempPt;
    return slide.insertShape(SlidesApp.ShapeType.HEXAGON, cx - radioTempPt, cy - h / 2, w, h);
  }
  return slide.insertShape(SlidesApp.ShapeType.ELLIPSE, cx - radioTempPt, cy - radioTempPt, radioTempPt * 2, radioTempPt * 2);
}

// tipoForma/colorForma son opcionales (círculo rojo por defecto, para no
// tener que tocar las llamadas ya existentes de GenerarMapaClienteCirculos())
// — se agregaron el 2026-09-16 para poder reusar toda esta lógica de
// agrupado/abanico/tope también con hexágonos (ver AJUSTE 16 en
// CAMBIO-2-CIRCULOS-MAPA-UBICACION.txt).
function dibujarCirculoIndividual_(slide, p, radioTempPt, tipoForma, colorForma) {
  tipoForma = tipoForma || SlidesApp.ShapeType.ELLIPSE;
  colorForma = colorForma || CIRCULO_COLOR;

  const forma = insertarFormaMarcador_(slide, p.x, p.y, radioTempPt, tipoForma);
  forma.getFill().setSolidFill(colorForma, CIRCULO_OPACIDAD_RELLENO);
  forma.getBorder().getLineFill().setSolidFill(colorForma);
  forma.getBorder().setWeight(1.5);

  insertarNumeroSobreForma_(slide, p.x, p.y, radioTempPt, String(p.partida), 14);
}

// Abanico ("spiderfy"): cuando 2 o más propiedades quedan tan cerca que sus
// círculos se solaparían, en vez de un círculo por cada una en su posición
// exacta (que quedarían tapándose entre sí, sobre todo con relleno sólido)
// o de una sola "zona combinada" genérica, se separan visualmente en un
// abanico alrededor del centro del grupo — el mismo patrón que usan Google
// Earth y la librería "Overlapping Marker Spiderfier" para este problema
// exacto. Cada propiedad conserva:
//   - un puntito pequeño en su posición geográfica REAL (sin mover), y
//   - una línea delgada ("pata") que conecta ese punto real con su círculo
//     numerado, ya desplazado a una posición donde no se solapa con las
//     demás del mismo grupo.
// Se activa SOLO para grupos con solapamiento real (ver agruparPuntosCercanos_
// más arriba) — una propiedad sin ninguna otra cerca sigue usando el círculo
// simple de dibujarCirculoIndividual_, sin línea ni puntito.
function dibujarClusterEnAbanico_(slide, puntosGrupo, radioTempPt, tipoForma, colorForma) {
  const n = puntosGrupo.length;
  const cx = puntosGrupo.reduce((s, p) => s + p.x, 0) / n;
  const cy = puntosGrupo.reduce((s, p) => s + p.y, 0) / n;

  // Radio del abanico: lo justo para que los círculos ya separados no se
  // vuelvan a tocar entre sí (geometría de cuerdas en un círculo), con un
  // margen del 25% para que se vea holgado y no al límite.
  const radioAbanico = Math.max(radioTempPt * 1.6, (radioTempPt / Math.sin(Math.PI / n)) * 1.25);

  // Halo amarillo suave detrás del abanico — agregado el 2026-09-16 a
  // pedido del usuario (mostró una imagen de referencia de otro sistema
  // con este mismo patrón: círculos numerados siempre separados y
  // legibles, más un halo amarillo alrededor de los que están cerca entre
  // sí en la realidad, para dar esa señal visual de "proximidad" sin
  // volver a juntar ni tapar los números). Se dibuja ANTES que las
  // líneas/puntos/círculos de más abajo para quedar DETRÁS de todo eso —
  // nunca tapa ni un número ni una pata.
  const HALO_COLOR = "#FFEB3B"; // amarillo, igual al de la imagen de referencia
  const HALO_OPACIDAD = 0.35;
  const radioHalo = radioAbanico + radioTempPt * 1.4;
  const halo = slide.insertShape(SlidesApp.ShapeType.ELLIPSE,
    cx - radioHalo, cy - radioHalo, radioHalo * 2, radioHalo * 2);
  halo.getFill().setSolidFill(HALO_COLOR, HALO_OPACIDAD);
  halo.getBorder().setTransparent();

  const LINEA_COLOR = "#666666";
  const PUNTO_REAL_PT = 6;

  // Orden estable por "Partida" para que el abanico salga siempre igual
  // entre corrida y corrida (más fácil de leer/comparar).
  const ordenados = puntosGrupo.slice().sort((a, b) => a.partida - b.partida);

  ordenados.forEach((p, i) => {
    const angulo = (i / n) * 2 * Math.PI - Math.PI / 2;
    const fx = cx + radioAbanico * Math.cos(angulo);
    const fy = cy + radioAbanico * Math.sin(angulo);

    // Pata: línea delgada de la posición real a la posición en el abanico.
    const linea = slide.insertLine(SlidesApp.LineCategory.STRAIGHT, p.x, p.y, fx, fy);
    linea.getLineFill().setSolidFill(LINEA_COLOR);
    linea.setWeight(0.75);

    // Puntito en la posición geográfica real (para no perder la ubicación exacta).
    const puntoReal = slide.insertShape(SlidesApp.ShapeType.ELLIPSE,
      p.x - PUNTO_REAL_PT / 2, p.y - PUNTO_REAL_PT / 2, PUNTO_REAL_PT, PUNTO_REAL_PT);
    puntoReal.getFill().setSolidFill(LINEA_COLOR);
    puntoReal.getBorder().getLineFill().setSolidFill(LINEA_COLOR);
    puntoReal.getBorder().setWeight(0.5);

    // Círculo (u hexágono) numerado, igual que uno individual, pero en la
    // posición del abanico.
    dibujarCirculoIndividual_(slide, { x: fx, y: fy, partida: p.partida }, radioTempPt, tipoForma, colorForma);
  });
}

// Dibuja los círculos (o zonas combinadas, ver arriba) con el número de
// "Partida" de cada propiedad sobre el mapa base, usando formas nativas de
// Slides en una presentación temporal (para lograr un radio físico exacto
// en centímetros), y devuelve el PNG final ya combinado.
//
// tipoForma/colorForma son opcionales (círculo rojo por defecto) — se
// agregaron el 2026-09-16 para poder reusar TODO este pipeline (agrupado,
// abanico, halo, tope de tamaño) también con hexágonos individuales en vez
// de círculos, a pedido del usuario (ver renderMapaConHexagonosIndividuales_
// y AJUSTE 16 en CAMBIO-2-CIRCULOS-MAPA-UBICACION.txt) sin duplicar toda
// esta función.
function renderMapaConCirculos_(baseMapBlob, points, center, zoom, mapPxW, mapPxH, tipoForma, colorForma) {
  tipoForma = tipoForma || SlidesApp.ShapeType.ELLIPSE;
  colorForma = colorForma || CIRCULO_COLOR;
  const tempPres = SlidesApp.create("__TEMP_MAPA__" + new Date().getTime());
  const tempId = tempPres.getId();

  try {
    const slide = tempPres.getSlides()[0];
    slide.getPageElements().forEach(pe => { try { pe.remove(); } catch (e) {} });

    const pageW = tempPres.getPageWidth();
    const pageH = tempPres.getPageHeight();
    const escalaPxAPt = pageH / mapPxH; // == pageW / mapPxW (misma proporción 16:9)

    const escalaTempAFinal = pageH / MAPA_FINAL_ALTO_PT;
    const radioFinalPt = CIRCULO_RADIO_CM * PT_POR_CM;
    const radioTempPt = radioFinalPt * escalaTempAFinal;

    const bg = slide.insertImage(baseMapBlob);
    bg.setLeft(0).setTop(0).setWidth(pageW).setHeight(pageH);

    const puntos = points.map(p => {
      const px = proyectarPuntoEnMapa_(p.lat, p.lng, center.lat, center.lng, zoom, mapPxW, mapPxH);
      return { x: px.x * escalaPxAPt, y: px.y * escalaPxAPt, partida: p.partida, lat: p.lat, lng: p.lng, ref: p.ref };
    });

    const gruposPorGrilla = agruparPuntosCercanos_(puntos, radioTempPt, 1.15);
    const grupos = fusionarGruposEnPantalla_(gruposPorGrilla, radioTempPt, 1.15);
    const noMostrados = [];
    grupos.forEach(grupo => {
      if (grupo.length === 1) {
        dibujarCirculoIndividual_(slide, grupo[0], radioTempPt, tipoForma, colorForma);
      } else if (grupo.length <= MAX_PUNTOS_POR_ABANICO) {
        dibujarClusterEnAbanico_(slide, grupo, radioTempPt, tipoForma, colorForma);
      } else {
        // AJUSTE 10 (ver CAMBIO-2-CIRCULOS-MAPA-UBICACION.txt): a pedido
        // del usuario, un abanico de más de MAX_PUNTOS_POR_ABANICO
        // propiedades tan próximas entre sí ya no se fuerza a acomodar
        // (quedaba ilegible en corredores muy densos) — se omite
        // directamente del mapa. No se pierden: quedan reportadas en la
        // hoja (ver GenerarMapaCliente) para que el usuario sepa cuáles
        // son y pueda revisarlas en el Panel Lateral de Mapas.
        grupo.forEach(p => noMostrados.push(p));
      }
    });

    tempPres.saveAndClose();

    const thumbnailUrl = "https://slides.googleapis.com/v1/presentations/" +
      tempId + "/pages/" + SlidesApp.openById(tempId).getSlides()[0].getObjectId() +
      "/thumbnail?thumbnailProperties.mimeType=PNG&thumbnailProperties.thumbnailSize=LARGE";

    const token = ScriptApp.getOAuthToken();
    const resp = UrlFetchApp.fetch(thumbnailUrl, {
      headers: { Authorization: "Bearer " + token },
      muteHttpExceptions: true
    });
    if (resp.getResponseCode() !== 200) {
      throw new Error("Slides thumbnail API falló al generar el mapa con círculos: " + resp.getContentText().substring(0, 300));
    }

    const json = JSON.parse(resp.getContentText());
    const imgResp = UrlFetchApp.fetch(json.contentUrl);
    return { blob: imgResp.getBlob().setContentType("image/png"), noMostrados: noMostrados };

  } finally {
    DriveApp.getFileById(tempId).setTrashed(true);
  }
}
