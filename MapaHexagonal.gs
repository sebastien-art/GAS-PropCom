/***************************************************************
 * MAPA HEXAGONAL — ver CAMBIO-2-CIRCULOS-MAPA-UBICACION.txt
 *
 * Este archivo tiene DOS renders hexagonales distintos, de 2 ajustes
 * distintos:
 *
 * 1) renderMapaHexagonal_() (AJUSTE 11) — cuadrícula hexagonal fija donde
 *    cada hexágono muestra el CONTEO de propiedades que caen ahí (no su
 *    "Partida" individual), coloreado por densidad (amarillo pálido =
 *    pocas, rojo oscuro = muchas). Fue el mapa definitivo brevemente,
 *    ahora es una alternativa de comparación —
 *    GenerarMapaClienteHexConteo(), en "GenerarMapaCliente.js".
 *
 * 2) renderMapaHexagonosIndividuales_() (AJUSTE 17) — ES EL MAPA
 *    DEFINITIVO ACTUAL (el que usa GenerarMapaCliente()). Un hexágono POR
 *    PROPIEDAD, con su propio número de "Partida" (nunca un conteo), en su
 *    posición REAL — solo se separan (lo mínimo, hasta quedar pegados borde
 *    con borde) los hexágonos que de verdad se solaparían (AJUSTE 26; antes
 *    se encajaban en una cuadrícula fija) — así ninguno se superpone y
 *    ninguna propiedad se omite nunca. Desde
 *    el AJUSTE 19, el radio del hexágono representa una distancia REAL
 *    fija en metros (HEX_RADIO_M) en vez de un tamaño fijo en pantalla —
 *    se ve más grande con el mapa acercado y más chico alejado, como pidió
 *    el usuario (igual que el radio físico de los círculos).
 ***************************************************************/

// Redondeo de coordenadas axiales de hexágono a la celda entera más
// cercana (algoritmo estándar de "cube rounding" para grillas hexagonales
// — ver la guía de Red Blob Games sobre hex grids). Necesario porque la
// conversión de píxeles a coordenadas de hexágono da valores fraccionarios
// que hay que ajustar a un hexágono real sin que las 3 coordenadas del
// "cubo" equivalente pierdan la propiedad x+y+z=0.
function hexRedondear_(qFloat, rFloat) {
  let x = qFloat, z = rFloat, y = -x - z;
  let rx = Math.round(x), ry = Math.round(y), rz = Math.round(z);
  const dx = Math.abs(rx - x), dy = Math.abs(ry - y), dz = Math.abs(rz - z);
  if (dx > dy && dx > dz) rx = -ry - rz;
  else if (dy > dz) ry = -rx - rz;
  else rz = -rx - ry;
  return { q: rx, r: rz };
}

// Interpola un color entre amarillo pálido (pocas propiedades) y rojo
// oscuro (muchas), según qué tan lleno está este hexágono en relación al
// más lleno de todo el mapa — mismo espíritu que colorPorConteo_ en
// "MapaZonasYCalor.js", pero en degradado continuo en vez de 3 escalones,
// para que se vea igual de gradual que la imagen de referencia del
// usuario (15 en rojo oscuro, 8 en rojo/naranja, 3 en amarillo, etc.).
function colorGradienteHex_(count, maxCount) {
  const ratio = maxCount > 1 ? (count - 1) / (maxCount - 1) : 1;
  const amarillo = [255, 235, 130];
  const rojoOscuro = [140, 10, 10];
  const rgb = amarillo.map((v, i) => Math.round(v + (rojoOscuro[i] - v) * ratio));
  return "#" + rgb.map(v => v.toString(16).padStart(2, "0")).join("");
}

// Dibuja la cuadrícula hexagonal con el conteo de propiedades por celda
// sobre el mapa base, usando formas nativas de Slides (mismo mecanismo de
// "presentación temporal + exportar miniatura a PNG" que renderMapaConCirculos_
// en "GenerarMapaCliente.js"), y devuelve el PNG final ya combinado.
//
// El tamaño del hexágono NO es un tamaño físico fijo (a diferencia del
// círculo, que sí usa un radio real en cm) — se calcula para que quepan
// NUM_HEX_ANCHO hexágonos a lo ancho del mapa, sin importar el zoom. Esto
// es intencional: la cuadrícula existe para AGRUPAR visualmente, no para
// representar un tamaño real en el terreno, igual que cualquier mapa de
// hexágonos de referencia (Uber H3, Kepler.gl, etc.).
//
// Historial: 7 (original) -> 13 (2026-09-16, 1er ajuste: "deben ser
// hexágonos muy chicos", cada hexágono cubría una porción demasiado
// grande del mapa real) -> 52 (2026-09-16, 2do ajuste, mismo día: el
// usuario vio la versión de 13 todavía "muy grande" y pidió reducir el
// TAMAÑO del hexágono a su cuarta parte) -> 36 (2026-09-16, 3er ajuste,
// mismo día: a 52 quedó demasiado chico — "aumentar un poco el tamaño de
// los hexágonos". El tamaño de letra del conteo (ver tamanoLetraHex_ más
// abajo) se ajusta solo en función del radio real, así que sigue viéndose
// bien proporcionado sin tocar nada más.
const NUM_HEX_ANCHO = 36;

// Tamaño de letra del conteo dentro de cada hexágono, en función del radio
// real del hexágono (R, en puntos) — a hexágonos más chicos, letra más
// chica, para que el número no se vea más grande que su propia celda.
// Los límites (11-16pt, subidos de 9-16 a pedido del usuario el
// 2026-09-16 junto con el tamaño del hexágono) evitan que quede ilegible
// en cuadrículas muy finas o desproporcionadamente grande en cuadrículas
// gruesas.
function tamanoLetraHex_(R) {
  return Math.round(Math.max(11, Math.min(16, R * 0.4)));
}

function renderMapaHexagonal_(baseMapBlob, points, center, zoom, mapPxW, mapPxH) {
  const tempPres = SlidesApp.create("__TEMP_MAPA_HEX__" + new Date().getTime());
  const tempId = tempPres.getId();

  try {
    const slide = tempPres.getSlides()[0];
    slide.getPageElements().forEach(pe => { try { pe.remove(); } catch (e) {} });

    const pageW = tempPres.getPageWidth();
    const pageH = tempPres.getPageHeight();
    const escalaPxAPt = pageH / mapPxH;

    const bg = slide.insertImage(baseMapBlob);
    bg.setLeft(0).setTop(0).setWidth(pageW).setHeight(pageH);

    // R = radio del hexágono (centro a vértice) en puntos de la
    // presentación temporal. Con hexágonos "de punta plana" (flat-top,
    // como el autoshape HEXAGON de Slides), el espaciado centro-a-centro
    // horizontal entre columnas vecinas es 1.5*R.
    const R = pageW / (NUM_HEX_ANCHO * 1.5);

    const puntos = points.map(p => {
      const px = proyectarPuntoEnMapa_(p.lat, p.lng, center.lat, center.lng, zoom, mapPxW, mapPxH);
      return { x: px.x * escalaPxAPt, y: px.y * escalaPxAPt };
    });

    // Conversión de píxeles a coordenadas axiales de hexágono (flat-top) —
    // fórmulas estándar de Red Blob Games.
    const celdas = new Map();
    puntos.forEach(p => {
      const qFloat = (2 / 3 * p.x) / R;
      const rFloat = (-1 / 3 * p.x + Math.sqrt(3) / 3 * p.y) / R;
      const { q, r } = hexRedondear_(qFloat, rFloat);
      const key = q + "_" + r;
      if (!celdas.has(key)) celdas.set(key, { q: q, r: r, count: 0 });
      celdas.get(key).count++;
    });

    const maxCount = Math.max.apply(null, Array.from(celdas.values()).map(c => c.count));

    celdas.forEach(c => {
      // Axial a píxeles (flat-top) — inversa de la conversión de arriba.
      const cx = R * (3 / 2 * c.q);
      const cy = R * (Math.sqrt(3) / 2 * c.q + Math.sqrt(3) * c.r);
      const w = 2 * R, h = Math.sqrt(3) * R;

      // Opacidad subida de 0.88 a 0.97 a pedido del usuario ("que se
      // mantenga de un solo color") — con más transparencia, el mapa de
      // fondo se alcanzaba a ver levemente a través del hexágono y podía
      // notarse como un color no del todo parejo; casi opaco se ve como
      // un color sólido y uniforme de punta a punta.
      const hex = slide.insertShape(SlidesApp.ShapeType.HEXAGON, cx - R, cy - h / 2, w, h);
      hex.getFill().setSolidFill(colorGradienteHex_(c.count, maxCount), 0.97);
      hex.getBorder().getLineFill().setSolidFill("#ffffff");
      hex.getBorder().setWeight(1.25);

      // El número va en un cuadro de texto aparte, no dentro del propio
      // hexágono — ver insertarNumeroSobreForma_ en "GenerarMapaCliente.js"
      // (mismo bug/fix que en los círculos: el margen interno que Slides
      // agrega al texto DENTRO de una figura no deja lugar para 2 cifras
      // en una sola línea, sin importar qué tan chica se ponga la letra).
      insertarNumeroSobreForma_(slide, cx, cy, R, String(c.count), tamanoLetraHex_(R));
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
      throw new Error("Slides thumbnail API falló al generar el mapa hexagonal: " + resp.getContentText().substring(0, 300));
    }

    const json = JSON.parse(resp.getContentText());
    const imgResp = UrlFetchApp.fetch(json.contentUrl);
    return imgResp.getBlob().setContentType("image/png");

  } finally {
    DriveApp.getFileById(tempId).setTrashed(true);
  }
}

/***************************************************************
 * AJUSTE 17 — HEXÁGONOS INDIVIDUALES PERO CONTIGUOS (sin abanico)
 *
 * El AJUSTE 16 reutilizó el pipeline del mapa de círculos (agrupar +
 * abanico + halo + tope que omite grupos muy grandes) para dibujar
 * hexágonos con el número de "Partida" de cada propiedad. El usuario
 * rechazó ese resultado por 2 motivos: (a) no quería que se volviera a
 * omitir propiedades ni a reportar un "no se dibujaron" en la hoja, y
 * (b) no quería hexágonos "flotando" separados con halo y líneas
 * (el look del abanico) sino CONTIGUOS entre sí, pegados unos con otros
 * como una cuadrícula real — igual que el mapa de conteo del AJUSTE 11,
 * pero con un número de Partida único por celda en vez de un conteo.
 *
 * Se resuelve encajando cada propiedad en la MISMA grilla hexagonal fija
 * que ya usa renderMapaHexagonal_ (misma R, mismas fórmulas de
 * píxel<->axial): si la celda "ideal" de una propiedad (la más cercana a
 * su posición real proyectada en el mapa) ya está ocupada por otra
 * propiedad, se la desplaza a la celda LIBRE más cercana, buscando en
 * "anillos" concéntricos de celdas vecinas cada vez más lejos (algoritmo
 * estándar de espiral hexagonal — ver anilloHex_ más abajo). Como la
 * grilla es infinita en la práctica (no hay tope de anillos razonable),
 * NINGUNA propiedad queda sin mostrarse — no hace falta ningún reporte de
 * "no se dibujaron". Y como cada hexágono siempre queda anclado a una
 * celda real de la grilla, dos hexágonos vecinos SIEMPRE quedan pegados
 * entre sí (contiguos), nunca superpuestos ni flotando con espacio de
 * por medio.
 ***************************************************************/

// AJUSTE 26 (ver CAMBIO-2-CIRCULOS-MAPA-UBICACION.txt): LA REJILLA FIJA de
// los ajustes 17/18 YA NO SE DIBUJA (se conserva solo para elegir el zoom, ver
// "CRITERIO DE ZOOM" más abajo — AJUSTE 27). Diagnóstico (AJUSTE 25, sección NO de
// "CTS LOGISTICS"): la rejilla movía TODOS los hexágonos a la celda más
// cercana — incluso los que no se solapaban con nadie (hasta ~0.8 km) — y
// al chocar empujaba al siguiente a la celda libre más cercana en un orden
// fijo, encadenándose hasta ~2.2 km de la posición real. Eso "juntaba" los
// hexágonos en el centro aunque no hiciera falta. El usuario pidió: solo
// juntarlos cuando se van a solapar.
//
// Ahora cada hexágono se dibuja EXACTAMENTE en su posición real proyectada
// en el mapa, y solo se separa lo indispensable entre los que SÍ se
// solapan. Dos hexágonos del mismo tamaño y orientación (punta plana) NO se
// solapan si y solo si la diferencia entre sus centros queda FUERA del
// hexágono del doble de tamaño (suma de Minkowski): con R = radio,
//   se solapan  <=>  |dy| < √3·R  y  √3·|dx| + |dy| < 2·√3·R
// Para cada par que se solapa se los aleja sobre la línea que une sus
// centros, mitad y mitad, hasta quedar justo tangentes (pegados borde con
// borde). Se repite hasta que ningún par se solape; un hexágono sin
// vecinos que lo pisen nunca se mueve.
//
// limites (opcional): {minX, maxX, minY, maxY} — el hexágono nunca sale de
// ese recuadro (se corrige la posición tras cada empuje), para que separar
// un cúmulo pegado al borde del mapa no lo corte.
function separarHexagonosSolapados_(puntos, R, limites) {
  const S3 = Math.sqrt(3);
  const ANGULO_DORADO = 2.399963; // reparte en direcciones distintas los puntos con coordenadas idénticas
  const dentro = (p) => {
    if (!limites) return;
    p.x = Math.min(limites.maxX, Math.max(limites.minX, p.x));
    p.y = Math.min(limites.maxY, Math.max(limites.minY, p.y));
  };
  puntos.forEach(dentro);
  for (let iter = 0; iter < 500; iter++) {
    let huboSolape = false;
    for (let i = 0; i < puntos.length; i++) {
      for (let j = i + 1; j < puntos.length; j++) {
        let dx = puntos[j].x - puntos[i].x, dy = puntos[j].y - puntos[i].y;
        if (Math.abs(dx) < 1e-6 && Math.abs(dy) < 1e-6) {
          dx = 0.01 * Math.cos(j * ANGULO_DORADO);
          dy = 0.01 * Math.sin(j * ANGULO_DORADO);
        }
        const ax = Math.abs(dx), ay = Math.abs(dy);
        if (ay < S3 * R && S3 * ax + ay < 2 * S3 * R) {
          huboSolape = true;
          // factor por el que hay que escalar el vector entre centros para
          // llegar al borde del hexágono doble (justo tangentes)
          const t = Math.min(S3 * R / Math.max(ay, 1e-9), 2 * S3 * R / (S3 * ax + ay));
          const k = (t - 1) / 2 + 0.001;
          puntos[i].x -= dx * k; puntos[i].y -= dy * k;
          puntos[j].x += dx * k; puntos[j].y += dy * k;
          dentro(puntos[i]); dentro(puntos[j]);
        }
      }
    }
    if (!huboSolape) break;
  }
}

// Calcula, para un zoom dado, dónde se dibuja el hexágono de cada
// propiedad (su centro en puntos de la presentación temporal: cx, cy): su
// posición real proyectada, separada solo de los hexágonos con los que se
// solaparía (ver separarHexagonosSolapados_), sin salirse del recuadro del
// mapa. SOLO dibuja: el ZOOM se elige con el criterio de siempre, ver
// zoomRejillaParaEncuadre_ más abajo.
function calcularAsignacionesHex_(points, center, zoom, mapPxW, mapPxH, escalaPxAPt, R) {
  const puntos = points.map(p => {
    const px = proyectarPuntoEnMapa_(p.lat, p.lng, center.lat, center.lng, zoom, mapPxW, mapPxH);
    return { x: px.x * escalaPxAPt, y: px.y * escalaPxAPt, partida: p.partida };
  });
  const S3 = Math.sqrt(3);
  const anchoPt = mapPxW * escalaPxAPt, altoPt = mapPxH * escalaPxAPt;
  separarHexagonosSolapados_(puntos, R, {
    minX: R, maxX: anchoPt - R, minY: S3 / 2 * R, maxY: altoPt - S3 / 2 * R
  });
  return puntos.map(p => ({ partida: p.partida, cx: p.x, cy: p.y }));
}

// AJUSTE 30 (ver CAMBIO-2-CIRCULOS-MAPA-UBICACION.txt): mide qué tan
// apiñado queda un mapa de UN SOLO hexágono por propiedad — el TAMAÑO DE LA
// MANCHA CONTIGUA MÁS GRANDE (cuántos hexágonos quedan pegados unos con
// otros en un mismo bloque, sin separación). Se arma el "grafo de contacto"
// (dos hexágonos son vecinos si al dibujarlos en su posición real se
// tocarían/solaparían) y se busca la componente conexa más grande —
// búsqueda estándar de componentes conexas.
//
// Reemplaza al criterio del AJUSTE 29 (mediana del desplazamiento, en
// radios de hexágono, de TODOS los hexágonos): ese número se diluía cuando
// la mayoría del dataset estaba bien repartida y solo una zona puntual
// estaba densa — "CTS LOGISTICS" daba una mediana baja (1.8 radios) pese a
// que el usuario mostró que, en la imagen real, esa zona densa era
// ilegible. El GRADO máximo de solape (cuántos hexágonos se solapan con el
// más rodeado) tampoco alcanzaba: "PRUEBA" y "EJEMPLO EDOMEX" daban el
// MISMO grado máximo (11) que "CTS LOGISTICS", pero eso no dice si esos 11
// forman una mancha sólida o varias cadenas sueltas. El tamaño de la mancha
// SÍ lo distingue, y de paso reveló que "PRUEBA" (mancha de 12, el 80% de
// sus 15 propiedades) y "EJEMPLO EDOMEX" (12, el 60% de 20) tenían el mismo
// problema que "CTS LOGISTICS" (25, el 45% de 56) aunque nadie lo hubiera
// reportado todavía.
function tamanoMayorManchaContigua_(points, center, zoom, mapPxW, mapPxH, R) {
  const escalaPxAPt = TEMP_PRES_ALTO_PT / mapPxH;
  const S3 = Math.sqrt(3);
  const posiciones = points.map(p => {
    const px = proyectarPuntoEnMapa_(p.lat, p.lng, center.lat, center.lng, zoom, mapPxW, mapPxH);
    return { x: px.x * escalaPxAPt, y: px.y * escalaPxAPt };
  });

  const n = posiciones.length;
  const visitado = new Array(n).fill(false);
  const vecinosDe = (i) => {
    const vecinos = [];
    for (let j = 0; j < n; j++) {
      if (j === i) continue;
      const ax = Math.abs(posiciones[j].x - posiciones[i].x);
      const ay = Math.abs(posiciones[j].y - posiciones[i].y);
      if (ay < S3 * R && S3 * ax + ay < 2 * S3 * R) vecinos.push(j);
    }
    return vecinos;
  };

  let mancha = 0;
  for (let i = 0; i < n; i++) {
    if (visitado[i]) continue;
    let tamano = 0;
    const pendientes = [i];
    visitado[i] = true;
    while (pendientes.length) {
      const actual = pendientes.pop();
      tamano++;
      vecinosDe(actual).forEach(v => { if (!visitado[v]) { visitado[v] = true; pendientes.push(v); } });
    }
    mancha = Math.max(mancha, tamano);
  }
  return mancha;
}

// ---------------------------------------------------------------------------
// CRITERIO DE ZOOM (el de siempre — NO cambiar sin avisar al usuario)
// ---------------------------------------------------------------------------
// El AJUSTE 26 cambió cómo se COLOCAN los hexágonos, pero por descuido
// también cambió cómo se ELIGE EL ZOOM (encontrarZoomHexagonosContiguos_
// probaba el encaje con la colocación nueva): en algunas pestañas la escala
// se alejaba un nivel respecto de lo que el usuario ya conocía ("MUEBLES
// DICO": zoom 11 -> 10). El usuario pidió que la escala de una sola imagen
// se mantenga como era antes. Por eso el ZOOM se sigue decidiendo con el
// mismo criterio de los ajustes 18-25: en qué celda de una rejilla
// hexagonal fija caería cada propiedad (con desplazamiento a la celda libre
// más cercana si está ocupada) y si todas caben dentro del recuadro. Esa
// rejilla YA NO se dibuja — solo se usa para elegir el zoom.
const DIRECCIONES_HEX_ = [[1, 0], [1, -1], [0, -1], [-1, 0], [-1, 1], [0, 1]];

function anilloHex_(centroQ, centroR, radio) {
  const resultado = [];
  let q = centroQ + DIRECCIONES_HEX_[4][0] * radio;
  let r = centroR + DIRECCIONES_HEX_[4][1] * radio;
  for (let i = 0; i < 6; i++) {
    for (let j = 0; j < radio; j++) {
      resultado.push({ q: q, r: r });
      q += DIRECCIONES_HEX_[i][0];
      r += DIRECCIONES_HEX_[i][1];
    }
  }
  return resultado;
}

function encontrarCeldaLibre_(qIdeal, rIdeal, ocupadas) {
  if (!ocupadas.has(qIdeal + "_" + rIdeal)) return { q: qIdeal, r: rIdeal };
  for (let radio = 1; radio <= 60; radio++) {
    const anillo = anilloHex_(qIdeal, rIdeal, radio);
    for (let i = 0; i < anillo.length; i++) {
      const c = anillo[i];
      if (!ocupadas.has(c.q + "_" + c.r)) return c;
    }
  }
  return { q: qIdeal, r: rIdeal };
}

// Centro (cx, cy) de la celda de rejilla de cada propiedad — solo para
// decidir el zoom (ver arriba).
function celdasRejillaParaZoom_(points, center, zoom, mapPxW, mapPxH, escalaPxAPt, R) {
  const ocupadas = new Set();
  return points.map(p => {
    const px = proyectarPuntoEnMapa_(p.lat, p.lng, center.lat, center.lng, zoom, mapPxW, mapPxH);
    const x = px.x * escalaPxAPt, y = px.y * escalaPxAPt;
    const ideal = hexRedondear_((2 / 3 * x) / R, (-1 / 3 * x + Math.sqrt(3) / 3 * y) / R);
    const celda = encontrarCeldaLibre_(ideal.q, ideal.r, ocupadas);
    ocupadas.add(celda.q + "_" + celda.r);
    return { cx: R * (3 / 2 * celda.q), cy: R * (Math.sqrt(3) / 2 * celda.q + Math.sqrt(3) * celda.r) };
  });
}

// AJUSTE 19 (ver CAMBIO-2-CIRCULOS-MAPA-UBICACION.txt): el usuario notó que
// el hexágono se veía del MISMO tamaño en pantalla en propuestas con
// propiedades muy dispersas (varias decenas de km) y en propuestas con
// propiedades muy juntas (pocos km) — el radio (R) era una constante fija
// en puntos de la presentación, independiente del zoom. Pidió que el
// hexágono represente siempre la MISMA distancia REAL en el terreno (como
// ya hacían los círculos con su radio físico) — es decir, que se vea más
// GRANDE en pantalla cuando el mapa está más acercado (propiedades juntas)
// y más CHICO cuando está más alejado (propiedades dispersas), en vez de
// mantenerse igual siempre.
//
// HEX_RADIO_M: radio real, en metros, que representa cada hexágono — 300m
// es la misma escala ya usada y aprobada para "qué tan cerca es cerca" en
// el agrupado del mapa de círculos (MAX_CLUSTER_DISTANCIA_M, ver
// GenerarMapaCliente.js), así que dos hexágonos vecinos siguen
// representando aproximadamente esa misma idea de "vecindad real".
const HEX_RADIO_M = 300;

// Límites de legibilidad: el radio real en metros, convertido a puntos,
// puede dar un hexágono MUCHO más chico o más grande que uno usable —
// una propuesta con naves repartidas en 80+ km (zoom muy alejado) da un
// radio de menos de 2pt, donde el número queda invisible; una propuesta
// con 2 naves a 20m una de otra (zoom muy cercano) daría un hexágono
// enorme sin necesidad. HEX_RADIO_PT_MIN es el mismo tamaño ya probado y
// aprobado como "hexágono chico pero legible" (ver AJUSTE 14);
// HEX_RADIO_PT_MAX evita hexágonos exagerados en clústeres muy compactos.
// Entre esos dos límites, el tamaño SÍ varía según la escala real, que es
// lo que pidió el usuario — los límites solo evitan los dos extremos
// donde esa variación dejaría de ser legible o razonable.
const HEX_RADIO_PT_MIN = 13;
const HEX_RADIO_PT_MAX = 26;

// Convierte HEX_RADIO_M a puntos de la presentación temporal, PARA EL ZOOM
// Y LATITUD dados — a diferencia de la versión anterior (constante), esto
// hay que recalcularlo cada vez que cambia el zoom candidato, porque
// cuántos metros representa un píxel del mapa depende del zoom (fórmula
// estándar de Web Mercator, la misma que usa Google Maps) y, en menor
// medida, de la latitud.
function radioHexEnPuntosTemp_(zoom, lat, mapPxH) {
  const metrosPorPixel = 156543.03392 * Math.cos(lat * Math.PI / 180) / Math.pow(2, zoom);
  const radioPxMapaBase = HEX_RADIO_M / metrosPorPixel;
  const escalaPxAPt = TEMP_PRES_ALTO_PT / mapPxH;
  const radioReal = radioPxMapaBase * escalaPxAPt;
  return Math.max(HEX_RADIO_PT_MIN, Math.min(HEX_RADIO_PT_MAX, radioReal));
}

// El zoom del mapa se elegía según la posición REAL de las propiedades
// (computeZoomForBounds_, pensado para círculos, que no se mueven de su
// posición exacta) — pero un hexágono puede DESPLAZARSE de su posición real
// si se solaparía con otro, y ese desplazamiento puede empujarlo más allá
// del encuadre que se calculó ANTES de saber que se iba a mover, cortando el
// hexágono en el borde del mapa. Se prueba el mismo zoom que ya se había
// calculado y, si algún hexágono queda fuera de un margen de seguridad, se
// aleja un nivel de zoom y se vuelve a probar — así el encuadre final
// contempla dónde termina cada hexágono después del desplazamiento, no solo
// dónde está la propiedad en la realidad. R se recalcula EN CADA ITERACIÓN
// (ver AJUSTE 19) porque depende del zoom que se está probando.
//
// Este criterio usa la rejilla de celdas (celdasRejillaParaZoom_) — ver el
// bloque "CRITERIO DE ZOOM" más arriba: es el mismo de los ajustes 18-25 y se
// conserva a propósito para que la escala no cambie.
function encontrarZoomHexagonosContiguos_(points, center, mapPxW, mapPxH, zoomInicial) {
  const pageH = TEMP_PRES_ALTO_PT; // mismo alto por defecto que usa SlidesApp.create() en el render real
  const pageW = pageH * (mapPxW / mapPxH);
  const escalaPxAPt = pageH / mapPxH;

  for (let zoom = zoomInicial; zoom >= 1; zoom--) {
    const R = radioHexEnPuntosTemp_(zoom, center.lat, mapPxH);
    const margenSeguro = R * 1.5; // medio hexágono extra de aire, para que ni el borde del hexágono más externo toque el borde del mapa
    const asignaciones = celdasRejillaParaZoom_(points, center, zoom, mapPxW, mapPxH, escalaPxAPt, R);
    const entranTodas = asignaciones.every(a =>
      a.cx >= margenSeguro && a.cx <= pageW - margenSeguro &&
      a.cy >= margenSeguro && a.cy <= pageH - margenSeguro
    );
    if (entranTodas) return zoom;
  }
  return 1; // el zoom más alejado posible — no debería llegar acá en la práctica
}

function hexARgb01_(hex) {
  const h = String(hex).replace("#", "");
  return { red: parseInt(h.substr(0, 2), 16) / 255, green: parseInt(h.substr(2, 2), 16) / 255, blue: parseInt(h.substr(4, 2), 16) / 255 };
}

// AJUSTE (velocidad, 2026-09-18 — ver CAMBIO-2-CIRCULOS-MAPA-UBICACION.txt):
// dibuja un hexágono POR PROPIEDAD (con su propio número de "Partida", nunca
// un conteo) encajado en la grilla hexagonal fija, para UNO O VARIOS mapas
// a la vez. Cada elemento de `items` es { baseMapBlob, points, center, zoom }
// y devuelve un PNG por elemento, en el mismo orden.
//
// Antes cada mapa creaba SU PROPIA presentación temporal, dibujaba cada
// hexágono con ~12 llamadas sueltas a SlidesApp, la guardaba, exportaba la
// miniatura y la mandaba a la papelera — un trabajo fijo de ~5 s POR MAPA
// que el seccionamiento en 4 cuadrantes multiplicaba por 4 (el usuario
// reportó que "se demoraba mucho"). Ahora:
//  - UNA sola presentación temporal para todos los mapas (1 slide por mapa).
//  - TODAS las formas (hexágonos + números) de todos los slides se crean en
//    UNA sola llamada Slides.Presentations.batchUpdate en vez de cientos de
//    llamadas sueltas.
//  - Las miniaturas se descargan EN PARALELO (UrlFetchApp.fetchAll).
// Se comprobó con datos reales que el PNG resultante es IDÉNTICO byte a byte
// al del método anterior (mismo dibujo, solo más rápido: con 3 cuadrantes el
// render bajó de ~14.6 s a ~7.4 s).
function renderMapasHexagonosIndividualesLote_(items, mapPxW, mapPxH) {
  if (!items.length) return [];

  const tempPres = SlidesApp.create("__TEMP_MAPA_HEXIND__" + new Date().getTime());
  const tempId = tempPres.getId();

  try {
    const slides = [tempPres.getSlides()[0]];
    slides[0].getPageElements().forEach(pe => { try { pe.remove(); } catch (e) {} });
    for (let i = 1; i < items.length; i++) slides.push(tempPres.appendSlide(SlidesApp.PredefinedLayout.BLANK));

    const pageW = tempPres.getPageWidth();
    const pageH = tempPres.getPageHeight();
    const escalaPxAPt = pageH / mapPxH;

    slides.forEach((slide, i) => {
      const bg = slide.insertImage(items[i].baseMapBlob);
      bg.setLeft(0).setTop(0).setWidth(pageW).setHeight(pageH);
    });
    const pageIds = slides.map(s => s.getObjectId());

    // Slides.Presentations.batchUpdate solo funciona sobre una versión YA
    // GUARDADA de la presentación (ver ajustarAnchosColumna_).
    tempPres.saveAndClose();

    const rojo = hexARgb01_(CIRCULO_COLOR);
    const blanco = { red: 1, green: 1, blue: 1 };
    const PT = (v) => ({ magnitude: v, unit: "PT" });

    const requests = [];
    items.forEach((it, si) => {
      // AJUSTE 19: R depende del zoom real de ESTE mapa (radio fijo en
      // METROS, no en puntos de la presentación) — ver radioHexEnPuntosTemp_.
      const R = radioHexEnPuntosTemp_(it.zoom, it.center.lat, mapPxH);
      const asignaciones = calcularAsignacionesHex_(it.points, it.center, it.zoom, mapPxW, mapPxH, escalaPxAPt, R);
      const fontSize = tamanoLetraHex_(R);
      const w = 2 * R, h = Math.sqrt(3) * R;
      // Caja de texto ancha (no del tamaño del hexágono) para que un número
      // de 2-3 dígitos no se parta — ver insertarNumeroSobreForma_.
      const anchoCaja = R * 3.4, altoCaja = R * 1.3;
      const elem = (ancho, alto, x, y) => ({
        pageObjectId: pageIds[si],
        size: { width: PT(ancho), height: PT(alto) },
        transform: { scaleX: 1, scaleY: 1, translateX: x, translateY: y, unit: "PT" }
      });

      asignaciones.forEach((a, j) => {
        // Los objectId de la API de Slides deben tener al menos 5 caracteres.
        const hexId = "hex_" + si + "_" + j, txtId = "txt_" + si + "_" + j;

        requests.push({ createShape: { objectId: hexId, shapeType: "HEXAGON", elementProperties: elem(w, h, a.cx - R, a.cy - h / 2) } });
        requests.push({ updateShapeProperties: { objectId: hexId,
          shapeProperties: {
            shapeBackgroundFill: { solidFill: { color: { rgbColor: rojo }, alpha: CIRCULO_OPACIDAD_RELLENO } },
            outline: { outlineFill: { solidFill: { color: { rgbColor: blanco } } }, weight: PT(1.25) }
          },
          fields: "shapeBackgroundFill.solidFill.color,shapeBackgroundFill.solidFill.alpha,outline.outlineFill.solidFill.color,outline.weight" } });

        requests.push({ createShape: { objectId: txtId, shapeType: "TEXT_BOX", elementProperties: elem(anchoCaja, altoCaja, a.cx - anchoCaja / 2, a.cy - altoCaja / 2) } });
        requests.push({ insertText: { objectId: txtId, text: String(a.partida) } });
        requests.push({ updateTextStyle: { objectId: txtId,
          style: { fontFamily: "Arial", fontSize: PT(fontSize), bold: true, foregroundColor: { opaqueColor: { rgbColor: blanco } } },
          fields: "fontFamily,fontSize,bold,foregroundColor" } });
        requests.push({ updateParagraphStyle: { objectId: txtId, style: { alignment: "CENTER" }, fields: "alignment" } });
        // Sin esto el número queda pegado arriba en vez de centrado en el
        // hexágono (ver insertarNumeroSobreForma_).
        requests.push({ updateShapeProperties: { objectId: txtId, shapeProperties: { contentAlignment: "MIDDLE" }, fields: "contentAlignment" } });
      });
    });

    Slides.Presentations.batchUpdate({ requests: requests }, tempId);

    const token = ScriptApp.getOAuthToken();
    const respsMiniatura = UrlFetchApp.fetchAll(pageIds.map(pid => ({
      url: "https://slides.googleapis.com/v1/presentations/" + tempId + "/pages/" + pid +
        "/thumbnail?thumbnailProperties.mimeType=PNG&thumbnailProperties.thumbnailSize=LARGE",
      headers: { Authorization: "Bearer " + token },
      muteHttpExceptions: true
    })));

    const urlsImagen = respsMiniatura.map(resp => {
      if (resp.getResponseCode() !== 200) {
        throw new Error("Slides thumbnail API falló al generar el mapa de hexágonos individuales: " + resp.getContentText().substring(0, 300));
      }
      return { url: JSON.parse(resp.getContentText()).contentUrl, muteHttpExceptions: true };
    });

    return UrlFetchApp.fetchAll(urlsImagen).map(r => r.getBlob().setContentType("image/png"));

  } finally {
    DriveApp.getFileById(tempId).setTrashed(true);
  }
}

// Versión de UN solo mapa (la usa GenerarMapaCliente() cuando no hace falta
// seccionar) — misma firma de siempre, ahora apoyada en la versión por lote.
function renderMapaHexagonosIndividuales_(baseMapBlob, points, center, zoom, mapPxW, mapPxH) {
  return renderMapasHexagonosIndividualesLote_([{ baseMapBlob: baseMapBlob, points: points, center: center, zoom: zoom }], mapPxW, mapPxH)[0];
}
