/*************************************************************
 * TABLA NATIVA DE SLIDES CON HIPERVÍNCULOS
 * Reemplaza el render de la tabla como imagen (captura) por una
 * tabla real de Slides, conservando los links de la columna REF
 * hacia las fichas técnicas. Usado por Slide ESPAÑOL.js y
 * Slide ENG/CHI.js.
 *************************************************************/

// Encabezados que deben mostrarse en la tabla del slide, buscados por NOMBRE
// (ver leerTablaConLinksPorNombre_) para no depender del orden de columnas
// de cada pestaña. "REF" debe estar incluido: es la columna con los links.
const HEADERS_TABLA_SLIDE_ES = [
  "Partida", "REF", "Zona Principal", "Sub Zona",
  "Superficie sugerida (m2)", "Superficie mínima disponible (m2) **",
  "Superficie máxima disponible (m2) **", "Precio por m2 *", "Disponibilidad"
];

const HEADERS_TABLA_SLIDE_ENG_CHI = [
  "Item / 项目", "REF / 参考编号", "State / 州", "Main zone / 主区域", "Sub-zone / 子区域",
  "Suggested Area (m2) / 建议面积", "Min Area (m2) ** / 最小面积", "Max Area (m2) ** / 最大面积",
  "Asking price / sqm (m2) * / 要价*", "Maint. / 维护费", "Availability / 可用性"
];

// Anchos de columna proporcionales — "Partida"/"Item" y "REF" solo muestran
// números o códigos cortos y no necesitan casi espacio; "Zona"/"Sub Zona"/
// "State"/"Main zone"/"Sub-zone" llevan nombres de lugares más largos y se
// ven apretados con el ancho parejo por defecto.
//
// El valor de PESO_COLUMNA_ANGOSTA cambió dos veces:
//  - 2026-09-16 (1er ajuste): bajado a 0.33 ("un tercio"), a pedido del
//    usuario, cuando la letra de la tabla todavía era chica (8-9pt).
//  - 2026-09-16 (2do ajuste, mismo día): subido a 0.55. Al agrandar la
//    letra en este mismo ajuste (ver FONT_MIN más abajo) un tercio de ancho
//    ya no alcanzaba para que "REF" (p.ej. "N3067") ni el encabezado
//    "Partida" cupieran en una sola línea — se veían cortados en dos
//    renglones ("Part/ida", "N30/67"). 0.55 es el valor mínimo con el que,
//    probado contra datos reales, ya no se corta ningún código de REF
//    típico (hasta 6-7 caracteres) a 15-16pt. Sigue siendo claramente más
//    angosta que las columnas normales (1) y las de zona (1.6).
const PESO_COLUMNA_ANGOSTA = 0.55;
const PESO_COLUMNA_ANCHA = 1.6;
const PESO_COLUMNA_NORMAL = 1;

// PESO_COLUMNA_ANGOSTA (0.55) se probó contra REF de hasta 6-7 caracteres
// (ver comentario arriba). Con la pestaña de ejemplo "EJEMPLO EDOMEX"
// aparecieron REF de 9 caracteres ("EDOMEX-01") que ya no entraban en una
// sola línea con ese ancho — la columna se partía en 2 renglones ("EDOMEX"
// / "-01") y deformaba la fila (visto el 2026-09-17). Agregado: cuando
// ALGÚN valor real de REF de la tabla supera UMBRAL_REF_LARGO_CHARS
// caracteres, la columna REF pasa a PESO_COLUMNA_REF_LARGA (más ancha) y
// las 3 columnas de "Superficie"/"Area" — que solo muestran números cortos
// y sobraba espacio, se ve en la captura — ceden ese ancho pasando a
// PESO_COLUMNA_SUPERFICIE_RECORTADA. El resto de columnas (Zona, Sub Zona,
// Precio, Disponibilidad) no cambia. Se decide sobre TODOS los datos de la
// tabla, no por página, para que —igual que el tamaño de letra
// (filasParaCalculo, ver AJUSTE 13)— todas las páginas de una misma tabla
// usen el mismo ancho de columna.
const UMBRAL_REF_LARGO_CHARS = 5;
const PESO_COLUMNA_REF_LARGA = 0.9;
const PESO_COLUMNA_SUPERFICIE_RECORTADA = 0.88;

function esColumnaSuperficie_(header) {
  const k = makeKeyTabla_(header);
  return k.includes("superficie") || k.includes("area (m2)");
}

function calcularRefEsLargo_(filas, refColOffset) {
  return filas.some(f => String(f.values[refColOffset] || "").trim().length > UMBRAL_REF_LARGO_CHARS);
}

function pesoColumna_(header, refEsLargo) {
  const k = makeKeyTabla_(header);
  if (k.startsWith("partida") || k.startsWith("item")) return PESO_COLUMNA_ANGOSTA;
  if (k.startsWith("ref")) return refEsLargo ? PESO_COLUMNA_REF_LARGA : PESO_COLUMNA_ANGOSTA;
  if (k.includes("zona") || k.includes("zone") || k.includes("state")) return PESO_COLUMNA_ANCHA;
  if (refEsLargo && esColumnaSuperficie_(header)) return PESO_COLUMNA_SUPERFICIE_RECORTADA;
  return PESO_COLUMNA_NORMAL;
}

// Color destacado para los valores de "Superficie sugerida (m2)" — a
// pedido del usuario (2026-09-16), esa columna se resalta en negrita y en
// este azul en vez del gris estándar del resto de la tabla. Se detecta por
// NOMBRE de encabezado (igual que pesoColumna_) para que funcione tanto en
// la tabla en español ("Superficie sugerida (m2)") como en la de inglés/
// chino ("Suggested Area (m2) / 建议面积"), sin importar el orden de
// columnas de cada pestaña.
const COLOR_SUPERFICIE_SUGERIDA = "#0659c0";
function esColumnaSuperficieSugerida_(header) {
  const k = makeKeyTabla_(header);
  return k.startsWith("superficie sugerida") || k.startsWith("suggested area");
}

// El servicio simplificado SlidesApp de Apps Script NO tiene forma de fijar
// el ancho de una columna de tabla (se comprobó: Table.setColumnWidth no
// existe). La única manera real es una llamada de bajo nivel a la propia
// API de Slides (Advanced Service "Slides", habilitada en el manifest) con
// una solicitud "updateTableColumnProperties" — y esa llamada solo funciona
// sobre una versión YA GUARDADA de la presentación, por eso debe invocarse
// después de pres.saveAndClose(), nunca antes (se verificó: llamarla antes
// de guardar falla con "The object (...) could not be found").
function ajustarAnchosColumna_(presId, tableObjectId, headers, boxWidth, refEsLargo) {
  const pesos = headers.map(h => pesoColumna_(h, refEsLargo));
  const totalPeso = pesos.reduce((a, b) => a + b, 0);
  const requests = headers.map((h, i) => ({
    updateTableColumnProperties: {
      objectId: tableObjectId,
      columnIndices: [i],
      tableColumnProperties: {
        columnWidth: { magnitude: boxWidth * (pesos[i] / totalPeso), unit: "PT" }
      },
      fields: "columnWidth"
    }
  }));
  Slides.Presentations.batchUpdate({ requests: requests }, presId);
}

// Máximo de filas de DATOS por slide de tabla. Si hay más, se divide en
// varios slides (ver insertarTablaPaginada_) en vez de encoger la letra
// hasta hacerla ilegible — agregado el 2026-09-16 a pedido del usuario tras
// ver una tabla de 25 filas apretada en un solo slide.
//
// Bajado de 15 a 13 el mismo día (ver AJUSTE 4): con leyenda, 15 filas de
// datos + encabezado ya no entraban en el espacio reservado (400pt) sin
// que la fila de datos cayera por debajo del piso legible de 10pt —
// Table.getHeight() no existe para medir el alto real ya renderizado, así
// que en vez de forzar filas apretadas por debajo de ese piso, se prefirió
// bajar el máximo de filas por slide.
//
// Bajado de nuevo, de 13 a 9, el mismo día (ver AJUSTE 5): el usuario pidió
// letra "al menos 3 puntos más grande" en las páginas de tabla con muchas
// filas — con 13 filas/slide, el cálculo de tamaño de letra caía en la
// rama de "no alcanza el espacio" (12pt, el piso normal). Con 9 filas/slide
// el mismo cálculo YA NO necesita esa rama: da lugar de sobra para el
// tamaño máximo de letra (FONT_MAX = 16pt, igual que las tablas cortas),
// que es el mismo tamaño ya usado y aprobado para tablas de pocas filas —
// no un valor nuevo sin probar. Costo: una tabla grande (ej. 57 filas) pasa
// de 5 a 7 slides de tabla.
//
// Subido de 9 a 10 el mismo día, a pedido explícito del usuario (menos
// slides de tabla). Con 10 filas/slide el cálculo de tamaño de letra sigue
// cayendo justo en el límite exacto de la rama normal (no la de "no
// alcanza el espacio") — da 14pt en vez del máximo de 16pt del ajuste
// anterior, pero sigue siendo claramente legible. Costo: una tabla de 57
// filas pasa de 7 a 6 slides de tabla.
//
// Subido de 10 a 15 el mismo día (ver AJUSTE 10): el usuario pidió
// puntualmente 15 y autorizó de antemano las 2 concesiones necesarias para
// que entrara sin pisarse — mover el título más arriba y aceptar una fila
// más baja. Ver moverTituloArribaParaTabla_() más abajo (se mueve el
// título real del slide, no solo el recuadro de la tabla) y el ajuste de
// tableBoxBase en "Slide ESPAÑOL.js"/"Slide ENG/CHI.js". Con 15 filas el
// cálculo de letra da justo al límite de la rama "no alcanza el espacio"
// (~12pt, el piso normal) — se probó con datos reales antes de dejarlo así.
const FILAS_POR_SLIDE_TABLA = 15;

// Cuánto se sube el título del slide de tabla (y se recorta el margen
// inferior hacia el pie de página) para ganarle espacio vertical a la
// tabla — ver moverTituloArribaParaTabla_ y AJUSTE 10. Elegido para que el
// título no llegue a chocar con el logo de la esquina superior derecha
// (se verificó la posición real de ambos elementos en la plantilla antes
// de fijar este valor) y para dejar un margen mínimo de aire (~16pt) antes
// del pie de página con el contacto.
// Subidos de 45/20 a 56/30 el 2026-09-16 (AJUSTE 12), para ganarle a la
// tabla el espacio adicional que hacía falta para que 15 filas llegara a
// ~14pt en vez de quedarse en 12pt (ver FONT_MIN/FONT_MAX arriba). 56pt de
// título es prácticamente el máximo seguro: el título real queda con su
// borde superior a ~5pt del borde del slide (61.23 - 56 ≈ 5), sin llegar a
// tocarlo. 30pt de recorte inferior deja ~6pt de aire antes del pie de
// página con el contacto — ajustado, pero se verificó visualmente que no
// se pisan.
const DESPLAZAR_TITULO_PT = 56;
const RECORTAR_MARGEN_INFERIOR_PT = 30;

// Busca el elemento de texto más alto del slide (el título — se identifica
// por ser una figura con texto no vacío ubicada arriba de todo, ya que no
// tiene una "descripción" propia para buscarlo por nombre) y lo sube
// DESPLAZAR_TITULO_PT puntos. Se llama ANTES de duplicar el slide para las
// páginas siguientes de la tabla (ver insertarTablaPaginada_), así el
// título ya movido se copia automáticamente a cada página nueva.
function moverTituloArribaParaTabla_(slide, deltaY) {
  const candidatos = slide.getPageElements().filter(pe => {
    if (pe.getTop() >= 200) return false;
    try { return pe.asShape().getText().asString().trim().length > 0; }
    catch (e) { return false; }
  });
  if (!candidatos.length) return;
  const titulo = candidatos[0];
  titulo.setTop(titulo.getTop() - deltaY);
}

// Lo normal es una tabla de 5-30 filas por propuesta. Por encima de este
// umbral se generan varios slides de tabla adicionales antes del mapa, así
// que se avisa cuántos y se deja cancelar en vez de agregarlos en silencio.
const MAX_FILAS_TABLA_SLIDE_RECOMENDADO = 40;

function advertirSiTablaGrande_(ui, numFilas) {
  if (numFilas <= MAX_FILAS_TABLA_SLIDE_RECOMENDADO) return true;
  const numSlides = Math.ceil(numFilas / FILAS_POR_SLIDE_TABLA);
  const resp = ui.alert(
    "Tabla muy grande",
    `Esta tabla tiene ${numFilas} filas (lo normal es 5-30) y se va a repartir en ${numSlides} slides de tabla antes de continuar con el resto de la propuesta.\n\n¿Deseas continuar de todas formas?`,
    ui.ButtonSet.YES_NO
  );
  return resp === ui.Button.YES;
}

function stripAccentsTabla_(s) {
  return String(s == null ? "" : s).normalize("NFD").replace(/[̀-ͯ]/g, "");
}
function makeKeyTabla_(s) {
  return stripAccentsTabla_(s).trim().toLowerCase().replace(/\s+/g, " ");
}

// Lee la tabla final (headers + filas) desde el Sheet buscando cada encabezado
// deseado POR NOMBRE (no por posición). Esto blinda la lectura contra pestañas
// antiguas donde el orden de columnas cambió con el tiempo (p.ej. "Zona
// Principal" quedó después de la columna separadora en tablas viejas) —
// aparecerá igual siempre que el encabezado exista con ese nombre en la fila.
// Captura además los hipervínculos reales de la columna REF (runs de rich
// text, para soportar celdas con varias REF agrupadas).
function leerTablaConLinksPorNombre_(sheet, headerRow, colStart, maxColsProbe, headersDeseados, refHeaderName) {
  const headerProbe = sheet.getRange(headerRow, colStart, 1, maxColsProbe).getDisplayValues()[0];
  const headerMap = new Map();
  headerProbe.forEach((h, i) => {
    const k = makeKeyTabla_(h);
    if (k && !headerMap.has(k)) headerMap.set(k, i);
  });

  const headers = [];
  const colOffsets = [];
  headersDeseados.forEach(h => {
    const idx = headerMap.get(makeKeyTabla_(h));
    if (idx !== undefined) { headers.push(h); colOffsets.push(idx); }
  });

  const refColOffset = headers.findIndex(h => makeKeyTabla_(h) === makeKeyTabla_(refHeaderName));
  if (refColOffset === -1) throw new Error(`No encontré la columna "${refHeaderName}" en el encabezado de la tabla.`);
  const partidaSheetOffset = colOffsets[0];
  const refSheetOffset = colOffsets[refColOffset];

  // Antes esto se leía fila por fila (2 llamadas a la hoja por cada fila:
  // una para los valores, otra para el rich text de REF) — con una tabla de
  // 25-28 filas eso eran 50+ llamadas a Sheets y varios segundos solo en
  // leer (medido: ~4.2s de los ~15s totales para generar una tabla de 28
  // filas). Ahora se lee TODO el bloque de filas de un solo golpe (2
  // llamadas en total, sin importar cuántas filas tenga la tabla) y se
  // recorre en memoria — mismo resultado, mucho más rápido. Ajustado el
  // 2026-09-16 tras medir que el proceso de generar se sentía lento.
  const filas = [];
  const lastRow = sheet.getLastRow();
  const numRowsDisponibles = lastRow - headerRow;

  if (numRowsDisponibles > 0) {
    const bloqueValores = sheet.getRange(headerRow + 1, colStart, numRowsDisponibles, maxColsProbe).getDisplayValues();
    const bloqueRefRich = sheet.getRange(headerRow + 1, colStart + refSheetOffset, numRowsDisponibles, 1).getRichTextValues();

    for (let i = 0; i < numRowsDisponibles; i++) {
      const rowProbe = bloqueValores[i];
      // BUG encontrado y corregido el 2026-09-16 (venía de ANTES de todos
      // estos cambios, heredado de la función original readFinalTableForSlides_INT
      // que esta reemplazó): antes también se cortaba la tabla en la primera
      // fila con la columna REF vacía, como si fuera el fin de la tabla. Una
      // REF vacía es un estado válido (una propiedad sin ficha/REF asignada
      // todavía), no el fin de los datos — se comprobó con "CTS LOGISTICS"
      // que la fila 29 (de 36 reales) tiene REF vacía y por eso la tabla se
      // leía como si tuviera solo 28 filas, cortando 8 filas reales. Ahora
      // solo se corta cuando "Partida" deja de ser un número consecutivo
      // (esa sí es la señal real de haber llegado al final de la tabla).
      if (!/^\d+$/.test(rowProbe[partidaSheetOffset])) break;

      const values = colOffsets.map(off => rowProbe[off]);

      const refRich = bloqueRefRich[i][0];
      let refRuns = [];
      if (refRich) {
        refRich.getRuns().forEach(run => {
          refRuns.push({ text: run.getText(), url: run.getLinkUrl() || "" });
        });
      }
      if (!refRuns.length) refRuns = [{ text: values[refColOffset] || "", url: "" }];

      // "Seleccionar..." es el texto del menú desplegable de una fila
      // agrupada cuando todavía no se eligió una REF específica (ver
      // TablaParques.js) — no es información real para el cliente, y además
      // al ser un texto largo dentro de la columna angosta de REF obligaba a
      // esa fila a partirse en 2 líneas y crecer más de lo calculado
      // (encontrado el 2026-09-16 al investigar por qué la leyenda de abajo
      // se pisaba con la tabla). Se deja en blanco en vez de mostrarlo.
      // NOTA: esta celda SIEMPRE trae "runs" de rich text aunque no tenga
      // ningún link (Sheets devuelve un run con la celda entera incluso para
      // un valor puesto con .setValue() sin formato), así que hay que
      // revisar el texto YA ARMADO — filtrar antes de este punto (como se
      // hizo en un primer intento) no alcanza, porque nunca cae en la rama
      // de "sin runs".
      const refTextoCompleto = refRuns.map(r => r.text).join("").trim();
      if (refTextoCompleto === "Seleccionar...") refRuns = [{ text: "", url: "" }];

      filas.push({ values: values, refRuns: refRuns });
    }
  }

  return { headers: headers, filas: filas, refColOffset: refColOffset };
}

// Inserta (o reemplaza) una tabla nativa de Slides con el mismo estilo
// visual que antes tenía la imagen generada (header azul marino, filas
// alternadas, Arial), pero con la columna REF como texto real e
// hipervínculo activo hacia la ficha técnica.
//
// filasParaCalculo (opcional, agregado el 2026-09-16 — AJUSTE 13): cuántas
// filas usar para CALCULAR el tamaño de letra/alto de fila, que puede ser
// distinto de filas.length (las filas que de verdad se dibujan). Se
// necesita para tablas de VARIAS páginas: si cada página calculara su
// propio tamaño según SU cantidad de filas, la última página (casi
// siempre con menos filas que las demás) terminaba con letra más grande —
// y como esa letra más grande hacía que la tabla ocupara TODO el alto
// disponible, la leyenda de esa página quedaba pegada o directamente
// pisada, mientras que en las demás páginas (con el recuadro completo)
// había aire de sobra. Pasándole SIEMPRE FILAS_POR_SLIDE_TABLA como
// filasParaCalculo (ver insertarTablaPaginada_ más abajo) todas las
// páginas de una misma tabla quedan con el mismo tamaño de letra, sin
// importar cuántas filas tenga cada una en particular. Si no se pasa,
// se usa filas.length (comportamiento de siempre para tablas de una sola
// página, donde SÍ se quiere que una tabla corta se vea compacta y con
// letra más grande).
function insertarTablaNativaConLinks_(slide, pres, headers, filas, refColOffset, box, desc, filasParaCalculo) {
  slide.getPageElements().forEach(pe => {
    try { if (pe.getDescription() === desc) pe.remove(); } catch (e) {}
  });

  const numFilasCalculo = filasParaCalculo == null ? filas.length : filasParaCalculo;
  const numCols = headers.length;
  const numRows = filas.length + 1;

  const HEADER_BG = "#1a3a5c";
  const ROW_BG_EVEN = "#f0f4f8";
  const ROW_BG_ODD = "#ffffff";
  const LINK_COLOR = "#1155cc";
  const TEXT_COLOR = "#222222";

  // El tamaño de letra se calcula primero según cuánto espacio hay por fila
  // (hasta 16pt cuando hay pocas filas y sobra espacio), y solo se reduce
  // hacia el mínimo legible conforme la tabla tiene más filas. El alto de
  // fila sigue al tamaño de letra elegido, así que una tabla con pocas filas
  // queda compacta (no estirada) pero con texto grande, y solo se estira
  // hasta llenar el cuadro cuando de verdad hace falta el espacio.
  //
  // FONT_MIN subido de 8 a 12 el 2026-09-16 (el usuario pidió al menos 4pt
  // más de letra tras ver una tabla de 25 filas en un solo slide). Ahora,
  // combinado con FILAS_POR_SLIDE_TABLA (ver más abajo), ningún slide de
  // tabla debería tener más de 15 filas de datos, así que en la práctica
  // casi siempre se usa el máximo de 16pt — este mínimo de 12 solo entra
  // en juego si de todas formas se genera una tabla más grande (el usuario
  // eligió continuar pese a la advertencia de tabla muy grande).
  // BUG encontrado el 2026-09-16 mientras se probaba el ajuste de la leyenda
  // (ver AJUSTE 4 en CAMBIO-1-TABLA-NATIVA-SLIDES.txt): este cálculo asumía
  // que las 16 "filas" de la tabla (15 de datos + encabezado) medían todas
  // lo mismo, pero encabezados largos como "Superficie mínima disponible
  // (m2) **" no entran en una sola línea dentro de su columna (angosta o
  // normal) y Slides estira esa fila a 2 líneas por su cuenta — sin avisar
  // y sin que Table.getHeight() lo refleje (sigue devolviendo null). Eso
  // empujaba el resto de la tabla más abajo de lo calculado y la hacía
  // pisar la leyenda incluso con el margen de seguridad de 150pt. Ahora se
  // presupuestan 2 "unidades" de fila para el encabezado (HEADER_ROW_UNITS)
  // en vez de 1, y ese alto se fuerza de verdad después con
  // ajustarAltosFila_ (mismo mecanismo de bajo nivel que ajustarAnchosColumna_
  // para el ancho), en vez de solo esperar que Slides lo respete.
  const HEADER_ROW_UNITS = 2;
  const numRowUnits = numFilasCalculo + HEADER_ROW_UNITS;
  // FONT_MIN/FONT_MAX subidos de 12/16 a 14/18 el 2026-09-16 (AJUSTE 12):
  // el primer intento de "+1pt" (ver AJUSTE 11) lo aplicaba solo AL
  // RENDERIZAR el texto, sin tocar este cálculo — pero para el caso de 15
  // filas, dataFontSize ya caía en la rama de abajo ("no alcanza el
  // espacio"), que fija rowH según el espacio disponible SIN mirar qué
  // fuente se quería; ajustarAltosFila_ terminaba fijando una fila
  // calculada para 12pt y Slides encogía el texto de 13pt de vuelta para
  // que entrara — el cambio nunca se veía. La fuente deseada tiene que
  // alimentar este cálculo (no aplicarse después) para que el alto de fila
  // realmente crezca con ella. Se liberó espacio adicional (ver
  // DESPLAZAR_TITULO_PT, RECORTAR_MARGEN_INFERIOR_PT y LEGEND_RESERVA_PT
  // más abajo) para que 15 filas alcance ~14pt en vez de quedarse en 12.
  const FONT_MIN = 14, FONT_MAX = 18, ROW_PADDING = 2.3;
  let dataFontSize = Math.max(FONT_MIN, Math.min(FONT_MAX, box.height / numRowUnits / ROW_PADDING));
  let rowH = dataFontSize * ROW_PADDING;
  if (rowH * numRowUnits > box.height) {
    rowH = box.height / numRowUnits;
    dataFontSize = Math.max(10, rowH / ROW_PADDING);
  }
  dataFontSize = Math.round(dataFontSize);
  const headerFontSize = dataFontSize + 1;
  const headerRowH = rowH * HEADER_ROW_UNITS;
  const tableHeight = rowH * filas.length + headerRowH;
  // Si se pasó filasParaCalculo (tabla de varias páginas, ver más arriba),
  // la tabla se ALINEA ARRIBA en vez de centrarse — así el encabezado
  // queda a la misma altura en todas las páginas, y el espacio que sobra
  // en una página con menos filas que el máximo queda TODO abajo, justo
  // donde hace falta el aire antes de la leyenda, en vez de repartirse
  // mitad arriba/mitad abajo sin necesidad. Para una tabla de una sola
  // página (filasParaCalculo no se pasa) se sigue centrando, como siempre
  // — ahí sí se quiere que una tabla corta se vea compacta y pareja.
  const tableTop = filasParaCalculo == null
    ? box.top + (box.height - tableHeight) / 2
    : box.top;

  const table = slide.insertTable(numRows, numCols, box.left, tableTop, box.width, tableHeight);
  table.setDescription(desc);

  // Un TextRange vacío ("") no admite getTextStyle()/getParagraphStyle() en Slides
  // (lanza "The object (...) has no text"), así que solo se aplica estilo si hay texto.
  const setCellText = (tr, text, fontSize) => {
    tr.setText(text);
    if (text.length === 0) return;
    tr.getTextStyle().setFontFamily("Arial").setFontSize(fontSize).setForegroundColor(TEXT_COLOR);
    tr.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
  };

  for (let c = 0; c < numCols; c++) {
    const cell = table.getCell(0, c);
    cell.getFill().setSolidFill(HEADER_BG);
    const tr = cell.getText();
    const headerText = String(headers[c] == null ? "" : headers[c]);
    tr.setText(headerText);
    if (headerText.length > 0) {
      tr.getTextStyle().setFontFamily("Arial").setFontSize(headerFontSize).setBold(true).setForegroundColor("#ffffff");
      tr.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    }
  }

  filas.forEach((fila, rIdx) => {
    const rowInTable = rIdx + 1;
    const bg = (rIdx % 2 === 0) ? ROW_BG_EVEN : ROW_BG_ODD;

    for (let c = 0; c < numCols; c++) {
      const cell = table.getCell(rowInTable, c);
      cell.getFill().setSolidFill(bg);
      const tr = cell.getText();

      if (c === refColOffset && fila.refRuns.length) {
        const fullText = fila.refRuns.map(run => run.text).join("");
        setCellText(tr, fullText, dataFontSize);

        let cursor = 0;
        fila.refRuns.forEach(run => {
          const len = run.text.length;
          if (len > 0 && run.url) {
            const runRange = tr.getRange(cursor, cursor + len);
            runRange.getTextStyle().setLinkUrl(run.url).setForegroundColor(LINK_COLOR).setUnderline(true);
          }
          cursor += len;
        });
      } else {
        const value = fila.values[c];
        const texto = String(value == null ? "" : value);
        setCellText(tr, texto, dataFontSize);
        if (texto.length > 0 && esColumnaSuperficieSugerida_(headers[c])) {
          tr.getTextStyle().setBold(true).setForegroundColor(COLOR_SUPERFICIE_SUGERIDA);
        }
      }
    }
  });

  return {
    table: table, rowH: rowH, headerRowH: headerRowH,
    dataFontSize: dataFontSize, headerFontSize: headerFontSize, tableTop: tableTop
  };
}

// ---------------------------------------------------------------------------
// ESTIMACIÓN DEL ALTO REAL DE LA TABLA (para pegar la leyenda justo debajo)
// ---------------------------------------------------------------------------
// Slides deja crecer una fila más allá del mínimo pedido cuando su texto se
// parte en 2+ líneas (el encabezado "Superficie mínima disponible (m2) **" en
// 4 líneas, o "Inmediata 23k + BTS" en 2), y NO informa ese alto real: se
// comprobó (2026-09-18) que la API devuelve rowHeight == minRowHeight aunque
// la fila se vea más alta, y Table.getHeight() da null. Como el alto real se
// desvía del nominal (en la última página de "CTS LOGISTICS": ~39pt más), no
// se puede colocar la leyenda a "alto nominal + margen" sin riesgo de pisarla.
// Se ESTIMA contando las líneas de cada celda. Modelo calibrado con una
// captura real (Arial): interlineado ≈ 1.17 × tamaño de letra y márgenes de
// celda ≈ 7.2pt por lado (encabezado de 4 líneas a 15pt ≈ 85pt; fila de 2
// líneas a 14pt ≈ 47pt).
const CELDA_MARGEN_H_PT = 14.4;   // izquierdo + derecho
const CELDA_MARGEN_V_PT = 14.4;   // superior + inferior
const INTERLINEADO_TABLA = 1.17;
// Se sobrestima un poco el ancho del texto: equivocarse por de más solo deja
// un poco más de aire, equivocarse por de menos haría que la leyenda pise la
// tabla.
const SEGURIDAD_ANCHO_TEXTO = 1.05;

// Anchos de carácter de Arial (por 1000 unidades de em) para lo que suele
// llevar la tabla; lo demás cae en 556 (ancho de un dígito/minúscula ancha).
const ANCHO_ARIAL_ = (function () {
  const w = { " ": 278, ".": 278, ",": 278, ":": 278, ";": 278, "/": 278, "-": 333, "(": 333, ")": 333, "*": 389, "+": 584, "%": 889 };
  "0123456789".split("").forEach(c => w[c] = 556);
  const may = { A: 667, B: 667, C: 722, D: 722, E: 667, F: 611, G: 778, H: 722, I: 278, J: 500, K: 667, L: 556, M: 833, N: 722, O: 778, P: 667, Q: 778, R: 722, S: 667, T: 611, U: 722, V: 667, W: 944, X: 667, Y: 667, Z: 611 };
  const min = { a: 556, b: 556, c: 500, d: 556, e: 556, f: 278, g: 556, h: 556, i: 222, j: 222, k: 500, l: 222, m: 833, n: 556, o: 556, p: 556, q: 556, r: 333, s: 500, t: 278, u: 556, v: 500, w: 722, x: 500, y: 500, z: 500 };
  Object.keys(may).forEach(k => w[k] = may[k]);
  Object.keys(min).forEach(k => w[k] = min[k]);
  return w;
})();

function anchoTextoPt_(texto, fontSize, negrita) {
  const sinAcentos = stripAccentsTabla_(texto);
  let unidades = 0;
  for (let i = 0; i < sinAcentos.length; i++) unidades += (ANCHO_ARIAL_[sinAcentos[i]] || 556);
  return unidades / 1000 * fontSize * (negrita ? 1.08 : 1) * SEGURIDAD_ANCHO_TEXTO;
}

// Cuántas líneas ocupa `texto` dentro de un ancho disponible, cortando por
// palabras (una palabra más ancha que el espacio se parte por caracteres).
function contarLineasCelda_(texto, anchoDisponible, fontSize, negrita) {
  const t = String(texto == null ? "" : texto).trim();
  if (!t) return 1;
  let lineas = 1, actual = 0;
  const espacio = anchoTextoPt_(" ", fontSize, negrita);
  t.split(/\s+/).forEach(palabra => {
    const w = anchoTextoPt_(palabra, fontSize, negrita);
    if (w > anchoDisponible) {
      const extra = Math.ceil(w / anchoDisponible);
      lineas += (actual > 0 ? 1 : 0) + extra - 1;
      actual = w - (extra - 1) * anchoDisponible;
    } else if (actual === 0) {
      actual = w;
    } else if (actual + espacio + w <= anchoDisponible) {
      actual += espacio + w;
    } else {
      lineas++;
      actual = w;
    }
  });
  return lineas;
}

// Alto estimado (pt) de la tabla ya renderizada: suma, por fila, el mayor
// entre el alto mínimo pedido y el que necesita su celda más alta.
function estimarAltoRealTabla_(headers, filas, refColOffset, boxWidth, refEsLargo, dataFontSize, headerFontSize, rowH, headerRowH) {
  const pesos = headers.map(h => pesoColumna_(h, refEsLargo));
  const totalPeso = pesos.reduce((a, b) => a + b, 0);
  const anchos = pesos.map(p => boxWidth * p / totalPeso - CELDA_MARGEN_H_PT);

  let maxLineasHeader = 1;
  headers.forEach((h, c) => {
    maxLineasHeader = Math.max(maxLineasHeader, contarLineasCelda_(h, anchos[c], headerFontSize, true));
  });
  let total = Math.max(headerRowH, maxLineasHeader * headerFontSize * INTERLINEADO_TABLA + CELDA_MARGEN_V_PT);

  filas.forEach(fila => {
    let maxLineas = 1;
    headers.forEach((h, c) => {
      const texto = (c === refColOffset && fila.refRuns.length)
        ? fila.refRuns.map(r => r.text).join("")
        : fila.values[c];
      const negrita = esColumnaSuperficieSugerida_(h);
      maxLineas = Math.max(maxLineas, contarLineasCelda_(texto, anchos[c], dataFontSize, negrita));
    });
    total += Math.max(rowH, maxLineas * dataFontSize * INTERLINEADO_TABLA + CELDA_MARGEN_V_PT);
  });
  return total;
}

// Fuerza el alto real de cada fila de la tabla (encabezado más alto que las
// filas de datos, ver HEADER_ROW_UNITS en insertarTablaNativaConLinks_) con
// una llamada de bajo nivel a la API de Slides — igual que ajustarAnchosColumna_
// para el ancho de columnas, y por la misma razón: SlidesApp no tiene un
// setter de alto de fila (se comprobó: TableRow.setMinimumHeight no existe),
// y esta llamada solo funciona sobre la versión YA GUARDADA de la
// presentación, por eso debe invocarse después de pres.saveAndClose(), igual
// que ajustarAnchosColumna_.
//
// Encontrado el 2026-09-16: sin esto, Slides decide el alto real de cada
// fila según su contenido y puede estirar la fila de encabezado más de lo
// presupuestado, empujando el resto de la tabla hacia abajo y pisando la
// leyenda aunque el cálculo de alto total (tableHeight) diera un número que
// parecía correcto.
function ajustarAltosFila_(presId, tableObjectId, numFilasDatos, rowH, headerRowH) {
  const requests = [{
    updateTableRowProperties: {
      objectId: tableObjectId,
      rowIndices: [0],
      tableRowProperties: { minRowHeight: { magnitude: headerRowH, unit: "PT" } },
      fields: "minRowHeight"
    }
  }];
  for (let r = 1; r <= numFilasDatos; r++) {
    requests.push({
      updateTableRowProperties: {
        objectId: tableObjectId,
        rowIndices: [r],
        tableRowProperties: { minRowHeight: { magnitude: rowH, unit: "PT" } },
        fields: "minRowHeight"
      }
    });
  }
  Slides.Presentations.batchUpdate({ requests: requests }, presId);
}

// BUG encontrado y corregido el 2026-09-16: el pedido original de "hacer
// menos ancha Partida y REF" venía junto con otro pedido que se pasó por
// alto — agregar al slide la leyenda que ya existe DEBAJO de la tabla en el
// Sheet (las líneas que empiezan con "*" y "**", ej. "Precios y
// disponibilidades sujetos a cambios..."). Esta función la busca: escanea
// hasta 6 filas después de la última fila de datos de la tabla, en la misma
// columna donde empieza la tabla, y junta las líneas consecutivas que
// empiecen con "*" (se corta apenas aparece una fila que no empieza con "*",
// para no juntar por error el bloque de "Mapa (Drive):" que Generar
// MapaCliente escribe más abajo). No se hardcodeó una fila fija porque la
// plantilla en español deja 1 fila en blanco antes de la leyenda y la de
// inglés/chino no deja ninguna — escanear una ventana funciona para ambas
// sin duplicar la función por idioma.
function leerLeyendaTabla_(sheet, headerRow, colStart, numFilasDatos) {
  const inicioBusqueda = headerRow + numFilasDatos + 1;
  const lineas = [];
  for (let r = inicioBusqueda; r < inicioBusqueda + 6; r++) {
    const texto = String(sheet.getRange(r, colStart).getDisplayValue() || "").trim();
    if (texto.indexOf("*") === 0) {
      lineas.push(texto);
    } else if (lineas.length > 0) {
      break;
    }
  }
  return lineas;
}

// Alto reservado para la leyenda debajo de la tabla. Se resta del recuadro
// de la tabla en TODAS las páginas (no solo la última) — así el recuadro es
// igual en todas las páginas y no hace falta recalcular nada según cuál
// termine siendo la última. Corregido el 2026-09-16: la primera versión
// ponía la leyenda en una posición fija cerca del borde del slide, y en
// páginas con muchas filas la tabla llegaba a ocupar casi todo el recuadro
// y se pisaban.
//
// Subido de 42 a 100 el mismo día: se comprobó que Table.getHeight() de
// Slides devuelve null (Slides no expone de forma confiable el alto real
// YA renderizado de una tabla, que puede crecer más allá del mínimo pedido
// si alguna celda necesita más de una línea — p.ej. la palabra "Seleccionar..."
// del menú desplegable de filas agrupadas sin REF elegido, que no cabe en
// una sola línea en la columna angosta de REF). Al no poder medir el alto
// real después de crearla, se usa un margen de seguridad más generoso en
// vez de uno calculado exacto.
// Bajado de 150 a 95 el 2026-09-16 (AJUSTE 12), para liberarle espacio a
// la tabla (ver FONT_MIN/FONT_MAX arriba): con LEGEND_GAP_PT=40 y la
// leyenda a 15pt itálica en 2 líneas (~35pt de texto), el espacio
// REALMENTE necesario es de unos 80pt — 150 tenía casi el doble de margen
// de sobra del que hace falta. 95 deja todavía ~15pt de colchón sobre el
// mínimo real, sin ser tan generoso como para quitarle a la tabla el
// espacio que necesita para el tamaño de letra más grande.
const LEGEND_RESERVA_PT = 95;

// Espacio en blanco entre el borde inferior de la tabla y el cuadro de
// texto de la leyenda. Antes eran solo 6pt — el usuario lo vio "muy
// pegado" en la presentación real (2026-09-16), así que se subió a 20pt,
// y luego a 40pt el mismo día: con tablas de 13 filas de datos (el máximo
// por slide, ver FILAS_POR_SLIDE_TABLA) se comprobó que la tabla sigue
// renderizando un poco más alta de lo calculado (mismo motivo que
// FILAS_POR_SLIDE_TABLA: no hay forma de medir el alto real ya renderizado),
// y con 20pt ese sobrante se comía todo el espacio en blanco, dejando la
// leyenda pegada al borde de la tabla otra vez aunque ya no se superponía
// con el texto. 40pt deja aire visible incluso en el peor caso (13 filas).
const LEGEND_GAP_PT = 40;

// Si va a haber leyenda, devuelve una copia del recuadro con el alto
// reducido en LEGEND_RESERVA_PT; si no hay leyenda, devuelve el mismo
// recuadro sin tocar (la tabla usa todo el espacio disponible como antes).
function boxConEspacioLeyenda_(box, hayLeyenda) {
  if (!hayLeyenda) return box;
  return { left: box.left, top: box.top, width: box.width, height: box.height - LEGEND_RESERVA_PT };
}

// Agrega la leyenda (ver leerLeyendaTabla_) como un cuadro de texto chico e
// itálico, debajo del recuadro de la tabla (top = borde inferior del
// recuadro que ya se le pasó a insertarTablaPaginada_, más LEGEND_GAP_PT de
// aire) — se llama una vez POR CADA slide de tabla (el usuario pidió que
// salga en todos, no solo en el último, cuando la tabla ocupa varios
// slides).
//
// bottomEstimadoTabla (opcional): borde inferior estimado de la tabla de ESTA
// página (ver estimarAltoRealTabla_). Solo lo reciben las páginas con menos
// filas que el máximo — en ellas la leyenda se pega a la tabla (a
// LEGEND_GAP_CERCA_PT de su borde estimado) en vez de quedar en el lugar fijo
// del fondo con un gran espacio en blanco en medio (pedido del usuario,
// 2026-09-18). Nunca queda MÁS abajo que el lugar fijo de siempre.
const LEGEND_GAP_CERCA_PT = 24;
function agregarLeyendaTabla_(slide, pres, lineas, top, bottomEstimadoTabla) {
  if (!lineas.length) return;
  const desc = "AUTO_LEYENDA_TABLA";
  slide.getPageElements().forEach(pe => {
    try { if (pe.getDescription() === desc) pe.remove(); } catch (e) {}
  });

  let topLeyenda = top + LEGEND_GAP_PT;
  if (bottomEstimadoTabla != null) {
    topLeyenda = Math.min(topLeyenda, bottomEstimadoTabla + LEGEND_GAP_CERCA_PT);
  }

  const left = 70, width = pres.getPageWidth() - 140;
  const caja = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, left, topLeyenda, width, LEGEND_RESERVA_PT - LEGEND_GAP_PT - 6);
  caja.setDescription(desc);

  const tr = caja.getText();
  tr.setText(lineas.join("\n"));
  // Historial del tamaño: 9 -> 12pt (2026-09-16, "al menos 3 puntos más")
  // -> 15pt (2026-09-16, mismo día: "sigue siendo pequeña, aumenta el
  // tamaño de fuente en un 25% más" — 12 * 1.25 = 15). Sigue habiendo
  // margen de sobra dentro del recuadro reservado (LEGEND_RESERVA_PT) para
  // 2 líneas a 15pt itálica, así que no hizo falta tocar LEGEND_RESERVA_PT
  // ni LEGEND_GAP_PT.
  tr.getTextStyle().setFontFamily("Arial").setFontSize(15).setItalic(true).setForegroundColor("#555555");
  tr.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.START);
}

// Inserta la tabla dividida en tantos slides como haga falta (de a
// FILAS_POR_SLIDE_TABLA filas de datos cada uno). El primer bloque va en
// slideOriginal (el slide de la plantilla ya reservado para la tabla); para
// cada bloque adicional se DUPLICA ese mismo slide (Slide.duplicate() lo
// inserta automáticamente justo después del original, con el mismo diseño/
// título/logo) y se reemplaza su tabla por el bloque de filas que sigue.
//
// Devuelve un arreglo con {slide, objectId, rowH, headerRowH, numFilasDatos}
// por cada slide de tabla creado — se necesita después de guardar la
// presentación para poder ajustarle el ancho de columnas (ajustarAnchosColumna_)
// y el alto de filas (ajustarAltosFila_) a CADA slide por separado, ya que
// ambas solo ajustan una tabla a la vez.
//
// Importante para quien llame a esta función: hay que guardar una
// referencia al slide del MAPA (u otro contenido que vaya después de la
// tabla) ANTES de llamar aquí, tomándola directamente del arreglo de slides
// (ej. slides[6]) y no por índice numérico — como se van a insertar slides
// nuevos entre la tabla y el mapa, el índice del mapa cambia, pero la
// referencia al objeto Slide sigue apuntando al slide correcto sin importar
// cuántos slides se inserten antes.
function insertarTablaPaginada_(slideOriginal, pres, headers, filas, refColOffset, box, desc) {
  const paginas = [];
  for (let i = 0; i < filas.length; i += FILAS_POR_SLIDE_TABLA) {
    paginas.push(filas.slice(i, i + FILAS_POR_SLIDE_TABLA));
  }
  if (!paginas.length) paginas.push([]);

  // AJUSTE 13 (ver CAMBIO-1-TABLA-NATIVA-SLIDES.txt): cuando hay más de una
  // página, TODAS calculan su tamaño de letra/alto de fila como si tuvieran
  // FILAS_POR_SLIDE_TABLA filas (el máximo), no la cantidad real de CADA
  // página — así ninguna página (típicamente la última, con menos filas)
  // termina con letra más grande que las demás, que era lo que hacía que
  // esa página en particular llenara todo el recuadro y pisara la leyenda.
  // Para una tabla de una sola página (paginas.length === 1) no se pasa
  // este valor, para no perder el comportamiento de siempre (una tabla
  // corta se ve compacta, con letra más grande).
  const filasParaCalculo = paginas.length > 1 ? FILAS_POR_SLIDE_TABLA : null;

  const resultado = [];
  let slideActual = slideOriginal;
  const refEsLargo = calcularRefEsLargo_(filas, refColOffset);

  paginas.forEach((pagina, idx) => {
    if (idx > 0) slideActual = slideActual.duplicate();
    const r = insertarTablaNativaConLinks_(slideActual, pres, headers, pagina, refColOffset, box, desc, filasParaCalculo);

    // Borde inferior ESTIMADO de la tabla de esta página, solo para las
    // páginas con menos filas que el máximo (típicamente la última): ahí
    // la leyenda va justo debajo de la tabla en vez de quedarse en su lugar
    // fijo al fondo, dejando un hueco enorme (ver agregarLeyendaTabla_ y
    // estimarAltoRealTabla_). En las páginas llenas es null y no cambia nada.
    let bottomParaLeyenda = null;
    if (pagina.length < FILAS_POR_SLIDE_TABLA) {
      const alto = estimarAltoRealTabla_(headers, pagina, refColOffset, box.width, refEsLargo,
        r.dataFontSize, r.headerFontSize, r.rowH, r.headerRowH);
      bottomParaLeyenda = r.tableTop + alto;
    }

    resultado.push({
      slide: slideActual, objectId: r.table.getObjectId(), table: r.table,
      rowH: r.rowH, headerRowH: r.headerRowH, numFilasDatos: pagina.length,
      bottomParaLeyenda: bottomParaLeyenda
    });
  });

  return resultado;
}
