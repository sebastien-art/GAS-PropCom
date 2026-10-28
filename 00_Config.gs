    /********** 00_Config.gs (ÚNICO lugar de constantes globales) **********/
  const OUTPUT_ROOT_FOLDER_ID = "13aW3gRhAHuVlF2Wqc4R1KLqe3ZucQSPX"; // PROPUESTAS CLIENTES (root) [PRUEBA]
  const MAPS_FOLDER_NAME = "Mapas Propuestas";
  const MAP_FILE_PREFIX = "MapaCliente_"; // Mapa de círculos/abanico — vuelve a ser el mapa definitivo (ver MAPA_DEFINITIVO_PREFIX).

  // Seccionamiento en 4 cuadrantes (ver CAMBIO-2-CIRCULOS-MAPA-UBICACION.txt,
  // AJUSTE "mapa seccionado"): con muchas propiedades MUY dispersas (ej.
  // "CTS LOGISTICS", 43 propiedades en ~30km), forzar todo en un solo mapa
  // hace que los hexágonos terminen en su tamaño mínimo legible (ver
  // HEX_RADIO_PT_MIN en "MapaHexagonal.js") y se vean apretados. A partir de
  // estos umbrales, GenerarMapaCliente() divide las propiedades en 4
  // cuadrantes geográficos (NE/NO/SE/SO, por área desde el centro) y genera
  // una captura de mapa POR CUADRANTE, cada una con su propio zoom (más
  // acercado, hexágonos más grandes) — ver dividirEnCuadrantes_() y
  // generarMapaHexParaPuntos_() en "GenerarMapaCliente.js". Las 4 (o menos,
  // si algún cuadrante queda vacío) se insertan de a 2 por slide en
  // "Slide ESPAÑOL.js"/"Slide ENG/CHI.js" (ver insertarMapaSeccionado_()).
  // AJUSTE 30 (ver CAMBIO-2-CIRCULOS-MAPA-UBICACION.txt): si el mapa único
  // tendría una MANCHA CONTIGUA (hexágonos pegados unos con otros sin
  // separación) de más de este tamaño, se secciona — ver
  // tamanoMayorManchaContigua_() en "MapaHexagonal.js".
  //
  // AJUSTE 32 (corrige el valor original de 10, ver
  // CAMBIO-2-CIRCULOS-MAPA-UBICACION.txt): "mancha contigua" mide hexágonos
  // TOCÁNDOSE, no SOLAPÁNDOSE — el algoritmo de separación ya garantiza 0
  // solapamiento real siempre (se verificó: "EJEMPLO EDOMEX" con mancha 12
  // da solapesReales=0 y las 20 propiedades se leen bien en el mapa único),
  // así que una mancha moderada no es automáticamente ilegible. El usuario
  // confirmó viendo la imagen real que "EJEMPLO EDOMEX" (mancha 12) SÍ cabía
  // en un solo mapa — el umbral de 10 lo seccionaba sin necesidad, igual que
  // a "PRUEBA" (también 12, mismo valor, muy probablemente el mismo caso).
  // Solo hay 2 puntos de dato REALMENTE confirmados por el usuario: 12 está
  // bien (no seccionar), 25 está mal (sí seccionar, "CTS LOGISTICS"). 20 es
  // una estimación a mitad de camino entre ambos — no un valor derivado con
  // la misma certeza que esos 2 extremos. Si alguna pestaña futura entre 13
  // y 24 se ve mal como mapa único, hay que bajar este número.
  const UMBRAL_MANCHA_CONTIGUA = 20;
  const ORDEN_CUADRANTES = ["NE", "NO", "SE", "SO"];
  const ZOOM_MAX_SECCION = 14; // tope de acercamiento de cada captura seccionada (igual que una propiedad sola)

  // AJUSTE 31 (ver CAMBIO-2-CIRCULOS-MAPA-UBICACION.txt): antes de caer a los
  // 4 cuadrantes, elegirSecciones_() (en "GenerarMapaCliente.js") prueba
  // primero si partir en solo 2 MITADES (sur/norte u oeste/este) ya resuelve
  // el apiñamiento — el usuario mostró que "CTS LOGISTICS" quedaba con una
  // sección de apenas 2 propiedades (casi vacía) cuando 2 mitades ya
  // alcanzaban. Esta lista reúne TODAS las etiquetas posibles (las de 2
  // mitades y las de 4 cuadrantes) para poder limpiar los archivos de
  // cualquier modo anterior antes de decidir de nuevo, y para que
  // getMapFilesForSheetStrict_INT() (en "Slide ENG/CHI.js") busque los
  // archivos sin importar cuál de los 2 modos se usó. El orden importa: es
  // el orden en el que se acomodan las capturas de izquierda a derecha.
  const TODAS_LAS_ETIQUETAS_SECCION = ["NE", "NO", "SE", "SO", "Sur", "Norte", "Oeste", "Este"];

  // El mapa que "Generar Slides" realmente busca e inserta en la presentación
  // (ver getMapFilesForSheetStrict_INT() en "Slide ENG/CHI.js"). Historial:
  //   - 2026-09-16 (AJUSTE 6): el usuario probó las 3 alternativas (círculos,
  //     zonas por municipio, calor) y eligió el mapa de calor.
  //   - 2026-09-16 (AJUSTE 7, mismo día): al ver el mapa de calor con datos
  //     reales, el usuario notó que el degradado de color TAPA los números
  //     de "Partida" cuando hay propiedades cercanas (justo el problema que
  //     el mapa de círculos/abanico ya resolvía separándolas en abanico).
  //     Mostró una imagen de referencia de otro sistema (círculos separados,
  //     número siempre legible, con un halo amarillo suave para marcar
  //     cercanía) y pidió volver a esa idea. Se vuelve a MAP_FILE_PREFIX
  //     (círculos/abanico, con el halo amarillo agregado — ver
  //     dibujarClusterEnAbanico_ en "GenerarMapaCliente.js") como el mapa
  //     definitivo. Para volver a otra alternativa más adelante, solo hay
  //     que cambiar este valor a "MapaZonas_" (zonas) o "MapaCalor_" (calor),
  //     sin tocar nada más.
  const MAPA_DEFINITIVO_PREFIX = MAP_FILE_PREFIX;