/**
 * @OnlyCurrentDoc
 * FUNCIÓN: Busca las fichas PDF en TODAS las subcarpetas de la carpeta maestra
 * y copia las que coinciden con los REF marcados con "ENVIAR".
 * REQUIERE: Activar "Drive API (v2)" en Servicios avanzados de Google.
 */

const FOLDER_ID_FICHAS_MAESTRA = "1E921Gm9a2wdyRYqL9ug1DUTLNxojfelb"; // Carpeta maestra donde están las fichas PDF
const FOLDER_ID_PROPUESTAS_PADRE = "11aIy3TrxtRXle-pr_275E8JkMboJrISR"; // Carpeta donde se crean las propuestas
const ENCABEZADO_SELECCION = "ENVIAR";
const ENCABEZADO_REF = "REF";

function generarFichas() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getActiveSheet();
  const nombreHoja = hoja.getName();

  // Avisa que inicia
  ui.alert(`🚀 Ejecutando en hoja: ${nombreHoja}`);

  try {
    const data = hoja.getDataRange().getValues();
    const headers = data[0];
    const idxEnviar = headers.indexOf(ENCABEZADO_SELECCION);
    const idxRef = headers.indexOf(ENCABEZADO_REF);

    if (idxEnviar === -1 || idxRef === -1) {
      ui.alert("⚠️ No se encontró alguna de las columnas requeridas (ENVIAR o REF).");
      return;
    }

    // Obtener referencias marcadas con "ok"
    const refsSeleccionadas = data
      .filter((fila, i) => i > 0 && String(fila[idxEnviar]).toLowerCase().trim() === "ok")
      .map(fila => String(fila[idxRef]).trim())
      .filter(ref => ref);

    if (refsSeleccionadas.length === 0) {
      ui.alert("⚠️ No hay referencias marcadas con 'ok' para enviar.");
      return;
    }

    ui.alert(`Paso 1/4\nBuscando fichas PDF en subcarpetas...\nSe buscarán ${refsSeleccionadas.length} referencias.`);

    // Buscar todos los PDFs en la carpeta maestra y sus subcarpetas
    const todosPDFs = getAllPdfsInFolder(FOLDER_ID_FICHAS_MAESTRA);
    ui.alert(`Paso 2/4\nSe encontraron ${todosPDFs.length} PDFs. Filtrando coincidencias...`);

    // Filtrar los PDFs que coinciden con las referencias
    const coincidencias = todosPDFs.filter(file => {
      const nombre = file.title.toLowerCase();
      return refsSeleccionadas.some(ref => nombre.includes(ref.toLowerCase()));
    });

    if (coincidencias.length === 0) {
      ui.alert("⚠️ No se encontraron fichas que coincidan con las referencias seleccionadas.");
      return;
    }

    ui.alert(`Paso 3/4\nSe encontraron ${coincidencias.length} coincidencias.\nCopiando archivos...`);

    // Crear carpeta destino
    const carpetaPropuesta = crearCarpetaPropuesta(nombreHoja);
    const copiados = new Set();
    let contador = 0;

    for (const file of coincidencias) {
      const nombreArchivo = file.title;

      if (copiados.has(nombreArchivo)) continue;
      copiados.add(nombreArchivo);

      Drive.Files.copy(file, carpetaPropuesta.getId(), { title: nombreArchivo });
      contador++;
    }

    ui.alert(`✅ Paso 4/4\nSe copiaron ${contador} archivo(s) correctamente.`);
  } catch (error) {
    ui.alert("❌ Error grave:\n" + error.message);
  }
}

/**
 * Busca recursivamente todos los PDFs en una carpeta y subcarpetas
 */
function getAllPdfsInFolder(folderId) {
  const archivos = [];
  const folder = DriveApp.getFolderById(folderId);
  const subcarpetas = folder.getFolders();

  const archivosDirectos = folder.getFilesByType(MimeType.PDF);
  while (archivosDirectos.hasNext()) {
    const file = archivosDirectos.next();
    archivos.push({ id: file.getId(), title: file.getName() });
  }

  while (subcarpetas.hasNext()) {
    const sub = subcarpetas.next();
    archivos.push(...getAllPdfsInFolder(sub.getId()));
  }

  return archivos;
}

/**
 * Crea una carpeta nueva dentro de la carpeta padre para guardar las propuestas.
 */
function crearCarpetaPropuesta(nombreHoja) {
  const carpetaPadre = DriveApp.getFolderById(FOLDER_ID_PROPUESTAS_PADRE);
  const nombreCarpeta = `${nombreHoja} ${new Date().toLocaleDateString("es-MX")}`;
  return carpetaPadre.createFolder(nombreCarpeta);
}
