/***************************************************************
 * ARCHIVADO DE PROPUESTAS - PARTE 1/4
 *
 * Objetivo:
 * - Revisar pestañas de propuestas
 * - Leer fecha de creación en AA1
 * - Detectar propuestas con más de 14 días
 *
 * NO elimina hojas todavía.
 * NO mueve archivos todavía.
 *
 * Parte 2:
 * creación del Spreadsheet archivado.
 ***************************************************************/


/**************** CONFIGURACIÓN ****************/

const DIAS_PARA_ARCHIVAR = 14;

const HOJAS_EXCLUIDAS_ARCHIVO = [
  "Menú",
  "Naves"
];


/**************** FUNCIÓN PRINCIPAL ****************/

function archivarPropuestas() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojas = ss.getSheets();

  const hoy = new Date();

  const propuestasParaArchivar = [];

  hojas.forEach(hoja => {

    const nombre = hoja.getName();


    // Ignorar hojas del sistema
    if (HOJAS_EXCLUIDAS_ARCHIVO.includes(nombre)) {
      return;
    }


   // Leer fecha interna AA1
let fechaCelda = hoja.getRange("AA1").getValue();


// Si no existe AA1, buscar fecha antigua
if (!fechaCelda) {

  fechaCelda = buscarFechaCreacionAntigua_(hoja);

  if (fechaCelda) {

    // Migrar automáticamente a AA1
    hoja.getRange("AA1")
      .setValue(
        Utilities.formatDate(
          fechaCelda,
          ss.getSpreadsheetTimeZone(),
          "yyyy-MM-dd HH:mm:ss"
        )
      );

    hoja.hideColumns(27);

  } else {

    Logger.log(
      "Sin fecha de creación: " + nombre
    );

    return;

  }

}
function buscarFechaCreacionAntigua_(hoja) {

  const ultimaFila = hoja.getLastRow();
  const rango = hoja.getRange(
    1,
    1,
    ultimaFila,
    hoja.getLastColumn()
  );

  const valores = rango.getValues();
  const fondos = rango.getBackgrounds();


  for (let r = ultimaFila - 1; r >= 0; r--) {

    for (let c = 0; c < valores[r].length; c++) {

      const fondo =
        String(fondos[r][c]).toLowerCase();


      // Busca celda amarilla
      if (
        fondo === "#ffff00" ||
        fondo === "yellow"
      ) {

        const valor =
          valores[r][c];


        if (valor instanceof Date) {
          return valor;
        }


        // Si está como texto dd/mm/yyyy
        const texto =
          String(valor).trim();


        const partes =
          texto.split("/");


        if (partes.length === 3) {

          return new Date(
            Number(partes[2]),
            Number(partes[1]) - 1,
            Number(partes[0])
          );

        }

      }

    }

  }


  return null;

}

    const fechaCreacion = new Date(fechaCelda);


    if (isNaN(fechaCreacion.getTime())) {
      Logger.log(
        "Fecha inválida en " + nombre
      );
      return;
    }


    const diferenciaDias =
      Math.floor(
        (hoy - fechaCreacion) /
        (1000 * 60 * 60 * 24)
      );


    if (diferenciaDias >= DIAS_PARA_ARCHIVAR) {


      const carpeta =
        getOrCreateSheetFolderInRoot_(nombre);


      propuestasParaArchivar.push({

        hoja: hoja,

        nombre: nombre,

        fechaCreacion: fechaCreacion,

        dias: diferenciaDias,

        carpeta: carpeta

      });

    }

  });


  if (propuestasParaArchivar.length === 0) {

    SpreadsheetApp.getActive().toast(
      "No hay propuestas para archivar",
      "Archivado",
      5
    );

    return;
  }


  Logger.log(
    "Propuestas encontradas:"
  );


  propuestasParaArchivar.forEach(p => {

    Logger.log(
      p.nombre +
      " | " +
      p.dias +
      " días"
    );

  });


  SpreadsheetApp.getActive().toast(
    propuestasParaArchivar.length +
    " propuestas listas para archivar",
    "Archivado",
    5
  );


  /*
    Aquí continuará Parte 2:

    - Crear Spreadsheet nuevo
    - Copiar hoja
    - Guardar en carpeta cliente
    - Actualizar menú
    - Eliminar hoja original
  */

Logger.log(
  "TOTAL PROPUESTAS PARA ARCHIVAR: " + propuestas.length
);

propuestas.forEach(p => {
  Logger.log(
    p.nombre + " | " + p.dias + " días"
  );
});
}


/**************** HELPERS ****************/


function getOrCreateSheetFolderInRoot_(name) {

  const folderName =
    name.replace(/[\/\\:*?"<>|]/g, " ");


  const root =
    DriveApp.getFolderById(
      OUTPUT_ROOT_FOLDER_ID
    );


  const folders =
    root.getFoldersByName(folderName);


  if (folders.hasNext()) {

    return folders.next();

  }


  return root.createFolder(folderName);

}
/***************************************************************
 * ARCHIVADO PARTE 2
 *
 * Crea respaldo independiente de una pestaña
 *
 * NO elimina la pestaña original todavía.
 *
 * Flujo:
 * 1) Crear Spreadsheet nuevo
 * 2) Guardarlo en carpeta del cliente
 * 3) Copiar pestaña completa
 * 4) Regresar links
 ***************************************************************/


function crearArchivoArchivado_(hojaOrigen) {


  const ssActual =
    SpreadsheetApp.getActiveSpreadsheet();


  const nombre =
    hojaOrigen.getName();


  // Carpeta del cliente (misma lógica actual)
  const carpeta =
    getOrCreateSheetFolderInRoot_(nombre);



  const nombreArchivo =
    nombre + " - Archivado";



  // Crear Spreadsheet nuevo
  const nuevoArchivo =
    SpreadsheetApp.create(nombreArchivo);



  const nuevoFile =
    DriveApp.getFileById(
      nuevoArchivo.getId()
    );



  // Meterlo en carpeta del cliente
  carpeta.addFile(nuevoFile);


  // Quitar de raíz de Drive
  DriveApp.getRootFolder()
    .removeFile(nuevoFile);



  // Hoja inicial creada automáticamente
  const hojaInicial =
    nuevoArchivo.getSheets()[0];



  // Copiar pestaña completa
  const copia =
    hojaOrigen.copyTo(
      nuevoArchivo
    );



  copia.setName(nombre);



  // Eliminar hoja vacía inicial
  if (
    nuevoArchivo.getSheets().length > 1
  ) {

    nuevoArchivo.deleteSheet(
      hojaInicial
    );

  }



  SpreadsheetApp.flush();



  return {

    nombre: nombre,

    archivoId:
      nuevoArchivo.getId(),

    urlSheet:
      nuevoArchivo.getUrl(),

    urlCarpeta:
      carpeta.getUrl(),

    fecha:
      new Date()

  };

}
/***************************************************************
 * ARCHIVADO PARTE 3
 *
 * Actualiza pestaña Menú con registro del archivado
 *
 * NO elimina pestañas todavía
 ***************************************************************/


function actualizarMenuArchivado_(resultado) {


  const ss =
    SpreadsheetApp.getActiveSpreadsheet();



  let menu =
    ss.getSheetByName("Menú");



  if (!menu) {

    menu =
      ss.insertSheet("Menú");

  }



  // Crear encabezados si está vacío
  if (menu.getLastRow() === 0) {


    menu.getRange(1,1,1,5)
      .setValues([[
        "Propuesta",
        "Estado",
        "Sheet Archivado",
        "Carpeta",
        "Fecha Archivado"
      ]]);


    menu.getRange(1,1,1,5)
      .setBackground("#b6d7a8")
      .setFontWeight("bold")
      .setHorizontalAlignment("center");


    menu.setFrozenRows(1);


  }



  // Buscar si ya existe
  const datos =
    menu.getRange(
      2,
      1,
      Math.max(menu.getLastRow()-1,1),
      1
    )
    .getValues();



  let fila = -1;


  for (let i=0;i<datos.length;i++){

    if (
      datos[i][0] === resultado.nombre
    ){

      fila = i + 2;
      break;

    }

  }



  // Si no existe, agregar abajo
  if (fila === -1){

    fila =
      menu.getLastRow()+1;

  }



  menu.getRange(fila,1,1,5)
    .setValues([[
      resultado.nombre,
      "Archivado",
      "",
      "",
      Utilities.formatDate(
        resultado.fecha,
        ss.getSpreadsheetTimeZone(),
        "dd/MM/yyyy"
      )
    ]]);



  // Link Sheet
  menu.getRange(fila,3)
    .setFormula(
      `=HYPERLINK("${resultado.urlSheet}","ABRIR SHEET")`
    )
    .setBackground("#007bff")
    .setFontColor("white")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");



  // Link Carpeta
  menu.getRange(fila,4)
    .setFormula(
      `=HYPERLINK("${resultado.urlCarpeta}","ABRIR CARPETA")`
    )
    .setBackground("#38761d")
    .setFontColor("white")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");



  // Ajustar columnas
  menu.setColumnWidth(1,220);
  menu.setColumnWidth(2,100);
  menu.setColumnWidth(3,140);
  menu.setColumnWidth(4,140);
  menu.setColumnWidth(5,120);


}
function testActualizarMenuArchivadoVarias(){


  const lista = [
    "EDGAR AGUILAR",
    "GONAC",
    "ALFA6659"
  ];


  const ss =
    SpreadsheetApp.getActiveSpreadsheet();



  lista.forEach(nombre => {


    const resultado = {

      nombre: nombre,

      urlSheet:
      "PRUEBA",

      urlCarpeta:
      "PRUEBA",

      fecha:new Date()

    };


    actualizarMenuArchivado_(resultado);


  });


}
/***************************************************************
 * ARCHIVADO AUTOMÁTICO PROP COM
 *
 * - Archiva propuestas >14 días
 * - Crea Sheet independiente
 * - Guarda en carpeta cliente
 * - Actualiza Menú
 * - Borra pestaña original
 ***************************************************************/


const DIAS_ARCHIVO = 14;



function EjecutarArchivadoAutomatico() {


  const ss =
    SpreadsheetApp.getActiveSpreadsheet();


  const hoy = new Date();


  const hojas =
    ss.getSheets();



  const ignorar = [
    "Menú",
    "Naves"
  ];



  let contador = 0;



  hojas.forEach(hoja => {


    const nombre =
      hoja.getName();



    if (ignorar.includes(nombre)) return;



    const fecha =
      obtenerFechaCreacion_(hoja);



    if (!fecha) {

      Logger.log(
        "Sin fecha de creación: " + nombre
      );

      return;

    }



    const dias =
      Math.floor(
        (hoy - fecha) /
        (1000*60*60*24)
      );



    if (dias <= DIAS_ARCHIVO) return;



    Logger.log(
      "Archivando: " +
      nombre +
      " | " +
      dias +
      " días"
    );



    const resultado =
      crearArchivoHistorico_(hoja, fecha);



    actualizarMenuArchivado_(resultado);



    // borrar pestaña original
    ss.deleteSheet(hoja);



    contador++;



  });



  SpreadsheetApp.getUi()
    .alert(
      "Archivado terminado",
      contador +
      " propuestas archivadas.",
      SpreadsheetApp.getUi().ButtonSet.OK
    );

}






/***************************************************************
 * BUSCAR FECHA CREACIÓN
 ***************************************************************/


function obtenerFechaCreacion_(hoja){


  // 1) Nueva ubicación AA1

  const aa =
    hoja.getRange("AA1").getValue();



  if (aa instanceof Date) {
    return aa;
  }



  // 2) Buscar fecha amarilla

  const rango =
    hoja.getDataRange();


  const valores =
    rango.getValues();


  const fondos =
    rango.getBackgrounds();



  for(let r=0;r<valores.length;r++){


    for(let c=0;c<valores[r].length;c++){


      const valor =
        valores[r][c];


      const color =
        String(fondos[r][c])
        .toLowerCase();



      if(
        valor instanceof Date &&
        color === "yellow"
      ){

        return valor;

      }


    }


  }



  return null;

}






/***************************************************************
 * CREA SHEET ARCHIVADO
 ***************************************************************/


function crearArchivoHistorico_(hoja, fecha){


  const nombre =
    hoja.getName();



  const carpeta =
    getOrCreateSheetFolderInRoot_(nombre);



  const nuevo =
    SpreadsheetApp.create(
      nombre + " - Archivado"
    );



  const archivo =
    DriveApp.getFileById(
      nuevo.getId()
    );



  carpeta.addFile(archivo);



  DriveApp.getRootFolder()
    .removeFile(archivo);



  const destino =
    SpreadsheetApp.openById(
      nuevo.getId()
    );



  const copia =
    hoja.copyTo(destino);



  copia.setName(nombre);



  // eliminar hoja vacía inicial

  const inicial =
    destino.getSheets()[0];


  if(destino.getSheets().length > 1){
    destino.deleteSheet(inicial);
  }



  return {

    nombre:nombre,

    urlSheet:
    destino.getUrl(),


    urlCarpeta:
    carpeta.getUrl(),


    fecha:new Date()

  };


}
function MostrarTodasLasPestanas() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  ss.getSheets().forEach(sheet => {

    if (sheet.isSheetHidden()) {
      sheet.showSheet();
      Logger.log("Mostrada: " + sheet.getName());
    }

  });

  SpreadsheetApp.getUi().alert(
    "Todas las pestañas ocultas fueron mostradas."
  );

}
/***************************************************************
 * ARCHIVAR PROPUESTAS MARCADAS EN MENU
 *
 * Lee columna F = OK
 * Crea Sheet independiente
 * Guarda en carpeta cliente
 * Actualiza Menú
 * Borra pestaña original
 ***************************************************************/

function ArchivarMarcadasMenu() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const menu = ss.getSheetByName("Menú");

  if (!menu) {
    throw new Error("No existe la pestaña Menú");
  }

  const ultimaFila = menu.getLastRow();

  if (ultimaFila < 2) {
    SpreadsheetApp.getUi().alert("No hay propuestas en Menú");
    return;
  }

  // Leer todos los datos actuales de la tabla (Columna A hasta F)
  const rangoDatos = menu.getRange(2, 1, ultimaFila - 1, 6);
  const datos = rangoDatos.getValues();

  let contador = 0;

  // Recorremos los datos copiados
  for (let i = 0; i < datos.length; i++) {
    const nombre = String(datos[i][0] || "").trim();
    const accion = String(datos[i][5] || "").trim().toUpperCase();

    // Solo procesa los marcados con "OK"
    if (accion !== "OK") continue;

    const hoja = ss.getSheetByName(nombre);

    if (!hoja) {
      Logger.log("No encontrada la pestaña: " + nombre);
      continue;
    }

    Logger.log("Archivando manual: " + nombre);

    // 1. Crear el Sheet independiente en Drive
    const resultado = crearArchivoHistorico_(hoja, new Date());

    // 2. BUSCAR LA FILA REAL Y ACTUAL EN EL MENÚ (Evita escribir en la fila equivocada)
    const nombresActuales = menu.getRange(2, 1, menu.getLastRow() - 1, 1).getValues();
    let filaReal = -1;

    for (let r = 0; r < nombresActuales.length; r++) {
      if (String(nombresActuales[r][0]).trim() === nombre) {
        filaReal = r + 2; // Fila exacta en el libro
        break;
      }
    }

    if (filaReal !== -1) {
      // Link Sheet archivado
      menu.getRange(filaReal, 3).setFormula(
        `=HYPERLINK("${resultado.urlSheet}","ABRIR SHEET")`
      );

      // Link carpeta cliente
      menu.getRange(filaReal, 4).setFormula(
        `=HYPERLINK("${resultado.urlCarpeta}","ABRIR CARPETA")`
      );

      // Fecha archivado
      menu.getRange(filaReal, 5).setValue(
        Utilities.formatDate(
          new Date(),
          ss.getSpreadsheetTimeZone(),
          "dd/MM/yyyy"
        )
      );

      // Cambiar OK por ARCHIVADO
      menu.getRange(filaReal, 6).setValue("ARCHIVADO");
    }

    // 3. Eliminar la pestaña original
    ss.deleteSheet(hoja);
    contador++;
  }

  SpreadsheetApp.getUi().alert(
    "Archivado terminado",
    contador + " propuestas archivadas.",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}