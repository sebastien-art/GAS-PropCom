function actualizarFechasCreacionMenu() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const menu = ss.getSheetByName("Menú");

  const ultimaFila = menu.getLastRow();

  // A:G
  const datos = menu.getRange(2, 1, ultimaFila - 1, 7).getValues();
  const formulas = menu.getRange(2, 3, ultimaFila - 1, 1).getFormulas();

  let actualizadas = 0;
  let errores = [];

  for (let i = 0; i < datos.length; i++) {

    const nombreHoja = datos[i][0];
    if (!nombreHoja) continue;

    let hoja = ss.getSheetByName(nombreHoja);

    // -------------------------------------------------
    // SI NO EXISTE EN ESTE ARCHIVO, BUSCAR EL ARCHIVADO
    // -------------------------------------------------
    if (!hoja) {

      const formula = formulas[i][0];

      if (formula) {

        const match = formula.match(/https:\/\/docs\.google\.com\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);

        if (match) {

          try {

            const archivoArchivado = SpreadsheetApp.openById(match[1]);

            hoja = archivoArchivado.getSheetByName(nombreHoja);

          } catch (e) {

            errores.push(nombreHoja + " (no pudo abrir archivo)");

            continue;

          }

        }

      }

    }

    if (!hoja) {

      errores.push(nombreHoja + " (hoja no encontrada)");

      continue;

    }

    //--------------------------------------------------
    // BUSCAR FECHA
    //--------------------------------------------------

    let fecha = null;

    // 1. Intentar AA1

    const valorAA1 = hoja.getRange("AA1").getValue();

    if (valorAA1 instanceof Date) {

      fecha = valorAA1;

    } else {

      // 2. Buscar celda amarilla

      const rango = hoja.getDataRange();

      const valores = rango.getValues();

      const fondos = rango.getBackgrounds();

      buscar:

      for (let r = 0; r < valores.length; r++) {

        for (let c = 0; c < valores[r].length; c++) {

          const color = fondos[r][c].toLowerCase();

          if (
            (color == "#ffff00" ||
             color == "#ff0" ||
             color == "yellow") &&
             valores[r][c] instanceof Date
          ) {

            fecha = valores[r][c];

            break buscar;

          }

        }

      }

    }

    //--------------------------------------------------
    // ESCRIBIR EN MENÚ (Columna G)
    //--------------------------------------------------

    if (fecha) {

      menu.getRange(i + 2, 7).setValue(fecha);

      actualizadas++;

    }

  }

  SpreadsheetApp.flush();

  SpreadsheetApp.getUi().alert(
    "Proceso terminado.\n\n" +
    "Fechas actualizadas: " + actualizadas +
    (errores.length
      ? "\n\nNo encontradas:\n" + errores.join("\n")
      : "")
  );

}