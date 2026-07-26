function validarHipervinculosColumna() {

  const ui = SpreadsheetApp.getUi();

  const respuesta = ui.prompt(
    "Validar hipervínculos",
    "¿Qué columna deseas revisar? (Ejemplo: B, C, D)",
    ui.ButtonSet.OK_CANCEL
  );

  if (respuesta.getSelectedButton() != ui.Button.OK) return;

  const letra = respuesta.getResponseText().trim().toUpperCase();

  if (!/^[A-Z]+$/.test(letra)) {
    ui.alert("Columna no válida.");
    return;
  }

  const hoja = SpreadsheetApp.getActiveSheet();

  // Convierte letra a número de columna
  const columna = letra
    .split("")
    .reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0);

  const ultimaFila = hoja.getLastRow();
  const formulas = hoja.getRange(2, columna, ultimaFila - 1, 1).getFormulas();

  let buenos = 0;
  let malos = 0;

  // Lista de GIDs existentes en este Spreadsheet
  const gids = {};
  SpreadsheetApp.getActive().getSheets().forEach(s => {
    gids[s.getSheetId()] = true;
  });

  for (let i = 0; i < formulas.length; i++) {

    const formula = formulas[i][0];
    if (!formula) continue;

    let valido = true;

    // -----------------------------------
    // HIPERVÍNCULO INTERNO (#gid=...)
    // -----------------------------------
    const gid = formula.match(/#gid=(\d+)/);

    if (gid) {

      if (!gids[gid[1]]) {
        valido = false;
      }

    } else {

      // -----------------------------------
      // HIPERVÍNCULO A OTRO SPREADSHEET
      // -----------------------------------
      const url = formula.match(/https:\/\/docs\.google\.com\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);

      if (url) {

        try {

          SpreadsheetApp.openById(url[1]);

        } catch (e) {

          valido = false;

        }

      }

    }

    // Solo hacer algo si es INVÁLIDO
    if (!valido) {
      hoja.getRange(i + 2, columna).setBackground("#ff9999");
      malos++;
    } else {
      buenos++;
    }

  }

  ui.alert(
    "Validación terminada\n\n" +
    "Enlaces válidos: " + buenos +
    "\nEnlaces inválidos: " + malos
  );

}