function irAMenu() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaMenu = ss.getSheetByName("Menú");

  if (!hojaMenu) {
    SpreadsheetApp.getUi().alert('No existe la pestaña "Menú"');
    return;
  }

  ss.setActiveSheet(hojaMenu);

  // Ir al último renglón con datos
  const ultimaFila = hojaMenu.getLastRow();

  if (ultimaFila > 0) {
    hojaMenu.setActiveRange(
      hojaMenu.getRange(ultimaFila, 1)
    );
  }
}