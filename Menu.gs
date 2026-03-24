function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Tabla, Mapa, Propuesta")
    .addItem("📁 Generar Carpeta de Fichas PDF", "generarFichas")
    .addItem("1. Crear Tabla Final", "CrearTablaFinal")
    .addItem("1b. Crear Tabla de Parques", "TablaParques")
    .addSeparator() // --- Línea separadora ---    
    .addItem("2a. Generar Mapa Cliente", "GenerarMapaCliente")
    .addItem("2b. Abrir mapa (panel lateral)", "AbrirMapaClienteSidebar")
    .addSeparator() // --- Línea separadora ---    
    .addItem("3. Crear Slide", "CrearEnSlides")
    .addSeparator() // --- Línea separadora ---    
    .addItem("4. Generar PDF final", "GenerarPDF")   
    .addSeparator() // --- Línea separadora ---    
    .addItem("5. Enviar PDF por correo", "EnviarPropuestaFinalPDFPorCorreo")
    .addSeparator()
    .addItem("Tabla Parques ENG/CHI", "Park_ENG_CHI")
    .addItem("Tabla Final ENG/CHI", "CrearTablaFinal_ENG_CHI")
    .addItem("Crear Slide ENG/CHI", "Slides_ENG_CHI")
    .addSeparator() // --- Línea separadora ---    
    .addItem('Sidebar NAVE', 'showNaveSidebar')
    .addToUi();
}