function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("💼 IEM | Gestión de Propuestas")
    
    // --- FLUJO PRINCIPAL ---
    .addItem("1. 📊 Crear Tabla de Parques", "TablaParques")
    .addSeparator()    
    .addItem("2a. 🗺️ Generar Mapa Comercial", "GenerarMapaCliente")
    .addItem("2b. 📍 Panel Lateral de Mapas", "AbrirMapaClienteSidebar")
    .addSeparator()    
    .addItem("3. 🎨 Crear Presentación (Slides)", "CrearEnSlides")
    .addSeparator()    
    .addItem("4. 📄 Generar PDF Consolidado", "GenerarPDF")   
    .addSeparator()    
    
    // --- GESTIÓN DE FICHAS Y DOCUMENTACIÓN ---
    .addItem("5a. 📋 Generar Fichas y REF en Inventario", "generarFichasPropuestasComerciales")
    .addItem("5b. 📁 Preparar Carpeta del Cliente", "prepararCliente")
    .addSeparator()
    
    // --- MÓDULO DE ENVÍO DE CORREOS ---
    .addItem("6a. ✉️ Enviar Solo Propuesta (PDF)", "EnviarPropuestaFinalPDFPorCorreo")
    .addItem("6b. ✉️ Enviar Solo Fichas Técnicas", "enviarFichasCliente")
    .addItem("6c. ✉️ Enviar Propuesta + Fichas Técnicas", "enviarPropuestaYFichasCliente")
    .addSeparator()   
    
    // --- OPCIONES INTERNACIONALES (ENG / CHI) ---
    .addItem("🌐 Tabla de Parques (ENG / CHI)", "Park_ENG_CHI") 
    .addItem("🌐 Crear Presentación (ENG / CHI)", "Slides_ENG_CHI")
    .addItem("🌐 Enviar Correo (ENG / CHI)", "SendFinalProposalPDFByEmail")
    .addSeparator()    
    
    // --- NAVEGACIÓN Y HERRAMIENTAS ---
    .addItem("🔍 Buscar NAVE (Panel Lateral)", "showNaveSidebar")
    .addItem("🏠 Ir a Menú Principal", "irAMenu")
    .addSeparator()    
    
    // --- MANTENIMIENTO Y SISTEMA ---
    .addItem("📦 Archivar Pestañas marcas OK", "ArchivarMarcadasMenu")
    .addItem("👁️ Mostrar Pestañas Ocultas", "MostrarTodasLasPestanas")
    .addToUi();
}