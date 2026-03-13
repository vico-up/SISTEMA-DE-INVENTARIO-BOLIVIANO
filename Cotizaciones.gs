/**
 * Módulo de Cotizaciones - Backend
 */

// Función para abrir la barra lateral
function showSidebar() {
  const html = HtmlService.createTemplateFromFile('sidebar')
    .evaluate()
    .setTitle('🛒 Nueva Cotización')
    .setWidth(400); // Ancho sugerido para sidebar
  
  SpreadsheetApp.getUi()
    .showSidebar(html);
}

// Helper para incluir archivos HTML (CSS/JS)
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
    .getContent();
}

/**
 * API: Obtiene la lista de productos del Inventario
 * Retorna: Array de objetos { sku, nombre, precio, stock }
 */
function getInventarioData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Inventario');
  
  if (!sheet) return [];
  
  const datos = sheet.getDataRange().getValues();
  // Asumimos fila 1 encabezados.
  // Col 0: SKU, 2: Nombre, 13: Precio Facturado (N)
  
  const productos = [];
  
  for (let i = 1; i < datos.length; i++) {
    const row = datos[i];
    // Solo requerimos el NOMBRE (Col 2). El SKU (Col 0) es opcional.
    if (row[2] && row[2].toString().trim() !== "") { 
      let precio = row[13]; 
      if (typeof precio !== 'number') precio = 0;
      
      productos.push({
        sku: row[0] ? row[0].toString() : "S/M", // S/M si no hay SKU
        marca: row[1] || "",
        nombre: row[2].toString(),
        precio: Math.round(precio), 
        stock: row[5] || 0,         
        tipo: 'Producto'
      });
    }
  }
  
  return productos;
}

/**
 * API: Obtiene la lista de servicios (Opcional por ahora)
 */
function getServiciosData() {
  // Servicios con precios redondeados
  return [
    { sku: 'SERV-001', nombre: 'Instalación de Cámara (Punto)', precio: 150, tipo: 'Servicio' }, // Ya son enteros
    { sku: 'SERV-002', nombre: 'Configuración DVR/NVR', precio: 250, tipo: 'Servicio' },
    { sku: 'SERV-003', nombre: 'Cableado por metro', precio: 10, tipo: 'Servicio' }
  ];
}

/**
 * Procesa la cotización: Llena la PLANTILLA MANUAL del usuario
 */
function processQuote(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetPlantilla = ss.getSheetByName('Plantilla_Cotizacion');
  
  if (!sheetPlantilla) {
    return { success: false, error: 'No se encontró la hoja "Plantilla_Cotizacion".' };
  }
  
  try {
    // 1. Datos Generales
    sheetPlantilla.getRange('E3').setValue(new Date());        // Fecha
    sheetPlantilla.getRange('B10').setValue(data.cliente || 'Consumidor Final'); // Cliente
    
    // 2. Limpiar Tabla Antigua (Desde fila 17 hasta la 40 para no tocar totales)
    sheetPlantilla.getRange('A17:E40').clearContent();
    
    // 3. Llenar Items (Desde fila 17)
    const items = data.items;
    const valores = items.map(item => [
      item.sku || '',          // A - Modelo
      item.nombre,             // B - Descripcion
      item.cantidad,           // C - Cantidad
      item.precio,             // D - Precio
      item.cantidad * item.precio // E - Subtotal Item
    ]);
    
    if (valores.length > 0) {
      // Limitamos a 24 filas (17-40) para evitar sobreescribir totales
      const numFilas = Math.min(valores.length, 24);
      sheetPlantilla.getRange(17, 1, numFilas, 5).setValues(valores.slice(0, numFilas));
    }
    
    // 4. Totales
    const subtotalGeneral = items.reduce((sum, item) => sum + (item.cantidad * item.precio), 0);
    const descuento = data.descuento || 0;
    
    sheetPlantilla.getRange('E41').setValue(subtotalGeneral); // Subtotal Suma
    sheetPlantilla.getRange('E42').setValue(descuento);       // Descuento
    sheetPlantilla.getRange('E43').setValue(subtotalGeneral - descuento); // Total
    
    SpreadsheetApp.flush();
    
    return { success: true, message: '¡Cotización generada! Revisa la hoja "Plantilla_Cotizacion".' };
    
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * Guarda un borrador de la cotización actual
 */
function saveQuoteDraft(data) {
  const userProps = PropertiesService.getUserProperties();
  userProps.setProperty('current_quote_draft', JSON.stringify(data));
  return { success: true };
}

/**
 * Carga el borrador guardado
 */
function loadQuoteDraft() {
  const userProps = PropertiesService.getUserProperties();
  const draft = userProps.getProperty('current_quote_draft');
  return draft ? JSON.parse(draft) : null;
}

/**
 * Lee los datos actuales de la hoja de plantilla para cargarlos en la barra lateral
 */
function extraerDatosDePlantilla() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Plantilla_Cotizacion');
  if (!sheet) return null;

  // Datos Cliente (B10)
  const cliente = sheet.getRange('B10').getValue();
  const descuento = sheet.getRange('E42').getValue() || 0;

  // Tabla de Items (A17:E40)
  const rows = sheet.getRange('A17:E40').getValues();
  const items = [];

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    if (row[1]) { // Si tiene descripción (Nombre)
      items.push({
        sku: row[0],
        nombre: row[1],
        cantidad: row[2] || 1,
        precio: row[3] || 0,
        tipo: row[0].toString().startsWith('SERV') ? 'Servicio' : 'Producto'
      });
    }
  }

  return {
    cliente: cliente,
    descuento: descuento,
    items: items
  };
}
