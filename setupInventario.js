
/**
 * SISTEMA DE INVENTARIO BOLIVIANO - Versión 3.5 (Modular)
 * 
 * Divide la configuración en 3 pasos seguros para evitar timeouts.
 */

const IVA_RATE = 0.13;
const IT_RATE = 0.03;
const FILAS_INICIALES = 20;

// ---------------------------------------------------------
// FUNCIONES DEL MENÚ
// ---------------------------------------------------------

/**
 * Paso 1: Crea las hojas y encabezados
 */
function paso1_Estructura() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  try {
    // 1. Gestión de Hojas
    let sheetInv = ss.getSheetByName('Inventario');
    if (sheetInv) {
      sheetInv.clear();
      sheetInv.setName('Inventario_OLD');
    }
    sheetInv = ss.insertSheet('Inventario');
    if (ss.getSheetByName('Inventario_OLD')) {
      ss.deleteSheet(ss.getSheetByName('Inventario_OLD'));
    }
    
    let sheetMarcas = ss.getSheetByName('_Marcas');
    if (sheetMarcas) {
      sheetMarcas.clear();
      sheetMarcas.setName('_Marcas_OLD');
    }
    sheetMarcas = ss.insertSheet('_Marcas');
    if (ss.getSheetByName('_Marcas_OLD')) {
      ss.deleteSheet(ss.getSheetByName('_Marcas_OLD'));
    }
    sheetMarcas.hideSheet();
    
    // 2. Encabezados
    const headers = [
      'SKU', 'Marca', 'Nombre Producto', 'Descripción', 'Categoría', 
      'Stock Actual', 'Stock Mínimo', 'Costo Unitario', 'Crédito Fiscal (13%)', 
      'Costos Extras', 'Precio Costo (Pc)', 'Margen Utilidad', 'Precio Venta (Pv)', 
      'Precio Facturado (PF)', 'IT (3%)', 'IVA Débito (13%)', 'IVA a Pagar', 
      'Utilidad Bruta', 'Utilidad Neta'
    ];
    
    sheetInv.getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setFontWeight('bold')
      .setBackground('#4a86e8')
      .setFontColor('white')
      .setHorizontalAlignment('center');
      
    sheetInv.setFrozenRows(1);
    
    // 3. Anchos
    const anchos = [100, 120, 200, 250, 120, 90, 90, 110, 110, 110, 110, 110, 110, 110, 100, 110, 110, 110, 110];
    anchos.forEach((w, i) => sheetInv.setColumnWidth(i + 1, w));
    
    SpreadsheetApp.flush();
    ui.alert('✅ PASO 1 COMPLETADO:\nEstructura creada correctamente.\n\n👉 Ahora ejecuta el Paso 2.');
    
  } catch (e) {
    ui.alert('❌ Error en Paso 1: ' + e.toString());
  }
}

/**
 * Paso 2: Aplica formatos y validaciones
 */
/**
 * Paso 2: Aplica formatos y validaciones
 */
/**
 * Paso 2A: Formato de Stocks (Numérico)
 */
function paso2a_Stocks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sheetInv = ss.getSheetByName('Inventario');
  
  if (!sheetInv) { ui.alert('⚠️ Error: Ejecuta el Paso 1 primero.'); return; }
  
  try {
    console.log('--- INICIO PASO 2A (STOCKS) ---');
    // 1. Formato Stocks
    const range = sheetInv.getRange(2, 6, FILAS_INICIALES, 2);
    range.setNumberFormat('0');
    SpreadsheetApp.flush();
    
    console.log('--- PASO 2A COMPLETADO ---');
    ui.alert('✅ PASO 2A COMPLETADO: Stocks formateados.');
    
  } catch (e) {
    console.error('ERROR EN PASO 2A: ' + e.toString());
    ui.alert('❌ Error en Paso 2A: ' + e.toString());
  }
}

/**
 * Paso 2B: Formato de Moneda (Bs)
 */
function paso2b_Moneda() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sheetInv = ss.getSheetByName('Inventario');
  
  if (!sheetInv) { ui.alert('⚠️ Error: Ejecuta el Paso 1 primero.'); return; }
  
  try {
    console.log('--- INICIO PASO 2B (MONEDA) ---');
    // 2. Formato Costos (Simplificado para evitar error de locale)
    sheetInv.getRange(2, 8, FILAS_INICIALES, 4).setNumberFormat('#,##0.00'); // H-K
    SpreadsheetApp.flush();
    
    // 3. Formato Precios (Simplificado)
    sheetInv.getRange(2, 13, FILAS_INICIALES, 7).setNumberFormat('#,##0.00'); // M-S
    SpreadsheetApp.flush();
    
    console.log('--- PASO 2B COMPLETADO ---');
    ui.alert('✅ PASO 2B COMPLETADO: Moneda formateada (Formato estándar para estabilidad).');
    
  } catch (e) {
    console.error('ERROR EN PASO 2B: ' + e.toString());
    ui.alert('❌ Error en Paso 2B: ' + e.toString());
  }
}

/**
 * Paso 2C: Formato de Porcentajes
 */
function paso2c_Porcentaje() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sheetInv = ss.getSheetByName('Inventario');
  
  if (!sheetInv) { ui.alert('⚠️ Error: Ejecuta el Paso 1 primero.'); return; }
  
  try {
    console.log('--- INICIO PASO 2C (PORCENTAJE) ---');
    // 4. Formato Margen
    sheetInv.getRange(2, 12, FILAS_INICIALES, 1).setNumberFormat('0.00%');
    SpreadsheetApp.flush();
    
    console.log('--- PASO 2C COMPLETADO ---');
    ui.alert('✅ PASO 2C COMPLETADO: Porcentajes formateados.\n\n👉 Ahora ve al Paso 3.');
    
  } catch (e) {
    console.error('ERROR EN PASO 2C: ' + e.toString());
    ui.alert('❌ Error en Paso 2C: ' + e.toString());
  }
}

/**
 * Paso 2D: Validación de Marcas
 */
function paso2d_Validacion() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sheetInv = ss.getSheetByName('Inventario');
  const sheetMarcas = ss.getSheetByName('_Marcas');
  
  if (!sheetInv || !sheetMarcas) { ui.alert('⚠️ Error: Ejecuta el Paso 1 primero.'); return; }
  
  try {
    console.log('--- INICIO PASO 2D (VALIDACIÓN) ---');
    
    // 1. Asegurar marcas iniciales
    sheetMarcas.showSheet();
    if (sheetMarcas.getLastRow() < 1) {
      const marcas = [['Genérico'], ['Sony'], ['Samsung'], ['TP-Link'], ['Xiaomi']];
      sheetMarcas.getRange(1, 1, marcas.length, 1).setValues(marcas);
    }
    SpreadsheetApp.flush();
    
    // 2. Crear y aplicar regla
    const rangeMarcas = sheetMarcas.getRange('A1:A100');
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(rangeMarcas, true)
      .setHelpText('Selecciona una marca de la lista.')
      .build();
    
    sheetInv.getRange(2, 2, FILAS_INICIALES).setDataValidation(rule);
    sheetMarcas.hideSheet();
    SpreadsheetApp.flush();
    
    console.log('--- PASO 2D COMPLETADO ---');
    ui.alert('✅ PASO 2D COMPLETADO: Lista desplegable de marcas activada.');
    
  } catch (e) {
    console.error('ERROR EN PASO 2D: ' + e.toString());
    ui.alert('❌ Error en Paso 2D: ' + e.toString());
  }
}

/**
 * Paso 3: Agrega fórmulas y datos de ejemplo
 */
function paso3_Datos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sheetInv = ss.getSheetByName('Inventario');
  
  if (!sheetInv) {
    ui.alert('⚠️ Error: Ejecuta el Paso 1 primero.');
    return;
  }
  
  try {
    // 1. Fórmulas
    const formulas = [];
    for (let i = 2; i <= FILAS_INICIALES + 1; i++) {
      formulas.push([
        `=IF(H${i}<>"", H${i} * ${IVA_RATE}, "")`,
        '',
        `=IF(H${i}<>"", (H${i} - I${i}) + IF(J${i}="", 0, J${i}), "")`,
        '',
        `=IF(AND(K${i}<>"", L${i}<>""), K${i} / (1 - L${i}), "")`,
        `=IF(M${i}<>"", M${i} / (1 - ${IVA_RATE}), "")`,
        `=IF(N${i}<>"", N${i} * ${IT_RATE}, "")`,
        `=IF(AND(N${i}<>"", M${i}<>""), N${i} - M${i}, "")`,
        `=IF(AND(P${i}<>"", I${i}<>""), P${i} - I${i}, "")`,
        `=IF(AND(M${i}<>"", K${i}<>""), M${i} - K${i}, "")`,
        `=IF(AND(R${i}<>"", O${i}<>""), R${i} - O${i}, "")`
      ]);
    }
    sheetInv.getRange(2, 9, formulas.length, 11).setFormulas(formulas);
    SpreadsheetApp.flush();
    
    // 2. Datos de Ejemplo
    const ejemplo = [
      ['CAM001', 'Sony', 'Sony A7 IV', 'Cámara Mirrorless', 'Fotografía', 5, 2, 10000, '', 500, '', 0.35],
      ['NET042', 'TP-Link', 'Archer AX50', 'Router Wi-Fi 6', 'Redes', 15, 5, 450, '', 0, '', 0.40]
    ];
    
    ejemplo.forEach((fila, idx) => {
      sheetInv.getRange(idx + 2, 1, 1, 8).setValues([fila.slice(0, 8)]);
      sheetInv.getRange(idx + 2, 10).setValue(fila[9]);
      sheetInv.getRange(idx + 2, 12).setValue(fila[11]);
    });
    
    SpreadsheetApp.flush();
    sheetInv.activate();
    ui.alert('✅ PASO 3 COMPLETADO:\nSistema listo para usar.\n\n¡Felicidades! 🎉');
    
  } catch (e) {
    ui.alert('❌ Error en Paso 3: ' + e.toString());
  }
}

// ---------------------------------------------------------
// MENÚ PRINCIPAL
// ---------------------------------------------------------

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📦 Inventario')
    .addItem('1️⃣ Paso 1: Estructura', 'paso1_Estructura')
    .addSeparator()
    .addItem('2️⃣-A Paso 2A: Stocks', 'paso2a_Stocks')
    .addItem('2️⃣-B Paso 2B: Moneda', 'paso2b_Moneda')
    .addItem('2️⃣-C Paso 2C: Porcentajes', 'paso2c_Porcentaje')
    .addItem('2️⃣-D Paso 2D: Validación', 'paso2d_Validacion')
    .addSeparator()
    .addItem('3️⃣ Paso 3: Datos y Fórmulas', 'paso3_Datos')
    .addSeparator()
    .addItem('📑 Nueva Cotización', 'showSidebar')
    .addItem('🛠️ Configurar Plantilla', 'setupPlantillaCotizacion')
    .addSeparator()
    .addItem('➕ Agregar Nueva Marca', 'agregarNuevaMarca')
    .addToUi();
}

/**
 * Agrega una marca a la lista de validación
 */
function agregarNuevaMarca() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt('➕ Nueva Marca', 'Introduce el nombre de la marca:', ui.ButtonSet.OK_CANCEL);
  
  if (resp.getSelectedButton() == ui.Button.OK) {
    const marca = resp.getResponseText().trim();
    if (marca) {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheetMarcas = ss.getSheetByName('_Marcas');
      if (sheetMarcas) {
        sheetMarcas.appendRow([marca]);
        ui.alert('✅ Marca "' + marca + '" agregada.');
      } else {
        ui.alert('⚠️ Error: No se encontró la hoja de marcas. Ejecuta la configuración primero.');
      }
    }
  }
}

/**
 * Crea la hoja de Plantilla para Cotizaciones (Diseño Profesional)
 */
function setupPlantillaCotizacion() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  try {
    let sheet = ss.getSheetByName('Plantilla_Cotizacion');
    if (sheet) { ss.deleteSheet(sheet); }
    sheet = ss.insertSheet('Plantilla_Cotizacion');
    
    // Configuración General
    sheet.setHiddenGridlines(true);
    sheet.setColumnWidth(1, 20); // Margen Izquierdo (A)
    sheet.setColumnWidth(2, 100); // Modelo (B)
    sheet.setColumnWidth(3, 300); // Descripción (C)
    sheet.setColumnWidth(4, 60);  // Cantidad (D)
    sheet.setColumnWidth(5, 100); // Precio (E)
    sheet.setColumnWidth(6, 100); // Subtotal (F)
    sheet.setColumnWidth(7, 20);  // Margen Derecho (G)

    // 1. ENCABEZADO (LA NUBE SERVICIOS)
    sheet.getRange('B2').setValue('LA NUBE SERVICIOS').setFontSize(18).setFontWeight('bold').setFontColor('#4a86e8');
    sheet.getRange('F2').setValue('COTIZACION').setFontSize(18).setFontWeight('bold').setHorizontalAlignment('right');
    
    const infoEmpresa = [
      ['Av. Blanco Galindo Km. 11'],
      ['Cochabamba - Bolivia'],
      ['www.lanube.com'],
      ['Teléfono: 62722006'],
      ['Email: victormartinez.vico@gmail.com'],
      ['Asesor de venta: Victor Gareca']
    ];
    sheet.getRange(3, 2, infoEmpresa.length, 1).setValues(infoEmpresa).setFontSize(9).setFontColor('#5f6368');

    // Caja de Metadatos (Derecha)
    const metaLabels = [['FECHA'], ['COTIZACION #'], ['CLIENTE ID'], ['VALIDO HASTA']];
    sheet.getRange('E3:E6').setValues(metaLabels).setFontSize(9).setFontWeight('bold').setBorder(true, true, true, true, null, null);
    sheet.getRange('F3:F6').setBorder(true, true, true, true, null, null); // Espacio para valores
    
    // 2. SECCIÓN CLIENTE
    sheet.getRange('B10:F10').merge().setValue('CLIENTE').setFontWeight('bold').setFontColor('white').setBackground('black');
    const labelsCliente = [['Nombre:'], ['Dirección:'], ['Ciudad:'], ['Teléfono:'], ['Email:']];
    sheet.getRange(11, 2, 5, 1).setValues(labelsCliente).setFontSize(10).setFontWeight('bold');
    sheet.getRange('B11:F15').setBorder(null, true, true, true, null, null);

    // 3. TABLA DE PRODUCTOS
    const headerRow = 17;
    const headers = ['MODELO', 'DESCRIPCION EQUIPO', 'CANT.', 'PRECIO', 'SUB-TOTAL'];
    sheet.getRange(headerRow, 2, 1, 5).setValues([headers])
      .setFontWeight('bold').setFontColor('white').setBackground('black').setHorizontalAlignment('center');
    
    // Bordes de la tabla (Filas 18 a 35)
    sheet.getRange(18, 2, 18, 5).setBorder(true, true, true, true, true, null);

    // 4. TOTALES
    const totalRow = 36;
    sheet.getRange(totalRow, 5).setValue('SUBTOTAL').setFontWeight('bold').setBorder(true, true, true, true, null, null);
    sheet.getRange(totalRow+1, 5).setValue('DESCUENTO').setFontWeight('bold').setBorder(true, true, true, true, null, null);
    sheet.getRange(totalRow+2, 5).setValue('TOTAL').setFontWeight('bold').setFontColor('white').setBackground('#4a86e8').setBorder(true, true, true, true, null, null);
    
    sheet.getRange(totalRow, 6, 3, 1).setBorder(true, true, true, true, null, null);

    // 5. TERMINOS Y CONDICIONES
    const terminosRow = 40;
    sheet.getRange(terminosRow, 2, 1, 4).merge().setValue('TERMINOS Y CONDICIONES').setFontWeight('bold').setFontColor('white').setBackground('black');
    const terminos = [
      ['1. El pago será cancelado el 60% previa instalacion y el 40% finalizado la instalacion'],
      ['2. La garantía de los equipos e instalacion es de 1 año'],
      ['3. Los materiales empleados se ajustará de acuerdo al montaje de la instalacion'],
      ['4. Plazo de entrega 2 - 3 dias habiles segun stock de productos ofrecidos'],
      ['5. Los items con * solo aplica si son necesarios por factibilidad tecnica']
    ];
    sheet.getRange(terminosRow + 1, 2, 5, 1).setValues(terminos).setFontSize(9);
    sheet.getRange(terminosRow, 2, 10, 4).setBorder(true, true, true, true, null, null);

    // Firma
    sheet.getRange(terminosRow + 7, 2).setValue('x________________________').setFontSize(10);
    sheet.getRange(terminosRow + 8, 2).setValue('Nombre Cliente:').setFontSize(9);

    // Footer Tagline
    sheet.getRange(52, 2, 1, 5).merge().setValue('Gracias por trabajar con nosotros!').setFontWeight('bold').setHorizontalAlignment('center');
    
    ui.alert('✅ Plantilla "LA NUBE SERVICIOS" creada correctamente.\n\nRecuerda insertar tu logo manualmente en la parte superior izquierda.');
    
  } catch (e) {
    ui.alert('❌ Error creando plantilla: ' + e.toString());
  }
}
