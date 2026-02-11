
/**
 * SISTEMA DE INVENTARIO BOLIVIANO - Versión 3.5.2 (Modular)
 * 
 * Divide la configuración en pasos seguros para evitar timeouts.
 * Incluye correción en fórmula de Utilidad Neta.
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
    
    const anchos = [100, 120, 200, 250, 120, 90, 90, 110, 110, 110, 110, 110, 110, 110, 100, 110, 110, 110, 110];
    anchos.forEach((w, i) => sheetInv.setColumnWidth(i + 1, w));
    
    SpreadsheetApp.flush();
    ui.alert('✅ PASO 1 COMPLETADO:\nEstructura creada correctamente.\n\n👉 Ahora ejecuta el Paso 2.');
    
  } catch (e) {
    ui.alert('❌ Error en Paso 1: ' + e.toString());
  }
}

/**
 * Paso 2A: Formato de Stocks (Numérico)
 */
function paso2a_Stocks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sheetInv = ss.getSheetByName('Inventario');
  
  if (!sheetInv) { ui.alert('⚠️ Error: Ejecuta el Paso 1 primero.'); return; }
  
  try {
    const range = sheetInv.getRange(2, 6, FILAS_INICIALES, 2);
    range.setNumberFormat('0');
    SpreadsheetApp.flush();
    ui.alert('✅ PASO 2A COMPLETADO: Stocks formateados.');
  } catch (e) {
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
    sheetInv.getRange(2, 8, FILAS_INICIALES, 4).setNumberFormat('#,##0.00');
    SpreadsheetApp.flush();
    sheetInv.getRange(2, 13, FILAS_INICIALES, 7).setNumberFormat('#,##0.00');
    SpreadsheetApp.flush();
    ui.alert('✅ PASO 2B COMPLETADO: Moneda formateada.');
  } catch (e) {
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
    sheetInv.getRange(2, 12, FILAS_INICIALES, 1).setNumberFormat('0.00%');
    SpreadsheetApp.flush();
    ui.alert('✅ PASO 2C COMPLETADO: Porcentajes formateados.');
  } catch (e) {
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
    sheetMarcas.showSheet();
    if (sheetMarcas.getLastRow() < 1) {
      const marcas = [['Genérico'], ['Sony'], ['Samsung'], ['TP-Link'], ['Xiaomi']];
      sheetMarcas.getRange(1, 1, marcas.length, 1).setValues(marcas);
    }
    SpreadsheetApp.flush();
    
    const rangeMarcas = sheetMarcas.getRange('A1:A100');
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(rangeMarcas, true)
      .setHelpText('Selecciona una marca de la lista.')
      .build();
    
    sheetInv.getRange(2, 2, FILAS_INICIALES).setDataValidation(rule);
    sheetMarcas.hideSheet();
    SpreadsheetApp.flush();
    ui.alert('✅ PASO 2D COMPLETADO: Lista desplegable de marcas activada.');
  } catch (e) {
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
  
  if (!sheetInv) { ui.alert('⚠️ Error: Ejecuta el Paso 1 primero.'); return; }
  
  try {
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
    ui.alert('✅ PASO 3 COMPLETADO: Sistema listo para usar.');
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
    .addItem('➕ Agregar Nueva Marca', 'agregarNuevaMarca')
    .addToUi();
}

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
