/**
 * Función principal que sirve la Aplicación Web (Ventana Completa).
 * Se ejecuta automáticamente cuando el usuario entra al enlace (URL) de la Web App.
 */
function doGet() {
  try {
    // Usamos createTemplateFromFile para poder "incluir" CSS y JS desde otros archivos
    return HtmlService.createTemplateFromFile("index").evaluate()
      .setTitle("Sistema de Inventario - La Nube")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch (error) {
    return HtmlService.createHtmlOutput(`<h2>Error: ${error.message}</h2>`);
  }
}

/**
 * Función para incluir el contenido de un archivo (como CSS o JS) dentro del index.html
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Obtiene la lista de marcas oficiales (Gestión de Marcas)
 */
function getMarcas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('_Marcas');
  if (!sheet) return [];
  const datos = sheet.getDataRange().getValues();
  return datos.map(row => row[0]).filter(marca => marca);
}

/**
 * Obtiene la lista de categorías oficiales
 */
function getCategorias() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('_Categorias');
  if (!sheet) return [];
  const datos = sheet.getDataRange().getValues();
  return datos.map(row => row[0]).filter(cat => cat);
}

/**
 * Guarda un nuevo producto en la hoja de Inventario, inyectando las fórmulas correspondientes.
 */
function guardarNuevoProducto(datos) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Inventario');
    
    // 1. Validaciones
    const valores = sheet.getDataRange().getValues();
    for (let i = 1; i < valores.length; i++) {
        if (datos.sku && valores[i][0] && valores[i][0].toString() === datos.sku.toString()) {
            return { success: false, error: 'El SKU ya existe en el inventario.' };
        }
        if (valores[i][2] && valores[i][2].toString().toLowerCase() === datos.nombre.toString().toLowerCase()) {
            return { success: false, error: 'Ya existe un producto con ese mismo nombre exacto.' };
        }
    }

    // 2. Insertar Valores Base (A a M)
    const rowIndex = sheet.getLastRow() + 1;
    const i = rowIndex;
    
    const costoExtra = datos.costoExtra || 0;
    const margen = datos.margen || 0; // Guardamos en decimales ej: 0.35
    
    const basicData = [[
        datos.sku || "",
        datos.marca || "",
        datos.nombre || "",
        datos.descripcion || "",
        datos.categoria || "",
        datos.stock || 0,
        datos.stockMinimo || 0,
        datos.costoUnitario || 0,
        "", // I: Formula
        costoExtra, // J
        "", // K: Formula
        margen // L
    ]];
    sheet.getRange(i, 1, 1, 12).setValues(basicData);

    // 3. Insertar Fórmulas Individuales para no chocar con las celdas editables J y L
    const IVA_RATE = 0.13;
    const IT_RATE = 0.03;
    
    if (datos.tieneFactura) {
        sheet.getRange(`I${i}`).setFormula(`=IF(H${i}<>"", H${i} * ${IVA_RATE}, "")`);
    } else {
        sheet.getRange(`I${i}`).setFormula(`=IF(H${i}<>"", 0, "")`); // No factura, cero crédito fiscal
    }
    
    sheet.getRange(`K${i}`).setFormula(`=IF(H${i}<>"", (H${i} - I${i}) + IF(J${i}="", 0, J${i}), "")`);
    sheet.getRange(`M${i}`).setFormula(`=IF(AND(K${i}<>"", L${i}<>""), K${i} / (1 - L${i}), "")`);
    sheet.getRange(`N${i}`).setFormula(`=IF(M${i}<>"", M${i} / (1 - ${IVA_RATE}), "")`);
    sheet.getRange(`O${i}`).setFormula(`=IF(N${i}<>"", N${i} * ${IT_RATE}, "")`);
    sheet.getRange(`P${i}`).setFormula(`=IF(AND(N${i}<>"", M${i}<>""), N${i} - M${i}, "")`);
    sheet.getRange(`Q${i}`).setFormula(`=IF(AND(P${i}<>"", I${i}<>""), P${i} - I${i}, "")`);
    sheet.getRange(`R${i}`).setFormula(`=IF(AND(M${i}<>"", K${i}<>""), M${i} - K${i}, "")`);
    sheet.getRange(`S${i}`).setFormula(`=IF(AND(R${i}<>"", O${i}<>""), R${i} - O${i}, "")`);
    
    // Format the new row
    sheet.getRange(i, 8, 1, 4).setNumberFormat('#,##0.00'); // H-K
    sheet.getRange(i, 13, 1, 7).setNumberFormat('#,##0.00'); // M-S
    sheet.getRange(i, 12).setNumberFormat('0.00%'); // L
    
    SpreadsheetApp.flush();
    return { success: true, message: 'Producto registrado correctamente.' };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

/**
 * Guarda una nueva marca directamente desde la Web App.
 */
function guardarNuevaMarca(marca) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('_Marcas');
    if (!sheet) return { success: false, error: "La hoja '_Marcas' no existe. Configurarla primero." };
    
    // Verificar duplicados para evitar redundancia
    const datos = sheet.getDataRange().getValues();
    const existe = datos.some(row => row && row[0] && row[0].toString().trim().toLowerCase() === marca.toLowerCase());
    if (existe) return { success: false, error: "La marca ya existe en tu base de datos." };

    sheet.appendRow([marca]);
    SpreadsheetApp.flush();
    return { success: true, message: "Marca guardada." };
  } catch(error) {
    return { success: false, error: error.message };
  }
}

/**
 * Guarda una nueva categoría directamente desde la Web App.
 */
function guardarNuevaCategoria(categoria) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('_Categorias');
    if (!sheet) {
      sheet = ss.insertSheet('_Categorias'); // Generar si fue borrado accidentalmente
      sheet.hideSheet();
    }
    
    const datos = sheet.getDataRange().getValues();
    const existe = datos.some(row => row && row[0] && row[0].toString().trim().toLowerCase() === categoria.toLowerCase());
    if (existe) return { success: false, error: "La categoría ya existe en tu base de datos." };

    sheet.appendRow([categoria]);
    SpreadsheetApp.flush();
    return { success: true, message: "Categoría guardada." };
  } catch(error) {
    return { success: false, error: error.message };
  }
}

/**
 * Sobreescribe valores bases específicos en una fila sin alterar fórmulas ajenas.
 */
function actualizarProductoBase(datos) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Inventario');
    const i = datos.fila;
    
    if (!i || i < 2) return { success: false, error: 'Identificador de fila inválido.' };

    sheet.getRange(i, 3).setValue(datos.nombre || ""); // Col C (Nombre)
    sheet.getRange(i, 4).setValue(datos.enlace || ""); // Col D (Descripción / Enlace)
    sheet.getRange(i, 6).setValue(datos.stock || 0); // Col F (Stock Actual)
    sheet.getRange(i, 8).setValue(datos.costoInicial || 0); // Col H (Costo Unitario)
    sheet.getRange(i, 10).setValue(datos.costoExtra || 0); // Col J (Costos Extras)
    sheet.getRange(i, 12).setValue(datos.margen || 0); // Col L (Margen Utilidad)
    
    SpreadsheetApp.flush();
    return { success: true, message: 'Producto actualizado correctamente.' };
  } catch (error) {
    return { success: false, error: error.message };
  }
}
