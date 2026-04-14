/**
 * Se ejecuta automáticamente al editar cualquier celda
 */
function onEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() !== 'Inventario') return;
  
  const col = e.range.getColumn();
  const row = e.range.getRow();
  
  // Solo nos importa si edita SKU (1) o Nombre (3), a partir de la fila 2
  if (row <= 1 || (col !== 1 && col !== 3)) return;
  
  const val = e.value;
  if (!val) return; // Si borró la celda, no validamos duplicados
  
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;
  
  const data = sheet.getRange(2, col, lastRow - 1, 1).getValues();
  let count = 0;
  
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().toLowerCase().trim() === val.toString().toLowerCase().trim()) {
      count++;
    }
  }
  
  // count será al menos 1 porque el valor ya está escrito en la hoja
  if (count > 1) { 
    const ui = SpreadsheetApp.getUi();
    const columnName = col === 1 ? 'SKU' : 'Nombre del Producto';
    
    // Muestra alerta
    ui.alert('⚠️ ALERTA DE DUPLICADO', 
      'El ' + columnName + ' "' + val + '" ya existe en otra fila del inventario.\n\nPor favor, ingresa un valor diferente para evitar confusiones al cotizar o buscar productos.', 
      ui.ButtonSet.OK);
      
    // Limpia la celda para obligar a introducir uno nuevo
    e.range.clearContent();
  }
}
