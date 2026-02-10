# Sistema de Inventario y Cotizaciones - Google Sheets

Sistema completo de gestión de inventario, cálculo de precios y generación de cotizaciones utilizando Google Sheets y Apps Script.

## 🚀 Instalación Inicial

### Paso 1: Crear el Google Sheets
1. Ve a [Google Sheets](https://sheets.google.com)
2. Crea una nueva hoja de cálculo
3. Nómbrala como "Sistema de Inventario" (o el nombre que prefieras)

### Paso 2: Abrir el Editor de Apps Script
1. En tu Google Sheet, ve al menú: **Extensiones → Apps Script**
2. Se abrirá el editor de código en una nueva pestaña

### Paso 3: Pegar el Código de Configuración
1. Borra el código de ejemplo que aparece (`function myFunction() {}`)
2. Abre el archivo `setupInventario.gs` desde tu carpeta local
3. Copia TODO el contenido del archivo
4. Pégalo en el editor de Apps Script
5. Haz clic en el **💾 icono de guardar** (o presiona `Ctrl + S`)
6. Dale un nombre al proyecto, por ejemplo: "Inventario Setup"

### Paso 4: Ejecutar la Configuración
1. En el editor de Apps Script, en el menú desplegable de funciones (arriba), selecciona **`setupInventarioSheet`**
2. Haz clic en el botón **▶️ Ejecutar**
3. La primera vez te pedirá permisos:
   - Haz clic en **"Revisar permisos"**
   - Selecciona tu cuenta de Google
   - Haz clic en **"Avanzado"** (abajo a la izquierda)
   - Haz clic en **"Ir a [nombre del proyecto] (no seguro)"**
   - Haz clic en **"Permitir"**
4. El script te preguntará si deseas agregar datos de ejemplo → selecciona **Sí** para ver cómo funciona
5. Verás un mensaje de **"✅ Configuración completa"**

### Paso 5: Volver a tu Google Sheet
1. Regresa a la pestaña de tu Google Sheet
2. Deberías ver:
   - Una hoja llamada **"Inventario"** con todas las columnas configuradas
   - Un nuevo menú en la barra superior: **📦 Inventario**
   - Si elegiste agregar ejemplos, verás 3 productos de demostración

---

## 📋 Estructura de la Hoja "Inventario"

| Columna | Descripción | Tipo | Validación |
|---------|-------------|------|------------|
| **SKU** | Código único del producto | Texto | Manual |
| **Marca** | Marca del producto | Lista desplegable | Solo valores de `_Marcas` |
| **Nombre del Producto** | Nombre descriptivo | Texto | Manual |
| **Descripción** | Detalles del producto | Texto | Manual |
| **Categoría** | Clasificación del producto | Texto | Manual |
| **Stock Actual** | Cantidad disponible | Número entero | Manual |
| **Stock Mínimo** | Alerta de reabastecimiento | Número entero | Manual |
| **Costo Unitario** | Precio de compra | Moneda ($) | Manual |
| **Margen (%)** | Porcentaje de ganancia | Porcentaje | Manual |
| **Precio de Venta** | Calculado automáticamente | Moneda ($) | **Automático** |

---

## ✨ Funcionalidades Implementadas

### 1. ✅ Validación de Marcas
- La columna **Marca** solo permite seleccionar de una lista predefinida
- Evita errores de tipeo
- Las marcas se gestionan desde la hoja oculta `_Marcas`

### 2. 🧮 Cálculo Automático de Precios
- El **Precio de Venta** se calcula automáticamente con la fórmula:
  ```
  Precio de Venta = Costo Unitario × (1 + Margen %)
  ```
- Ejemplo: Si costo = $100 y margen = 25% → Precio de venta = $125

### 3. 🚨 Alertas de Stock Bajo
- Si el **Stock Actual** es menor o igual al **Stock Mínimo**, la celda se pinta en rojo
- Te ayuda a identificar visualmente qué productos necesitas reabastecer

### 4. 📦 Menú Personalizado
Después de ejecutar el script, verás un menú nuevo llamado **"📦 Inventario"** con opciones para:
- **⚙️ Configurar Hoja Inventario**: Vuelve a ejecutar la configuración inicial
- **➕ Agregar Nueva Marca**: Agrega marcas nuevas sin editar la hoja oculta

---

## 🏷️ Cómo Agregar Nuevas Marcas

### Método 1: Usando el Menú (Recomendado)
1. Ve al menú **📦 Inventario → ➕ Agregar Nueva Marca**
2. Escribe el nombre de la marca en el cuadro de diálogo
3. Haz clic en **OK**
4. La marca se agregará automáticamente y estará disponible en la lista desplegable

### Método 2: Editando Manualmente
1. Ve a la hoja **_Marcas** (está oculta por defecto)
   - Para mostrarla: clic derecho en las pestañas → **Mostrar → _Marcas**
2. Agrega la nueva marca en una celda vacía de la columna A
3. Vuelve a ocultar la hoja si lo deseas

---

## 📝 Próximos Pasos

La hoja de **Inventario** ya está lista para usarse. Los siguientes módulos que implementaremos serán:

- [ ] Hoja **Cotizaciones** (histórico de cotizaciones)
- [ ] Hoja **Plantilla_Cotización** (diseño visual para PDF)
- [ ] Panel lateral para búsqueda rápida de productos
- [ ] Generación automática de cotizaciones en PDF
- [ ] Descuento automático de inventario al cerrar una venta

---

## 🆘 Solución de Problemas

### El menú "📦 Inventario" no aparece
- Cierra y vuelve a abrir la hoja de Google Sheets
- El menú se genera automáticamente cuando abres el archivo gracias a la función `onOpen()`

### No puedo seleccionar marcas en la columna "Marca"
- Asegúrate de que la hoja `_Marcas` existe y tiene al menos una marca
- Ejecuta nuevamente la función `setupInventarioSheet` desde el menú

### El precio de venta no se calcula automáticamente
- Verifica que hayas ingresado valores numéricos en **Costo Unitario** y **Margen (%)**
- La columna **Precio de Venta** tiene una fórmula en la fila 2, que se debe copiar hacia abajo

---

## 📧 Contacto

Para más información o soporte, consulta la [documentación completa](implementation_plan.md).
