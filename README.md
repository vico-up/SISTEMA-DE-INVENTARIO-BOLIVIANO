# Sistema de Inventario y Cotizaciones - Bolivia 🇧🇴

Sistema profesional de gestión de inventarios y cálculo de impuestos (IVA/IT) optimizado para Google Sheets.

## 🚀 Estado Actual: **v1.0.0 Estable**
El sistema ha sido rediseñado para garantizar estabilidad total y evitar errores de conexión ("Service Spreadsheets failed").

---

## 🛠️ Instalación y Uso

### 1. Preparación en Apps Script
1. Abre tu Google Sheet.
2. Ve a **Extensiones > Apps Script**.
3. Copia el contenido del archivo `setupInventario.gs` de este repositorio.
4. Pégalo en el editor de Apps Script (borrando cualquier código previo).
5. Guarda (`Ctrl+S`) y actualiza tu hoja de cálculo (`F5`).

### 2. Configuración en 3 Pasos Seguros
Para evitar bloqueos de Google, usa el nuevo menú modular **📦 Inventario** en este orden:
1. **1️⃣ Paso 1: Estructura** (Crea las hojas y encabezados).
2. **2️⃣-A / B / C Paso 2** (Aplica formatos y validaciones uno por uno).
3. **3️⃣ Paso 3: Datos y Fórmulas** (Inserta las fórmulas automáticas y ejemplos).

---

## 📁 Estructura del Proyecto

- `setupInventario.gs`: El archivo "vivo". Cualquier cambio aquí se refleja en tu sistema.
- `backups/`: Copias de seguridad de versiones estables. Úsalas si algo falla.
- `README.md`: Esta guía.
- `.gitignore`: Archivos excluidos del repositorio.

---

## 🔄 Workflow con GitHub

Para guardar tus avances y mantener tu código seguro:

1. Modifica `setupInventario.gs`.
2. Ejecuta en tu terminal:
   ```bash
   git add .
   git commit -m "Descripción de tu mejora"
   git push
   ```

---

## 🧪 Próximas Funcionalidades
- [ ] Módulo de Cotizaciones Automáticas (v1.1.0).
- [ ] Exportación a PDF con logo.
- [ ] Panel lateral de búsqueda rápida.

---

## 🆘 Soporte y Documentación
Consulta el archivo [`implementation_plan.md`](implementation_plan.md) para detalles técnicos sobre el sistema tributario boliviano.
