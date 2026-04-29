# Tool Allocation Dashboard — Intel

Dashboard interactivo para visualizar y localizar productos en máquinas de testeo de la planta de producción.

---

## Requisitos previos

- **Windows 7 o superior** (o Linux/macOS con Node.js)
- **Node.js 16+** con npm (descargable desde [nodejs.org](https://nodejs.org))
- **Acceso a la red interna de Intel** para conectarse a la base de datos

---

## Instalación

### 1. Instalar Node.js

1. Ve a [https://nodejs.org](https://nodejs.org)
2. Descarga la versión **LTS** (Long Term Support)
3. Ejecuta el instalador y sigue los pasos por defecto
4. Abre una terminal nueva y verifica:
   ```
   node --version
   npm --version
   ```

### 2. Ejecutar el Dashboard

**Opción A: Doble-clic (más fácil)**
- Navega a la carpeta del proyecto
- Haz doble-clic en **`start.bat`**
- Se instalaran las dependencias (solo la primera vez) y se abrirá el navegador automáticamente

**Opción B: Terminal (línea de comando)**
```bash
cd "ruta\a\Tool Allocation Rev1"
npm install
npm start
```

El dashboard se abrirá en `http://localhost:3000`

---

## Primer uso — Mapeo de columnas

La primera vez que abras el dashboard:

1. Haz clic en el botón **⚙** (engranaje) en la esquina superior derecha
2. Haz clic en **"Discover Columns"** para que se conecte a la BD y liste todas las columnas disponibles
3. En los desplegables, selecciona qué columna corresponde a cada campo:
   - **Tool / Machine ID** ← nombre/ID de la máquina
   - **Cell / Slot** ← nombre de la celda (L01, L02, etc.)
   - **Product** ← producto asignado
   - **Work Week (WW)** ← semana
   - **Week Day** ← día de la semana
   - **TOS Version** ← (opcional) versión del software
4. Haz clic en **"Save & Close"**

Después de esto, se guardará la configuración y no tendrás que hacerlo de nuevo.

---

## Funcionalidades

### Filtros
- **Área**: HDBI, HDMX, PPV
- **Sub-área**: HDBI / BDC (HDBI), PTC / SST / HST (PPV)
- **Máquina**: Selecciona una o varias máquinas
- **Semana (WW)**: Filtra por semana laboral
- **Día**: Filtra por día específico
- **Búsqueda de producto**: Escribe parte del nombre del producto para destacarlo

### Visualización
- Cada máquina se muestra como una tarjeta con sus celdas
- Las celdas ocupadas están coloreadas según el tipo de producto
- Pasa el mouse sobre una celda para ver detalles (Tool, Cell, Product, TOS)
- Haz clic en una tarjeta para ver todos los detalles en una tabla

### Descarga
- **Export CSV**: Descarga los datos filtrados en formato CSV para análisis en Excel

---

## Estructura de directorios

```
Tool Allocation Rev1/
├── server.js              ← Servidor Node.js (backend)
├── package.json           ← Dependencias del proyecto
├── .env                   ← Credenciales de BD (no compartir)
├── .gitignore             ← Archivos ignorados por git
├── start.bat              ← Lanzador para Windows
├── README.md              ← Este archivo
├── CRVLE OPS PLNG...csv   ← Archivo de referencia visual
├── HDBI-HDMX-PPV...csv    ← Mapeo de áreas a máquinas
└── public/                ← Sitio web (frontend)
    ├── index.html         ← Página principal
    ├── app.js             ← Lógica del dashboard (JavaScript)
    └── styles.css         ← Estilos y diseño
```

---

## Seguridad

- **Credenciales de BD**: Se guardan en `.env` (excluido de git)
- **SQL Injection**: Todas las consultas usan parámetros seguros (no hay concatenación de strings)
- **LocalStorage**: La configuración de columnas se guarda en el navegador

---

## Troubleshooting

### Error: "npm is not recognized"
→ Node.js no está instalado. Descárgalo desde [nodejs.org](https://nodejs.org) e instálalo.

### Error de conexión a BD
→ Verifica:
- Tienes acceso a la red interna de Intel
- Las credenciales en `.env` son correctas
- El servidor `sql3243-fm1-in.amr.corp.intel.com:3181` está disponible

### El dashboard no muestra datos después del mapeo
→ Ejecuta "Discover Columns" nuevamente y verifica que los nombres de columnas sean exactos.

### El navegador no abre automáticamente
→ Abre manualmente: `http://localhost:3000`

---

## Datos de ejemplo (si necesitas probar sin BD viva)

Actualmente, el dashboard espera datos en vivo de `ops.vw_SUM_PRODUCT_ALLOCATION`.

Para usar datos exportados en lugar de conectión en vivo, necesitarías:
1. Exportar datos de la vista a un CSV
2. Modificar `server.js` para servir esos datos

---

## Soporte

Para problemas o sugerencias, verifica:
- Que Node.js está actualizado: `node --version`
- Que npm se actualizó: `npm --version`
- Los logs del servidor en la terminal donde ejecutaste `start.bat`

---

**Versión**: 1.0  
**Última actualización**: 2026-04-22
