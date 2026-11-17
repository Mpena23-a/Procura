# COLVEN · Dashboard de Seguimiento de Equipos y Mantenimiento

Reporte gerencial estático (HTML + JavaScript) que lee directamente el Excel
**`data/Seguimiento eq Mantenimiento.xlsx`** y muestra:

| Sección | Hoja del Excel | Qué muestra |
|---|---|---|
| **Resumen general** | Todas | KPIs consolidados, tarjetas por segmento y puntos de atención automáticos |
| **Mantenimiento** | `MTTO` | Estado, prioridad, responsables, tipo de orden/región y antigüedad |
| **Operaciones** | `OPER` | Igual que Mantenimiento, con principales proveedores |
| **Materiales críticos** | `Materiales Criticos` | Inventario y niveles mínimo/máximo (muestra un aviso si aún está vacía) |
| **Anticipos y montos** | `RQ (2)` | Monto total, anticipos, saldo por pagar e incidencias |

La primera vista de cada sección es un resumen. Todo es interactivo:

- **Cada cuadro (KPI o tarjeta de segmento)** abre justo debajo un panel de detalle
  con indicadores, gráficos y la tabla de los registros que componen ese número.
  Un segundo clic (o la ✕) lo cierra. Algunos paneles traen un botón para
  **filtrar la tabla** o **abrir la sección** correspondiente.
- **Barras de los gráficos** de estado, prioridad y responsables filtran la tabla de detalle.
- **Ver detalle completo** abre la tabla con búsqueda, filtros y orden por columna.
  Los gráficos **Tipo de orden y región** y **Principales proveedores** (y los de
  Anticipos) se ajustan a esos filtros y muestran la etiqueta *Filtrado · N de M*.

---

## 1. Publicar en GitHub Pages (una sola vez)

1. Crea un repositorio nuevo en GitHub (por ejemplo `dashboard-mantenimiento`).
2. Sube **todo el contenido de esta carpeta** (`index.html`, `assets/`, `data/`, `README.md`).
   - Desde la web: *Add file → Upload files* y arrastra los archivos y carpetas.
3. Ve a **Settings → Pages**.
   - *Source*: **Deploy from a branch**
   - *Branch*: **main** y carpeta **/ (root)** → **Save**
4. En 1–2 minutos la página estará disponible en
   `https://<tu-usuario>.github.io/<nombre-del-repositorio>/`

> ⚠️ **Confidencialidad:** con una cuenta gratuita, GitHub Pages solo funciona con
> repositorios **públicos**, por lo que cualquiera con el enlace podría ver los datos
> y descargar el Excel. Para mantenerlo privado se necesita GitHub Pro/Team/Enterprise
> (Pages en repositorio privado con acceso restringido) u otro hosting interno.

## 2. Actualizar los datos

1. Actualiza el Excel manteniendo **la misma estructura** (mismas hojas y encabezados).
2. En GitHub entra a la carpeta **`data/`** → *Add file → Upload files*.
3. Sube el archivo con **el mismo nombre**: `Seguimiento eq Mantenimiento.xlsx`
   (GitHub reemplaza el anterior) → *Commit changes*.
4. Espera 1–2 minutos y recarga la página. El botón **Descargar Excel** entregará la versión nueva.

> Si prefieres otro nombre de archivo, cámbialo en `assets/js/config.js` (`dataFile`).

### Botón “Cargar Excel” (con contraseña)
Pide una contraseña y luego permite **previsualizar** cualquier Excel con la misma
estructura. El cambio es solo en el navegador de quien lo carga; **no modifica la
página publicada** (eso solo se hace subiendo el archivo a `data/`, paso anterior).
Tras ingresarla bien, no se vuelve a pedir hasta cerrar la pestaña.

- La contraseña se guarda en `assets/js/config.js` como huella **SHA-256**, no en texto plano.
- Para cambiarla, calcula el SHA-256 de la nueva (en PowerShell):
  ```powershell
  $b=[Text.Encoding]::UTF8.GetBytes('NuevaClave'); ([Security.Cryptography.SHA256]::Create().ComputeHash($b) | % { $_.ToString('x2') }) -join ''
  ```
  y pega el resultado en `uploadPasswordHash`.
- ⚠️ Es una protección básica de una página estática: evita cargas casuales, pero no
  sustituye un control de acceso real. Los datos publicados siguen siendo visibles para
  quien tenga el enlace.

## 3. Reglas y supuestos

- **Estado (estimado):** se calcula a partir de palabras clave del campo *Comentarios*:
  | Palabras clave | Estado |
  |---|---|
  | “NO FUE APROBADO”, “CANALIZAR” | Canalizado / no aprobado |
  | “TRABAJO 100%”, “COMPLETADO” | Completado |
  | “POR APROB” | Por aprobación |
  | “POR ANTICIPO”, “PENDIENTE PAGO”, “SIN INFORMACIÓN DE PAGO”, “PARA INICIO” | Pago / anticipo pendiente |
  | “ENTREGA PENDIENTE”, “FABRICACIÓN” | Entrega pendiente |
  | Sin número de orden | En gestión (sin orden) |
  | “ANTICIPO”, “PAGO”, “TERMINA”, “EXTENDER”, “OTROSÍ” | En ejecución |
  | Comentario vacío | Sin comentario |
- **Región:** por el código de la orden (`AP` = Apure, `BA` = Barinas, `AB` = Apure / Barinas);
  si no hay orden, por la descripción.
- **Antigüedad:** días desde la *Fecha de Emisión de la Orden* hasta hoy. El umbral de
  alerta (45 días) se ajusta en `assets/js/config.js`.
- **Anticipos:** `Anticipo = Monto × % Anticipo`. Si la columna % contiene texto
  (p. ej. “Factura retrasada”) se registra como **incidencia**.
- Las filas que solo tienen el correlativo (No.) se ignoran. La hoja oculta `Hoja4` (listas) no se usa.
- Los encabezados se reconocen por nombre, por lo que se pueden reordenar columnas sin romper el tablero.

## 4. Ver localmente

Abrir `index.html` con doble clic no carga el Excel automáticamente (restricción del
navegador); usa **Cargar Excel** o publícalo en GitHub Pages.

## Estructura

```
index.html
assets/
  css/styles.css     ← estilo visual COLVEN
  js/config.js       ← nombre del archivo, moneda, umbral de antigüedad
  js/app.js          ← lectura del Excel, KPIs, gráficos y tablas
  img/logo-colven.png
data/
  Seguimiento eq Mantenimiento.xlsx
```

Librerías (CDN): [SheetJS](https://sheetjs.com/) para leer Excel y [Chart.js](https://www.chartjs.org/) para los gráficos.
