# Consolidado Legal por País

Aplicación web que reemplaza el armado manual de `Consolidado Paises.xlsx`: lee los documentos de cada proveedor directamente desde las carpetas de SharePoint (`Consolidado Cobros - Pagos [Compliance]`), interpreta boletas, facturas, notas de cobro y minutas, y construye el reporte de gastos por país × mes en CLP — con el mismo formato de tabla dinámica del Excel original.

El proceso de negocio (roles, flujo mensual, convenciones) está en [PROCESO.md](PROCESO.md).

## Cómo funciona

1. **Fuente de datos: SharePoint en vivo** ([fuentes.js](fuentes.js)). Login con la cuenta Global66 (MSAL) y lectura vía Microsoft Graph, en el momento de abrir el reporte. La carpeta manual queda solo como camino de emergencia.
2. **Parsers por proveedor** ([parsers.js](parsers.js)): la subcarpeta determina el lector — boletas/notas de cobro de Moraga, facturas exentas de Dentons/Aninat, notas de cobro UF de Sensus, minutas xlsx de Andes Latam, planillas tabulares de Colombia/Argentina. Archivos "ANULADA" se excluyen; si un mes de Moraga tiene boleta y nota de cobro, la boleta manda (no se duplica).
3. **Base histórica**: la hoja `Base resumen` del `Consolidado Paises.xlsx` se ingesta tal cual — los valores ya reportados se mantienen y tienen prioridad: si un país+mes ya está en la base, los archivos de ese período se omiten (no hay doble conteo). La lectura automática rige para los meses nuevos.
4. **Normalización**: todo se convierte a CLP con los parámetros configurables (USD, PEN, COP, ARS, UF, IVA/IGV) que replican los supuestos del Excel. Se editan en ⚙ Parámetros y quedan en localStorage.
5. **Clasificación Gasto/Honorario/Juicios** ([clasificador.js](clasificador.js)): automática por regla o con IA (Gemini, mismo patrón de API key que SmartCheck), con un **mantenedor** en la UI donde la corrección manual queda guardada y siempre prevalece.
6. **Reporte**: tarjetas de totales y tabla dinámica de cuatro niveles — **país → proveedor/abogado → concepto cobrado → solicitante** — con columnas por mes agrupadas por año (Total 2025, Total 2026, Total general). El desplegable *Abrir por* cambia el orden de la jerarquía (por solicitante, por proveedor o por concepto), el botón *Expandir todo* abre el árbol completo y cada fila muestra cuántos movimientos agrupa. Filtros por concepto, año y mes, y exportación a CSV con solicitante y equivalente en USD.

   El **solicitante** sale del propio documento cuando lo trae ("Solicitado por…"). Las boletas de honorarios no lo incluyen, así que se toma del anexo de detalle del mismo mes: si el anexo nombra varios, la fila queda como *Varios (n)* y los nombres aparecen en la nota del archivo. Donde ninguna fuente lo aporta —la base histórica, por ejemplo— se agrupa como *Sin solicitante registrado*. El panel "Archivos leídos" muestra qué se procesó, con qué supuestos, y qué quedó fuera — nada entra al número sin dejar rastro.

Todo el procesamiento ocurre en el navegador; ningún documento sale hacia servicios de terceros.

## Cómo se usa (100% web)

**Abre https://bmackenna-g66.github.io/consolidado-legal/ y pulsa *Conectar con Microsoft*.** La app lee los documentos directamente desde SharePoint en ese momento y arma el reporte. No hay carpeta que sincronizar, ni copias locales que se atrasen, ni nada que mantener: si un encargado sube una factura, aparece en la siguiente carga.

Requisito único: el **Application (client) ID** del registro de Azure AD (ver más abajo). La app lo pide la primera vez y lo guarda en el navegador de cada persona. El inicio de sesión ocurre en la ventana oficial de Microsoft — la app nunca ve contraseñas, solo recibe un permiso de lectura temporal.

> **Alternativa mientras no exista el clientId:** el botón *Elegir carpeta…* permite generar el reporte desde una carpeta descargada de SharePoint (o su ZIP descomprimido). Es un camino de emergencia, no el flujo normal.

## Ejecutar en local (solo desarrollo)

`Abrir Consolidado Legal.command` levanta el servidor local y abre el navegador. Sirve para desarrollar; para el uso real basta el sitio publicado.

### Versión publicada

La aplicación está en **https://bmackenna-g66.github.io/consolidado-legal/**. Es solo el programa: **no contiene ningún dato**. Al abrirla, cada persona elige su carpeta de SharePoint sincronizada (o conecta con Microsoft) y el reporte se arma en su propio navegador. Las carpetas `datos/` y `muestras/` están en `.gitignore`, y el enlace interno de SharePoint no vive en el código — la app lo pide y lo guarda localmente.

Para usarla sobre la carpeta ya sincronizada de este equipo, sigue siendo más cómodo el lanzador local descrito arriba.

## Habilitar el modo SharePoint (una vez, TI)

1. En [Azure Portal → App registrations](https://portal.azure.com) crear registro "Consolidado Legal Compliance":
   - Supported account types: *Single tenant*.
   - Platform: **Single-page application (SPA)** con la URL donde se sirva la app como Redirect URI (ej. `https://<org>.github.io/<repo>/consolidado-legal/` y `http://localhost:4173` para desarrollo).
2. API permissions → Microsoft Graph → **Delegated** → `Files.Read.All` (+ Grant admin consent si la política lo exige).
3. Copiar el **Application (client) ID**. Se puede pegar directamente en la app (botón *Conectar con Microsoft* lo pide la primera vez y lo guarda en el navegador) o dejarlo fijo en `config.js` → `AZURE.clientId`.

Con eso, *Conectar con Microsoft* abre la ventana oficial de Microsoft, cada persona inicia sesión con su propia cuenta y la app lee SharePoint en vivo, bajo demanda. **La aplicación nunca ve ni almacena contraseñas**: solo recibe un token de acceso temporal emitido por Microsoft.

Sin este registro la app funciona igual en modo carpeta local.

## Estructura

| Archivo | Rol |
|---|---|
| [config.js](config.js) | Proveedores/carpetas, parámetros por defecto, utilidades de fecha y números |
| [fuentes.js](fuentes.js) | Lectura de carpeta local y de SharePoint (MSAL + Graph) |
| [parsers.js](parsers.js) | Lectores por proveedor (PDF vía pdf.js, XLSX vía SheetJS) y deduplicación |
| [reporte.js](reporte.js) | Pivot país × mes, tarjetas, export CSV |
| [app.js](app.js) | Orquestación, filtros, panel de parámetros, mantenedor, fichas |
| [manual.js](manual.js) | Fichas de carga manual: almacenamiento, cálculo CLP/USD, export/import |

## Carga manual (ficha)

El botón **➕ Carga manual** abre una ficha para registrar movimientos que no tienen documento en las carpetas, o cuyo PDF está escaneado y no se puede leer. Campos: período, **solicitante**, proveedor, país, tipo de concepto, categoría, **moneda de origen**, monto, **monto en CLP** (calculado con los parámetros, editable) y **equivalente en USD** (derivado).

Cada ficha entra al reporte como una fuente más: suma en el pivot, aparece en el mantenedor de clasificación y se incluye en el CSV. Se pueden editar y borrar, y el botón *Exportar / Importar fichas* mueve el registro completo entre equipos como un JSON — útil hasta que la captura se automatice.

### Trabajo en equipo: una plantilla por persona

Las fichas se guardan en el navegador de quien las carga, así que para consolidar el trabajo de varias personas se usa la carpeta compartida como punto de encuentro:

1. Cada persona abre la app (el link publicado sirve), llena sus fichas y pulsa **⬇ Excel para SharePoint**.
2. Sube ese archivo — `Carga manual - <nombre> - <año-mes>.xlsx` — a la carpeta del proveedor en SharePoint, junto a las facturas.
3. La app reconoce esa plantilla por sus encabezados y la lee como un documento más: los movimientos digitados por todos aparecen en el consolidado, con su solicitante.

El archivo se puede llenar también a mano en Excel, respetando los encabezados (`Fecha, Solicitante, Proveedor, Pais, Concepto, Categoria, Moneda, Monto origen, Monto CLP, Monto USD, Detalle`). Así los datos quedan siempre dentro de SharePoint, sin depender de que nadie comparta su navegador.

## Base maestra en Excel

El botón **⬇ Base maestra Excel** descarga un libro con todo lo que la app leyó y calculó:

| Hoja | Contenido |
|---|---|
| Base maestra | Un movimiento por fila con 24 columnas: período, país, proveedor, solicitante, concepto cobrado, tipo, moneda y monto de origen, CLP, USD, detalle, **origen del dato**, carpeta de proveedor, carpeta de mes, archivo, ruta completa, estado y nota de lectura |
| Por país / proveedor / concepto / solicitante / tipo / origen | Totales en CLP y USD de cada corte |
| Resumen país x mes | La misma tabla del reporte en pantalla |
| Detalle línea por línea | El desglose que traen los documentos por dentro, sin agrupar: fecha, categoría, **solicitante**, **profesional/abogado**, descripción del trabajo, horas y valor. Sale de los anexos de Moraga (bloque "DETALLE TRABAJOS"), de los timesheets y gastos de las minutas de Andes Latam, y de los hitos de Sensus |
| Archivos leídos | Los 90 documentos con su estado, cuántos movimientos aportó cada uno y con qué supuestos |
| Parámetros | Tipos de cambio usados, fecha de generación y totales de control |

> Los montos viven a nivel de documento (una boleta, una factura); el detalle interno aporta horas y descripciones, no montos por línea. Por eso son dos hojas y no una: sumar el detalle **no** da el total del consolidado.

También se puede generar sin abrir el navegador, ejecutando los mismos lectores en Node:

```bash
cd scripts && npm install xlsx pdfjs-dist@3.11.174 && node base-maestra.mjs
```

## Agregar un proveedor nuevo

No requiere tocar código: crear la carpeta en la raíz de SharePoint y subir los documentos. La app detecta la carpeta nueva, deduce el país del nombre (terminar en `- Chile`, `- Colombia`, `- Perú` o `- Argentina`) y aplica los lectores genéricos, marcando los movimientos con "🆕 Proveedor nuevo detectado" en el panel de archivos.

Solo hace falta editar [config.js](config.js) (lista `PROVEEDORES`) si el proveedor necesita un lector específico porque su formato no es una factura PDF con "Monto Total" ni una planilla con columnas Fecha + Monto.

## Limitaciones conocidas

- Un PDF **escaneado sin capa de texto** no se puede leer: queda marcado ❌ en el panel de archivos y debe registrarse a mano (o re-subir el PDF electrónico).
- Los tipos de cambio son manuales (mismos supuestos que el Excel). Fase futura: API BCCh/SBS.
- Las cuantías de juicios en curso ("Juicios y otros") no viven en las carpetas de proveedor y siguen fuera de este reporte, igual que las investigaciones internas.
