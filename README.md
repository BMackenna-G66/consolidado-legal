# Consolidado Legal por País

Aplicación web que reemplaza el armado manual de `Consolidado Paises.xlsx`: lee los documentos de cada proveedor directamente desde las carpetas de SharePoint (`Consolidado Cobros - Pagos [Compliance]`), interpreta boletas, facturas, notas de cobro y minutas, y construye el reporte de gastos por país × mes en CLP — con el mismo formato de tabla dinámica del Excel original.

El proceso de negocio (roles, flujo mensual, convenciones) está en [PROCESO.md](PROCESO.md).

## Cómo funciona

1. **Fuente de datos** (a elección al abrir la app):
   - **Carpeta local / OneDrive**: seleccionas la carpeta de SharePoint sincronizada con OneDrive. Funciona sin ninguna configuración y la app recuerda la carpeta para las siguientes veces.
   - **SharePoint en línea**: login con la cuenta Global66 (MSAL) y lectura vía Microsoft Graph. Requiere el registro en Azure AD descrito abajo.
2. **Parsers por proveedor** ([parsers.js](parsers.js)): la subcarpeta determina el lector — boletas/notas de cobro de Moraga, facturas exentas de Dentons/Aninat, notas de cobro UF de Sensus, minutas xlsx de Andes Latam, planillas tabulares de Colombia/Argentina. Archivos "ANULADA" se excluyen; si un mes de Moraga tiene boleta y nota de cobro, la boleta manda (no se duplica).
3. **Base histórica**: la hoja `Base resumen` del `Consolidado Paises.xlsx` se ingesta tal cual — los valores ya reportados se mantienen y tienen prioridad: si un país+mes ya está en la base, los archivos de ese período se omiten (no hay doble conteo). La lectura automática rige para los meses nuevos.
4. **Normalización**: todo se convierte a CLP con los parámetros configurables (USD, PEN, COP, ARS, UF, IVA/IGV) que replican los supuestos del Excel. Se editan en ⚙ Parámetros y quedan en localStorage.
5. **Clasificación Gasto/Honorario/Juicios** ([clasificador.js](clasificador.js)): automática por regla o con IA (Gemini, mismo patrón de API key que SmartCheck), con un **mantenedor** en la UI donde la corrección manual queda guardada y siempre prevalece.
6. **Reporte**: tarjetas de totales y tabla dinámica de cuatro niveles — **país → proveedor/abogado → concepto cobrado → solicitante** — con columnas por mes agrupadas por año (Total 2025, Total 2026, Total general). El desplegable *Abrir por* cambia el orden de la jerarquía (por solicitante, por proveedor o por concepto), el botón *Expandir todo* abre el árbol completo y cada fila muestra cuántos movimientos agrupa. Filtros por concepto, año y mes, y exportación a CSV con solicitante y equivalente en USD.

   El **solicitante** sale del propio documento cuando lo trae ("Solicitado por…"). Las boletas de honorarios no lo incluyen, así que se toma del anexo de detalle del mismo mes: si el anexo nombra varios, la fila queda como *Varios (n)* y los nombres aparecen en la nota del archivo. Donde ninguna fuente lo aporta —la base histórica, por ejemplo— se agrupa como *Sin solicitante registrado*. El panel "Archivos leídos" muestra qué se procesó, con qué supuestos, y qué quedó fuera — nada entra al número sin dejar rastro.

Todo el procesamiento ocurre en el navegador; ningún documento sale hacia servicios de terceros.

## Ejecutar (uso diario)

**Doble clic en `Abrir Consolidado Legal.command`.** Eso hace todo:

1. Copia los archivos desde la carpeta de SharePoint sincronizada con OneDrive (`~/Library/CloudStorage/OneDrive-Global81SPA/Consolidado Cobros - Pagos [Compliance]`) a `datos/` y genera su manifiesto.
2. Levanta el servidor local (`servidor.py`, sin caché).
3. Abre el navegador **con el reporte ya generado** — sin elegir carpetas ni hacer clics.

Para ver las cifras actualizadas después de que alguien suba documentos nuevos: doble clic otra vez.

> En macOS, la primera vez puede pedir confirmación por ser un archivo descargado: clic derecho → Abrir.
> Si el navegador mostrara datos viejos, recarga con Cmd+Shift+R.

Requisito: la carpeta de SharePoint debe estar sincronizada con OneDrive (botón *Sincronizar* en SharePoint). Si no lo está, la app abre en la portada y permite elegir cualquier carpeta a mano; en Chrome/Edge la recuerda para las siguientes veces.

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

Las fichas se guardan en el navegador de quien las carga (localStorage), así que conviene exportarlas para respaldo o para consolidar el trabajo de varias personas.

## Agregar un proveedor nuevo

No requiere tocar código: crear la carpeta en la raíz de SharePoint y subir los documentos. La app detecta la carpeta nueva, deduce el país del nombre (terminar en `- Chile`, `- Colombia`, `- Perú` o `- Argentina`) y aplica los lectores genéricos, marcando los movimientos con "🆕 Proveedor nuevo detectado" en el panel de archivos.

Solo hace falta editar [config.js](config.js) (lista `PROVEEDORES`) si el proveedor necesita un lector específico porque su formato no es una factura PDF con "Monto Total" ni una planilla con columnas Fecha + Monto.

## Limitaciones conocidas

- Un PDF **escaneado sin capa de texto** no se puede leer: queda marcado ❌ en el panel de archivos y debe registrarse a mano (o re-subir el PDF electrónico).
- Los tipos de cambio son manuales (mismos supuestos que el Excel). Fase futura: API BCCh/SBS.
- Las cuantías de juicios en curso ("Juicios y otros") no viven en las carpetas de proveedor y siguen fuera de este reporte, igual que las investigaciones internas.
