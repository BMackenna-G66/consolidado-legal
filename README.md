# Consolidado Legal por País

Aplicación web que reemplaza el armado manual de `Consolidado Paises.xlsx`: lee los documentos de cada proveedor directamente desde las carpetas de SharePoint (`Consolidado Cobros - Pagos [Compliance]`), interpreta boletas, facturas, notas de cobro y minutas, y construye el reporte de gastos por país × mes en CLP — con el mismo formato de tabla dinámica del Excel original.

El proceso de negocio (roles, flujo mensual, convenciones) está en [PROCESO.md](PROCESO.md).

## Cómo funciona

1. **Fuente de datos** (a elección al abrir la app):
   - **Carpeta local / OneDrive**: seleccionas la carpeta sincronizada (o una copia descargada). Funciona sin ninguna configuración.
   - **SharePoint en línea**: login con la cuenta Global66 (MSAL) y lectura vía Microsoft Graph. Requiere el registro en Azure AD descrito abajo.
2. **Parsers por proveedor** ([parsers.js](parsers.js)): la subcarpeta determina el lector — boletas/notas de cobro de Moraga, facturas exentas de Dentons/Aninat, notas de cobro UF de Sensus, minutas xlsx de Andes Latam, planillas tabulares de Colombia/Argentina. Archivos "ANULADA" se excluyen; si un mes de Moraga tiene boleta y nota de cobro, la boleta manda (no se duplica).
3. **Base histórica**: la hoja `Base resumen` del `Consolidado Paises.xlsx` se ingesta tal cual — los valores ya reportados se mantienen y tienen prioridad: si un país+mes ya está en la base, los archivos de ese período se omiten (no hay doble conteo). La lectura automática rige para los meses nuevos.
4. **Normalización**: todo se convierte a CLP con los parámetros configurables (USD, PEN, COP, ARS, UF, IVA/IGV) que replican los supuestos del Excel. Se editan en ⚙ Parámetros y quedan en localStorage.
5. **Clasificación Gasto/Honorario/Juicios** ([clasificador.js](clasificador.js)): automática por regla o con IA (Gemini, mismo patrón de API key que SmartCheck), con un **mantenedor** en la UI donde la corrección manual queda guardada y siempre prevalece.
6. **Reporte**: tarjetas de totales, pivot país → proveedor → categoría con columnas por mes agrupadas por año (Total 2025, Total 2026, Total general), filtro por concepto y exportación a CSV. El panel "Archivos leídos" muestra qué se procesó, con qué supuestos, y qué quedó fuera — nada entra al número sin dejar rastro.

Todo el procesamiento ocurre en el navegador; ningún documento sale hacia servicios de terceros.

## Ejecutar

Es una app estática, sin build:

```bash
cd consolidado-legal && python3 -m http.server 4173
```

y abrir `http://localhost:4173`. Para publicarla (GitHub Pages u otro hosting estático) basta copiar esta carpeta. La carpeta `muestras/` está en `.gitignore` porque contiene documentos reales — no debe commitearse.

Modo prueba: `http://localhost:4173/?demo` carga las muestras de `muestras/manifest.json`.

## Habilitar el modo SharePoint (una vez, TI)

1. En [Azure Portal → App registrations](https://portal.azure.com) crear registro "Consolidado Legal Compliance":
   - Supported account types: *Single tenant*.
   - Platform: **Single-page application (SPA)** con la URL donde se sirva la app como Redirect URI (ej. `https://<org>.github.io/<repo>/consolidado-legal/` y `http://localhost:4173` para desarrollo).
2. API permissions → Microsoft Graph → **Delegated** → `Files.Read.All` (+ Grant admin consent si la política lo exige).
3. Copiar el **Application (client) ID** en `config.js` → `AZURE.clientId`.

Sin este registro la app funciona igual en modo carpeta local.

## Estructura

| Archivo | Rol |
|---|---|
| [config.js](config.js) | Proveedores/carpetas, parámetros por defecto, utilidades de fecha y números |
| [fuentes.js](fuentes.js) | Lectura de carpeta local y de SharePoint (MSAL + Graph) |
| [parsers.js](parsers.js) | Lectores por proveedor (PDF vía pdf.js, XLSX vía SheetJS) y deduplicación |
| [reporte.js](reporte.js) | Pivot país × mes, tarjetas, export CSV |
| [app.js](app.js) | Orquestación, filtros, panel de parámetros, modo demo |

## Limitaciones conocidas

- Un PDF **escaneado sin capa de texto** no se puede leer: queda marcado ❌ en el panel de archivos y debe registrarse a mano (o re-subir el PDF electrónico).
- Los tipos de cambio son manuales (mismos supuestos que el Excel). Fase futura: API BCCh/SBS.
- Las cuantías de juicios en curso ("Juicios y otros") no viven en las carpetas de proveedor y siguen fuera de este reporte, igual que las investigaciones internas.
