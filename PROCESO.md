# Proceso de Consolidación de Gastos Legales por País

**Área:** Compliance / Legal — Global66
**Objeto:** Consolidar mensualmente los gastos legales y administrativos de Chile, Colombia, Perú y Argentina, y disponer de un reporte siempre actualizado (pivot país × mes en CLP) que reemplaza el armado manual del Excel `Consolidado Paises.xlsx`.

---

## 1. Principio de funcionamiento

La carpeta de SharePoint **`Consolidado Cobros - Pagos [Compliance]`** es la única fuente de verdad. Cada encargado sube los documentos de su proveedor a la subcarpeta que le corresponde, y la aplicación **lee los archivos directamente desde las carpetas** cada vez que se abre el reporte. No hay pasos intermedios de copiado ni de re-digitación: si el archivo está en la carpeta correcta, aparece en el reporte.

```
Consolidado Cobros - Pagos [Compliance]/
├── Alvaro Moraga - Chile/<año>/Mensual/<N- Mes>/   → NCCL*.pdf (nota de cobro), BH*.pdf (boletas)
├── Dentos - Chile/                                  → FACT*.pdf (facturas exentas)
├── Aninat/                                          → Factura*.pdf, NRG*.pdf
├── Sensus Legis/<Mes>/                              → NC*.pdf (nota de cobro UF), listado xlsx
├── Andes Latam - Perú/Mensual/<N- Mes>/             → minutas de liquidación .xlsx
├── Garrigues - Colombia/                            → planilla mensual .xlsx
├── Carlos Gomez - Colombia/                         → planilla mensual .xlsx
└── Gastos Juridicos - Argentina/                    → planilla de facturas/OC .xlsx
```

## 2. Roles

| Rol | Responsabilidad |
|---|---|
| **Encargado de país/proveedor** | Subir los documentos del mes a su subcarpeta antes del cierre (día 5 hábil del mes siguiente). Verificar en el reporte que su carga aparezca sin alertas. |
| **Owner del consolidado (Compliance)** | Revisar el panel de alertas del reporte, resolver archivos no reconocidos, mantener los parámetros (tipos de cambio, valor UF, IVA) y comunicar el cierre mensual. |
| **TI (una sola vez)** | Registrar la aplicación en Azure AD para habilitar el acceso directo a SharePoint (ver README). |

## 3. Flujo mensual

1. **Días 1–5:** cada encargado sube a su subcarpeta los documentos del mes anterior (factura, nota de cobro o minuta, según el proveedor). Regla de oro: **un archivo por documento del proveedor, en la carpeta del mes que corresponde al servicio**, no al pago.
2. **Día 5:** el owner abre el reporte. La app recorre las carpetas, interpreta cada archivo con el lector específico de su proveedor y arma el consolidado en CLP.
3. **Revisión de alertas:** el panel "Archivos" del reporte muestra tres estados: ✅ procesado, ⚠️ procesado con supuestos (ej. monto en UF convertido con el valor UF configurado), ❌ no reconocido. Los ❌ se corrigen (archivo mal ubicado, escaneado sin texto, formato nuevo) o se registran a mano.
4. **Actualización de parámetros:** el owner mantiene en la pantalla de configuración los tipos de cambio del mes (USD→CLP, PEN→CLP, COP→CLP, ARS→CLP) y el valor de la UF. Quedan guardados en el navegador y son exportables.
5. **Cierre:** con cero alertas ❌, el reporte del mes queda firme. Se puede exportar a CSV/Excel para respaldo o presentación.

## 4. Reglas por proveedor (qué lee la app)

| Proveedor | Archivo esperado | Dato que se extrae | Moneda |
|---|---|---|---|
| **Alvaro Moraga (Chile)** | `BH<N>*.pdf` boleta de honorarios | Total honorarios (bruto) y fecha | CLP |
| | `NCCL <MMAAAA>*.pdf` nota de cobro | Honorario Bruto en UF del período | UF → CLP |
| **Dentons (Chile)** | `FACT <N>*.pdf` factura exenta | Monto Total y fecha de emisión. Los archivos con "ANULADA" en el nombre se ignoran | CLP |
| **Aninat (Chile)** | `Factura*<N>*.pdf`, `NRG*.pdf` | Monto Total y fecha | CLP |
| **Sensus Legis / Yáñez & Acuña (Chile)** | `NC*.pdf` nota de cobro | Total UF del período (hitos Ley 20.009) + IVA | UF → CLP |
| **Andes Latam (Perú)** | minuta `.xlsx` (hojas `1- ok`, `2 - ok`, …) | Honorarios por hora (USD) y gastos reembolsables (PEN), con categoría por fila | USD / PEN → CLP |
| **Garrigues / Carlos Gómez (Colombia)** | planilla `.xlsx` con columnas Fecha, Descripción, Categoría, Valor | Cada fila como un gasto | COP → CLP |
| **Argentina** | planilla `.xlsx` de facturas/OC con columnas Proveedor, Fecha, Total | Cada factura | ARS → CLP |

Cualquier `.xlsx` con columnas reconocibles (Fecha / Descripción / Categoría / Monto) también se ingesta con el lector genérico, aunque el proveedor no esté en la lista.

## 5. Reglas de consolidación

- **Lo ya reportado manda.** La hoja `Base resumen` del `Consolidado Paises.xlsx` se carga como **base histórica**: sus valores se mantienen tal cual (no se recalculan). Si la base histórica ya tiene datos para un país+mes, los archivos de carpetas de ese período se omiten automáticamente para no duplicar. La lectura automática de archivos aplica **desde el primer mes que no esté en la base histórica** en adelante.
- **Misma lógica del Excel, automatizada.** La app construye internamente la misma "base resumen" plana (fecha, categoría, país, CLP, concepto) que antes se armaba a mano, y de ahí derivan las tablas pivot. La base completa se puede exportar a CSV.
- **Carga manual (ficha).** Lo que no llega por documento —un gasto sin factura en la carpeta, o un PDF escaneado ilegible— se registra en la ficha de carga manual con solicitante, proveedor, país, concepto, moneda de origen y monto. Entra al reporte igual que un documento leído y queda identificado como "(carga manual)" en el panel de archivos, para que siempre se distinga lo automático de lo digitado.
- **Clasificación Gasto / Honorario / Juicios.** Cada partida se clasifica automáticamente: primero por regla, y opcionalmente con IA (botón "Clasificar partidas con IA"). El **mantenedor de clasificación** permite corregir cualquier partida; la corrección manual queda guardada y siempre prevalece sobre la automática. La base histórica conserva el concepto ya reportado salvo corrección manual.

## 6. Convenciones que evitan problemas

- **No renombrar las carpetas de proveedor existentes**: el nombre de la carpeta determina el país y el lector que se aplica.
- **Proveedor nuevo**: basta crear su carpeta en la raíz y subir los documentos — se incorpora al reporte automáticamente. Nombrar la carpeta terminando en el país (ej. `Estudio Pérez - Chile`) para que se asigne bien; si el país no aparece en el nombre, el reporte lo muestra como "Sin país" hasta corregirlo. Los proveedores nuevos usan los lectores genéricos (factura PDF con "Monto Total", o planilla con columnas Fecha + Monto); si su formato es distinto, se agrega un lector específico.
- **Carpetas de mes**: usar el formato existente `N- Mes` (ej. `3- Marzo`). El número manda; el nombre es informativo.
- **PDFs con texto**: subir el PDF electrónico original (SII, nota de cobro generada), no un escaneo/foto. Un PDF escaneado no se puede leer automáticamente y quedará en ❌.
- **Anulaciones**: si un documento se anula, agregar "ANULADA" al nombre del archivo; la app lo excluye del consolidado.
- **No duplicar**: antes de subir una versión corregida, eliminar la anterior (los archivos `(1)`, `Copia de …` generan riesgo de doble conteo; la app detecta duplicados por número de documento y monto, pero la carpeta limpia es la primera defensa).
- **Contingencias judiciales (cuantías demandadas)** no viven en las carpetas de proveedor: se registran en la planilla de juicios (hoja "Juicios") y el reporte las muestra separadas del gasto efectivo, igual que hoy.

## 7. Qué reemplaza y qué no

- ✅ Reemplaza: el armado manual de `Base resumen` y las tablas dinámicas del `Consolidado Paises.xlsx`.
- ✅ Mantiene: la misma lógica de negocio (montos en CLP con IVA, apertura Gastos/Honorarios, pivot país × mes con totales por año).
- ❌ No reemplaza (por ahora): el registro de cuantías de juicios en curso y las investigaciones internas, que siguen siendo captura manual.

## 8. Evolución sugerida

1. **Fase 1 (esta app):** lectura automática de los formatos actuales de cada proveedor.
2. **Fase 2:** pedir a los proveedores de solo-PDF un anexo estándar en Excel (misma plantilla para todos), reduciendo los lectores específicos.
3. **Fase 3:** tipos de cambio automáticos (API BCCh/SBS) y publicación del reporte con refresco programado.
