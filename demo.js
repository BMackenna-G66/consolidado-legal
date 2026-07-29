// Datos FICTICIOS de demostración: permiten ver el reporte completo (tablas,
// mantenedor, filtros) en la versión publicada sin exponer información real.
// Los montos y casos son inventados.

const D = [
  // pais, proveedor, categoria, concepto, mes, anio, clp, fuente
  ['Chile', 'Reportado (base histórica)', 'Asesoría Profesional', 'Gastos', 1, 2025, 3200000, 'base'],
  ['Chile', 'Reportado (base histórica)', 'Asesoría Profesional', 'Gastos', 2, 2025, 3150000, 'base'],
  ['Chile', 'Reportado (base histórica)', 'Asesoría Profesional', 'Gastos', 3, 2025, 3300000, 'base'],
  ['Chile', 'Reportado (base histórica)', 'Honorarios varios', 'Gastos', 4, 2025, 2900000, 'base'],
  ['Chile', 'Reportado (base histórica)', 'Querellas penales', 'Juicios y otros', 1, 2025, 15000000, 'base'],
  ['Chile', 'Reportado (base histórica)', 'Demandas laborales (ex-empleados)', 'Juicios y otros', 2, 2025, 42000000, 'base'],
  ['Colombia', 'Reportado (base histórica)', 'Mercantil Corp. M&A', 'Gastos', 1, 2025, 700000, 'base'],
  ['Colombia', 'Reportado (base histórica)', 'Asesorías legales', 'Gastos', 2, 2025, 460000, 'base'],
  ['Colombia', 'Reportado (base histórica)', 'Conciliaciones extrajudiciales', 'Gastos', 4, 2025, 105000, 'base'],
  ['Perú', 'Reportado (base histórica)', 'Documentación contractual', 'Gastos', 1, 2025, 1050000, 'base'],
  ['Perú', 'Reportado (base histórica)', 'Cartas notariales y gestiones', 'Gastos', 2, 2025, 255000, 'base'],
  ['Perú', 'Reportado (base histórica)', 'Honorarios Licencia Fintech', 'Honorarios', 6, 2025, 3800000, 'base'],
  ['Argentina', 'Reportado (base histórica)', 'Legalización de firmas', 'Gastos', 4, 2025, 1450000, 'base'],
  ['Argentina', 'Reportado (base histórica)', 'Asesoría Profesional', 'Gastos', 2, 2025, 98000, 'base'],
];

const F = [
  // Archivos "nuevos" que la app leería de las carpetas (posteriores a la base)
  { archivo: 'BH5210 Global (demo).pdf', ruta: '/Alvaro Moraga - Chile/2026/Mensual/1- Enero/BH5210 Global (demo).pdf', estado: 'ok', nota: 'Bruto 3.100.000 − retención 449.500 = neto 2.650.500 (demo)', regs: [['Chile', 'Alvaro Moraga', 'Asesoría Profesional', 'Gastos', 1, 2026, 2650500]] },
  { archivo: 'FACT 8001- DENTONS (demo).pdf', ruta: '/Dentos - Chile/FACT 8001- DENTONS (demo).pdf', estado: 'ok', nota: 'Monto total 1.900.000 (demo)', regs: [['Chile', 'Dentons', 'Honorarios varios', 'Honorarios', 2, 2026, 1900000]] },
  { archivo: 'NC Global enero 2026 (demo).pdf', ruta: '/Sensus Legis/Enero/NC Global enero 2026 (demo).pdf', estado: 'supuesto', nota: 'UF 18 × $39.894 + IVA 19% = 854.529 (demo, valor UF configurable)', regs: [['Chile', 'Sensus Legis', 'Honorarios Ley 20.009', 'Honorarios', 1, 2026, 854529]] },
  { archivo: 'Andes Latam - enero 26 (demo).xlsx', ruta: '/Andes Latam - Perú/Mensual/1- Enero/Andes Latam - enero 26 (demo).xlsx', estado: 'supuesto', nota: '12 filas (USD→CLP 980 +IGV 18%) (demo)', regs: [
    ['Perú', 'Andes Latam', 'Documentación contractual', 'Gastos', 1, 2026, 640000],
    ['Perú', 'Andes Latam', 'Junta General y Poderes', 'Gastos', 1, 2026, 380000],
    ['Perú', 'Andes Latam', 'Gastos reembolsables', 'Gastos', 1, 2026, 45000],
  ] },
  { archivo: 'Factura escaneada (demo).pdf', ruta: '/Aninat/Factura escaneada (demo).pdf', estado: 'error', nota: 'PDF sin montos legibles (¿escaneado?) — ejemplo de alerta', regs: [] },
  { archivo: 'Copia de minuta (demo).xlsx', ruta: '/Andes Latam - Perú/Mensual/1- Enero/Copia de minuta (demo).xlsx', estado: 'omitido', nota: 'Duplicado — se omite (ejemplo)', regs: [] },
];

function reg([pais, proveedor, categoria, concepto, mes, anio, clp, fuente]) {
  return { dia: 1, mes, anio, pais, proveedor, categoria, concepto, moneda: 'CLP', montoOrigen: clp, clp, archivo: '(demo)', detalle: 'Dato ficticio de demostración', fuente: fuente || 'archivo' };
}

export function resultadosDemo() {
  const base = { archivo: 'Consolidado Paises (demo).xlsx', ruta: '/Consolidado Paises (demo).xlsx', estado: 'ok', fuente: 'base', nota: `${D.length} movimientos ya reportados — datos ficticios de demostración`, registros: D.map(reg) };
  const archivos = F.map(f => ({ archivo: f.archivo, ruta: f.ruta, estado: f.estado, nota: f.nota, registros: f.regs.map(reg) }));
  return [base, ...archivos];
}
