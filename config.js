// Configuración central del Consolidado Legal.
// Todo lo que un owner necesita ajustar vive aquí o en la pantalla de Parámetros (localStorage).

export const AZURE = {
  // Completar tras registrar la app en Azure AD (ver README.md). Vacío ⇒ solo modo carpeta local.
  clientId: '',
  tenantId: 'global66.onmicrosoft.com',
  // Enlace compartido de la carpeta raíz en SharePoint (repositorio privado).
  shareUrl: 'https://global66-my.sharepoint.com/:f:/g/personal/compliance4_global66_onmicrosoft_com/IgBVjhlzYfeBQpXWsxYcaDQJAajHV2IQt6MVbpahvsCtajo',
};

// Parámetros de conversión editables en la UI; estos son los valores por defecto.
// Mismos supuestos que el Excel original (USD→CLP 980, PEN→USD 3.68, IGV Perú 18%).
export const DEFAULT_PARAMS = {
  USD_CLP: 980,      // CLP por USD
  PEN_CLP: 270.31,   // CLP por PEN (gastos Perú se facturan en soles)
  COP_CLP: 0.23,     // CLP por COP
  ARS_CLP: 0.84,     // CLP por ARS
  UF_CLP: 39894.11,  // CLP por UF
  IVA_PE: 0.18,      // IGV aplicado a honorarios Perú en el consolidado
  IVA_CL: 0.19,      // IVA notas de cobro afectas (Sensus)
};

export const MESES = ['ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO','JULIO','AGOSTO','SEPTIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE'];

const MES_NOMBRES = {
  enero: 1, febrero: 2, marzo: 3, abril: 4, mayo: 5, junio: 6, julio: 7,
  agosto: 8, septiembre: 9, setiembre: 9, octubre: 10, noviembre: 11, diciembre: 12,
};

// Mapea carpeta de proveedor → parser y país. `match` se prueba contra la ruta relativa del archivo.
export const PROVEEDORES = [
  { id: 'moraga',   pais: 'Chile',     nombre: 'Alvaro Moraga',      match: /alvaro moraga/i },
  { id: 'dentons',  pais: 'Chile',     nombre: 'Dentons',            match: /dento(n)?s/i },
  { id: 'aninat',   pais: 'Chile',     nombre: 'Aninat',             match: /aninat/i },
  { id: 'sensus',   pais: 'Chile',     nombre: 'Sensus Legis',       match: /sensus/i },
  { id: 'andes',    pais: 'Perú',      nombre: 'Andes Latam',        match: /andes latam/i },
  { id: 'colombia', pais: 'Colombia',  nombre: 'Garrigues',          match: /garrigues/i },
  { id: 'colombia', pais: 'Colombia',  nombre: 'Carlos Gómez',       match: /carlos gomez/i },
  { id: 'argentina', pais: 'Argentina', nombre: 'Gastos Jurídicos',  match: /argentina/i },
];

const PAIS_POR_NOMBRE = [
  [/chile|santiago|\bcl\b/i, 'Chile'],
  [/colombia|bogot[áa]|\bco\b/i, 'Colombia'],
  [/per[úu]|lima|\bpe\b/i, 'Perú'],
  [/argentina|buenos aires|\bar\b/i, 'Argentina'],
];

// Devuelve el proveedor configurado o, si la carpeta es nueva, uno deducido de
// su nombre. Así un proveedor recién agregado entra al reporte sin tocar código:
// el país sale del nombre de la carpeta y se usan los lectores genéricos.
export function proveedorDeRuta(ruta) {
  // El proveedor lo define la carpeta raíz, no cualquier coincidencia en la ruta:
  // "Andes Latam - Perú/2026/6. Junio/Dentons perú/…" es de Andes Latam, no de Dentons.
  const carpeta = String(ruta).replace(/^\/+/, '').split('/')[0];
  const conocido = PROVEEDORES.find(p => p.match.test(carpeta));
  if (conocido) return conocido;
  if (!carpeta || carpeta.includes('.')) return null; // archivo suelto en la raíz
  const pais = (PAIS_POR_NOMBRE.find(([re]) => re.test(carpeta)) || [])[1] || null;
  // "Estudio Pérez - Chile" → nombre "Estudio Pérez"
  const nombre = carpeta.replace(/\s*[-–]\s*(chile|colombia|per[úu]|argentina)\s*$/i, '').trim() || carpeta;
  return { id: 'nuevo', pais: pais || 'Sin país', nombre, nuevo: true, paisDeducido: Boolean(pais) };
}

// Deduce mes/año desde la ruta ("/Mensual/3- Marzo/", "/2025/", "Sensus Legis/Julio/")
// como respaldo cuando el documento no trae fecha parseable.
export function fechaDeRuta(ruta) {
  let mes = null, anio = null;
  // Acepta "3- Marzo", "6. Junio" y variantes
  const mNum = ruta.match(/\/(\d{1,2})\s*[-.]\s*[A-Za-zÁÉÍÓÚáéíóúñ]+\//);
  if (mNum) mes = parseInt(mNum[1], 10);
  if (!mes) {
    const mNombre = ruta.toLowerCase().match(/\/(enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|setiembre|octubre|noviembre|diciembre)\//);
    if (mNombre) mes = MES_NOMBRES[mNombre[1]];
  }
  const mAnio = ruta.match(/\/(20\d{2})\//);
  if (mAnio) anio = parseInt(mAnio[1], 10);
  return { mes, anio };
}

export function parseFechaTexto(texto) {
  // "07 de Octubre de 2025", "11 DICIEMBRE 2025", "30 de enero de 2026"
  const m = texto.toLowerCase().match(/(\d{1,2})\s*(?:de\s+)?(enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|setiembre|octubre|noviembre|diciembre)\s*(?:de\s+)?(20\d{2})/);
  if (m) return { dia: +m[1], mes: MES_NOMBRES[m[2]], anio: +m[3] };
  // pdf.js puede fragmentar "03-12-2025" como "03 - 1 2 - 2025": compactamos dígitos y guiones
  const compacto = texto.replace(/(\d)\s+(?=[\d\-\/])/g, '$1').replace(/\s*([-\/])\s*/g, '$1');
  const iso = compacto.match(/(\d{1,2})[-\/](\d{1,2})[-\/](20\d{2})/);
  if (iso && +iso[2] >= 1 && +iso[2] <= 12) return { dia: +iso[1], mes: +iso[2], anio: +iso[3] };
  return null;
}

// "3.567.925" | "2.770.923" | "1.318,33" | "70,18" → número
export function numeroCL(s) {
  if (s == null) return NaN;
  s = String(s).trim().replace(/\$/g, '').replace(/\s/g, '');
  if (/,\d{1,2}$/.test(s)) return parseFloat(s.replace(/\./g, '').replace(',', '.'));
  return parseFloat(s.replace(/\./g, ''));
}

// "1318.33" estilo US o número Excel
export function numeroUS(s) {
  if (typeof s === 'number') return s;
  if (s == null) return NaN;
  return parseFloat(String(s).replace(/[^\d.\-]/g, ''));
}

// Extrae montos de un texto reconociendo los dos estilos de separadores:
// "3.034,55" (es) y "3,034.55" / "207.68" (us). Devuelve [{valor, indice}].
export function montosDelTexto(texto) {
  const out = [];
  const re = /\d{1,3}(?:[.,]\d{3})*(?:[.,]\d{1,2})?|\d+(?:[.,]\d{1,2})?/g;
  let m;
  while ((m = re.exec(texto))) {
    const s = m[0];
    let v = NaN;
    if (/^\d{1,3}(\.\d{3})+,\d{1,2}$/.test(s)) v = parseFloat(s.replace(/\./g, '').replace(',', '.'));      // 3.034,55
    else if (/^\d{1,3}(,\d{3})+\.\d{1,2}$/.test(s)) v = parseFloat(s.replace(/,/g, ''));                     // 3,034.55
    else if (/^\d{1,3}(\.\d{3})+$/.test(s)) v = parseFloat(s.replace(/\./g, ''));                            // 3.567.925
    else if (/^\d{1,3}(,\d{3})+$/.test(s)) v = parseFloat(s.replace(/,/g, ''));                              // 3,567,925
    else if (/^\d+,\d{1,2}$/.test(s)) v = parseFloat(s.replace(',', '.'));                                   // 1234,55
    else if (/^\d+(\.\d{1,2})?$/.test(s)) v = parseFloat(s);                                                 // 207.68 | 200
    if (isFinite(v) && v > 0) out.push({ valor: v, indice: m.index });
  }
  return out;
}

// Moneda declarada en el documento; si no hay marca, la del país.
export function monedaDelTexto(texto, paisPorDefecto) {
  if (/US\$|USD|D[ÓO]LARES/i.test(texto)) return 'USD';
  if (/\bS\/\b|SOLES/i.test(texto)) return 'PEN';
  const porPais = { Chile: 'CLP', Colombia: 'COP', 'Perú': 'PEN', Argentina: 'ARS' };
  return porPais[paisPorDefecto] || 'CLP';
}
