// Lectores por proveedor. Cada parser recibe el contenido del archivo y devuelve
// { registros: [...], estado: 'ok'|'supuesto'|'error', nota }
// Registro: { dia, mes, anio, pais, proveedor, categoria, concepto, moneda, montoOrigen, clp, archivo, detalle }

import { fechaDeRuta, parseFechaTexto, numeroCL, numeroUS, proveedorDeRuta, montosDelTexto, monedaDelTexto } from './config.js';

const XLSX = window.XLSX;
const pdfjsLib = window.pdfjsLib;

// ---------- utilidades ----------

async function textoPdf(arrayBuffer) {
  const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
  let out = '';
  for (let p = 1; p <= pdf.numPages; p++) {
    const page = await pdf.getPage(p);
    const tc = await page.getTextContent();
    out += tc.items.map(i => i.str).join(' ') + '\n';
  }
  return out;
}

// Números estilo chileno con miles: 3.567.925 (mínimo 4 dígitos reales).
// Excluye RUT: un número seguido de guion + dígito verificador ("77.645.526-1",
// "76.026.270−6", "77.955.200 - 4") no es un monto.
function numerosCLP(texto) {
  const out = [];
  const re = /(?:\$\s*)?(\d{1,3}(?:\.\d{3}){1,4})(?!\d)(?!\s*[-−–]\s*[\dkK])/g;
  let m;
  while ((m = re.exec(texto))) {
    const v = numeroCL(m[1]);
    if (v >= 1000 && v !== 20009) out.push(v); // 20.009 = N° de ley, no monto
  }
  return out;
}

// Solicitantes mencionados en un documento ("Solicitado por X", "Solicitante: X").
// Si hay uno solo se puede atribuir; si hay varios se marca como tal y se listan.
function solicitantesDelTexto(texto) {
  const nombres = new Set();
  const re = /Solicitad[oa]\s+por\s+([A-ZÁÉÍÓÚÑ][\wÁÉÍÓÚÑáéíóúñ.'-]*(?:\s+[A-ZÁÉÍÓÚÑ][\wÁÉÍÓÚÑáéíóúñ.'-]*){0,3})|Solicitante\s*:?\s*([A-ZÁÉÍÓÚÑ][\wÁÉÍÓÚÑáéíóúñ.'-]*(?:\s+[A-ZÁÉÍÓÚÑ][\wÁÉÍÓÚÑáéíóúñ.'-]*){0,3})/g;
  let m;
  while ((m = re.exec(texto))) {
    const n = (m[1] || m[2] || '').trim().replace(/\s+/g, ' ');
    if (n && !/^sin\b/i.test(n) && n.length > 3) nombres.add(n);
  }
  return [...nombres];
}

// Solicitantes que aparecen en las hojas de detalle de un libro Excel
function solicitantesDeHojas(wb) {
  const nombres = new Set();
  for (const hoja of wb.SheetNames) {
    const filas = filasDeHoja(wb.Sheets[hoja]);
    const hIdx = filas.findIndex(r => r && r.some(c => /solicitante/i.test(String(c ?? ''))));
    if (hIdx < 0) continue;
    const col = filas[hIdx].findIndex(c => /solicitante/i.test(String(c ?? '')));
    for (const r of filas.slice(hIdx + 1)) {
      const v = String(r[col] ?? '').trim();
      if (v && v.length > 3 && !/^total/i.test(v)) nombres.add(v);
    }
  }
  return [...nombres];
}

function solicitanteResumen(lista) {
  if (!lista.length) return '';
  if (lista.length === 1) return lista[0];
  return `Varios (${lista.length})`;
}

function fechaDoc(texto, ruta, etiqueta) {
  if (etiqueta) {
    const zona = texto.match(new RegExp(etiqueta + '[\\s\\S]{0,60}', 'i'));
    if (zona) {
      const f = parseFechaTexto(zona[0]);
      if (f) return { ...f, origen: 'doc' };
    }
  }
  const f = parseFechaTexto(texto);
  if (f) return { ...f, origen: 'doc' };
  const r = fechaDeRuta(ruta);
  if (r.mes) return { dia: 1, mes: r.mes, anio: r.anio, origen: 'ruta' };
  return null;
}

function filasDeHoja(ws) {
  return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
}

function buscaColumna(headers, ...patrones) {
  for (const pat of patrones) {
    const i = headers.findIndex(h => h != null && pat.test(String(h)));
    if (i >= 0) return i;
  }
  return -1;
}

// ---------- parsers PDF (Chile) ----------

// Boleta de honorarios Moraga: usa el monto NETO (bruto - retención), igual que el Excel.
export async function parseMoragaBH(buf, ruta, params, archivo) {
  const texto = await textoPdf(buf);
  // Preferir la zona "Total Honorarios … Total:" donde viven bruto/retención/neto
  const zona = texto.match(/Total\s+Honorarios[\s\S]{0,200}/i);
  let nums = zona ? [...new Set(numerosCLP(zona[0]))].sort((a, b) => b - a) : [];
  if (nums.length < 2) nums = [...new Set(numerosCLP(texto))].sort((a, b) => b - a);
  if (!nums.length) return { registros: [], estado: 'error', nota: 'PDF escaneado sin texto — registrar a mano o re-subir el PDF electrónico' };
  const bruto = nums[0];
  const retencion = nums.find(n => n >= bruto * 0.10 && n <= bruto * 0.17);
  const neto = retencion ? (nums.find(n => Math.abs(n - (bruto - retencion)) < 2) ?? bruto - retencion) : bruto;
  // La carpeta del mes (mes del servicio) manda sobre la fecha de emisión de la boleta
  const rf = fechaDeRuta(ruta);
  const f = rf.mes ? { dia: 1, mes: rf.mes, anio: rf.anio, origen: 'ruta' } : fechaDoc(texto, ruta, 'Fecha');
  const mDesc = texto.match(/Por atenci[oó]n profesional:?\s*([^\n]{0,80}?)\s*Total/i);
  return {
    registros: [{
      ...fechaReg(f), pais: 'Chile', proveedor: 'Alvaro Moraga',
      categoria: 'Asesoría Profesional', concepto: 'Gastos', moneda: 'CLP',
      solicitante: solicitanteResumen(solicitantesDelTexto(texto)),
      montoOrigen: neto, clp: neto, archivo, detalle: mDesc ? mDesc[1].trim() : 'Boleta de honorarios',
    }],
    estado: retencion ? 'ok' : 'supuesto',
    nota: retencion ? `Bruto ${fmt(bruto)} − retención ${fmt(retencion)} = neto ${fmt(neto)}` : 'No se identificó retención; se usó el monto mayor',
  };
}

// Nota de cobro Moraga (NCCL): Honorario Bruto en UF.
export async function parseMoragaNCCL(buf, ruta, params, archivo) {
  const texto = await textoPdf(buf);
  let uf = null;
  const mBruto = texto.match(/Honorario\s+Bruto[\s\S]{0,60}?UF\s*([\d.,]+)/i);
  if (mBruto) uf = numeroCL(mBruto[1]);
  if (!uf) {
    const todos = [...texto.matchAll(/UF\s*([\d]{1,3}(?:[.,]\d{1,2})?)/g)].map(m => numeroCL(m[1])).filter(v => v > 0);
    if (todos.length) uf = Math.max(...todos);
  }
  const solicitantesAnexo = solicitantesDelTexto(texto);
  if (!uf) return { registros: [], estado: 'omitido', solicitantesAnexo,
    nota: 'Nota de cobro sin monto (solo detalle de horas) — el cobro lo aporta la boleta del mes'
      + (solicitantesAnexo.length ? `; aporta ${solicitantesAnexo.length} solicitante(s)` : '') };
  // La carpeta del mes manda: las descripciones internas traen fechas de otros meses.
  // El año, si la ruta no lo trae, sale del período de facturación del documento.
  const rf = fechaDeRuta(ruta);
  const anioTexto = (texto.match(/hasta el[^\n]{0,40}?(20\d{2})/i) || texto.match(/(20\d{2})/) || [])[1];
  const f = rf.mes
    ? { dia: 1, mes: rf.mes, anio: rf.anio || (anioTexto ? +anioTexto : null), origen: 'ruta' }
    : fechaDoc(texto, ruta, 'hasta el');
  const clp = Math.round(uf * params.UF_CLP);
  return {
    registros: [{
      ...fechaReg(f), pais: 'Chile', proveedor: 'Alvaro Moraga',
      categoria: 'Asesoría Profesional', concepto: 'Gastos', moneda: 'UF',
      solicitante: solicitanteResumen(solicitantesDelTexto(texto)),
      montoOrigen: uf, clp, archivo, detalle: `Nota de cobro ${uf} UF`,
    }],
    solicitantesAnexo,
    estado: 'supuesto',
    nota: `UF ${uf} × $${fmt(params.UF_CLP)} = ${fmt(clp)} (valor UF configurable)`,
  };
}

// Facturas exentas electrónicas (Dentons, Aninat): Monto Total + fecha de emisión.
export async function parseFacturaExenta(buf, ruta, params, archivo, proveedor) {
  const texto = await textoPdf(buf);
  // Preferir el monto junto a la etiqueta "Monto Total"; si no, el mayor del documento
  const zonaTotal = texto.match(/Monto\s+Total[\s\S]{0,80}/i);
  let nums = zonaTotal ? numerosCLP(zonaTotal[0]) : [];
  if (!nums.length) nums = numerosCLP(texto);
  if (!nums.length) return { registros: [], estado: 'error', nota: 'PDF escaneado sin texto — registrar a mano o re-subir el PDF electrónico' };
  const monto = Math.max(...nums);
  const f = fechaDoc(texto, ruta, 'Fecha\\s*Emis');
  const mDesc = texto.match(/(Honorarios?[^\n]{0,90}|Asesor[ií]a[^\n]{0,90})/i);
  const esGasto = /NRG|nota\s+de\s+gasto/i.test(archivo + ' ' + texto.slice(0, 400));
  return {
    registros: [{
      ...fechaReg(f), pais: proveedor.pais, proveedor: proveedor.nombre,
      categoria: esGasto ? 'Gastos notariales y otros' : 'Honorarios varios',
      concepto: esGasto ? 'Gastos' : 'Honorarios', moneda: 'CLP',
      solicitante: solicitanteResumen(solicitantesDelTexto(texto)),
      montoOrigen: monto, clp: monto, archivo, detalle: mDesc ? mDesc[1].trim() : 'Factura',
    }],
    estado: 'ok', nota: `Monto total ${fmt(monto)}`,
  };
}

// Nota de cobro Sensus / Yáñez & Acuña: "Total: UF 74" + IVA.
export async function parseSensusNC(buf, ruta, params, archivo) {
  const texto = await textoPdf(buf);
  const m = texto.match(/Total\s*:?\s*UF\s*([\d.,]+)/i);
  if (!m) return { registros: [], estado: 'error', nota: 'No se encontró "Total: UF" en la nota de cobro' };
  const uf = numeroCL(m[1]);
  const f = fechaDoc(texto, ruta, 'Fecha');
  const neto = uf * params.UF_CLP;
  const clp = Math.round(neto * (1 + params.IVA_CL));
  return {
    registros: [{
      ...fechaReg(f), pais: 'Chile', proveedor: 'Sensus Legis',
      categoria: 'Honorarios Ley 20.009', concepto: 'Honorarios', moneda: 'UF',
      solicitante: solicitanteResumen(solicitantesDelTexto(texto)),
      montoOrigen: uf, clp, archivo, detalle: `Hitos Ley 20.009: ${uf} UF`,
    }],
    estado: 'supuesto',
    nota: `UF ${uf} × $${fmt(params.UF_CLP)} + IVA ${params.IVA_CL * 100}% = ${fmt(clp)}`,
  };
}

// Factura genérica en cualquier moneda (SUNAT Perú, facturas argentinas, etc.).
// Toma el monto de la línea "Total" (ignorando subtotales) y, si no la encuentra,
// el mayor del documento. La moneda sale del propio documento.
export async function parseFacturaGenerica(buf, ruta, params, archivo, proveedor) {
  const texto = await textoPdf(buf);
  const montos = montosDelTexto(texto);
  if (!montos.length) return { registros: [], estado: 'error', nota: 'PDF escaneado sin texto — registrar a mano o re-subir el PDF electrónico' };

  // Preferir los montos que aparecen cerca de un "Total" que no sea subtotal
  let candidatos = [];
  const reTotal = /(?<!sub)\bT\s?O\s?T\s?A\s?L\b|Importe\s+total/gi;
  let mt;
  while ((mt = reTotal.exec(texto))) {
    const zona = texto.slice(mt.index, mt.index + 120);
    if (/valor\s+venta|detracci/i.test(zona)) continue;
    candidatos.push(...montosDelTexto(zona).map(x => x.valor));
  }
  if (!candidatos.length) candidatos = montos.map(x => x.valor);
  const monto = Math.max(...candidatos);

  const moneda = monedaDelTexto(texto, proveedor.pais);
  const tasas = { USD: params.USD_CLP, PEN: params.PEN_CLP, COP: params.COP_CLP, ARS: params.ARS_CLP, CLP: 1 };
  const clp = Math.round(monto * (tasas[moneda] ?? 1));

  const rf = fechaDeRuta(ruta);
  const f = parseFechaTexto(texto) || (rf.mes ? { dia: 1, mes: rf.mes, anio: rf.anio } : null);
  const esHonorario = /honorario|fee|asesor/i.test(archivo + ' ' + texto.slice(0, 600));
  return {
    registros: [{
      ...fechaReg(f ? { ...f, origen: 'doc' } : null),
      pais: proveedor.pais, proveedor: proveedor.nombre,
      categoria: esHonorario ? 'Honorarios legales' : 'Gastos legales',
      concepto: esHonorario ? 'Honorarios' : 'Gastos',
      solicitante: solicitanteResumen(solicitantesDelTexto(texto)),
      moneda, montoOrigen: monto, clp, archivo,
      detalle: archivo.replace(/\.[a-z]+$/i, ''),
    }],
    estado: 'supuesto',
    nota: `${moneda} ${monto.toLocaleString('es-CL')} × ${tasas[moneda] ?? 1} = ${fmt(clp)} — verificar monto y moneda`,
  };
}

// ---------- parsers XLSX ----------

// Minuta Andes Latam. Fuente principal: el bloque "Resumen cobro" de la hoja
// "1 GLOBAL 66" → "Total cobro" en USD (ya incluye IGV), que es el monto facturado.
// Sirve para los dos formatos de minuta (con y sin hojas "ok").
export function parseAndesLatam(buf, ruta, params, archivo) {
  const wb = XLSX.read(buf, { type: 'array', cellDates: true });
  const hoja = wb.SheetNames.find(n => /global\s*66/i.test(n)) || wb.SheetNames[0];
  if (hoja) {
    const filas = filasDeHoja(wb.Sheets[hoja]);
    const valorTrasEtiqueta = (fila, i) => {
      for (let j = i + 1; j < fila.length; j++) {
        const v = numeroUS(fila[j]);
        if (isFinite(v) && v > 0) return v;
      }
      return NaN;
    };
    const fechaTrasEtiqueta = (fila, i) => {
      for (let j = i + 1; j < fila.length; j++) {
        const c = fila[j];
        if (c instanceof Date && !isNaN(c)) return { dia: c.getDate(), mes: c.getMonth() + 1, anio: c.getFullYear() };
        const m = String(c ?? '').match(/(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
        if (m) return { dia: +m[1], mes: +m[2], anio: +m[3] < 100 ? 2000 + +m[3] : +m[3] };
      }
      return null;
    };
    let total = NaN, hasta = null, desde = null, factura = '';
    for (const fila of filas) {
      if (!fila) continue;
      for (let i = 0; i < fila.length; i++) {
        const c = String(fila[i] ?? '');
        if (/total\s*cobro/i.test(c)) total = valorTrasEtiqueta(fila, i);
        else if (/fecha\s*hasta/i.test(c)) hasta = fechaTrasEtiqueta(fila, i);
        else if (/fecha\s*desde/i.test(c)) desde = fechaTrasEtiqueta(fila, i);
        else if (/factura\s*n/i.test(c)) {
          const v = fila.slice(i + 1).find(x => x != null && String(x).trim());
          if (v) factura = String(v).trim();
        }
      }
    }
    if (isFinite(total) && total > 0) {
      const solicitantes = solicitantesDeHojas(wb);
      const rf = fechaDeRuta(ruta);
      const f = hasta || desde;
      const mes = (f && f.mes) || rf.mes || null;
      const anio = (f && f.anio) || rf.anio || null;
      const clp = Math.round(total * params.USD_CLP);
      return {
        registros: [{
          dia: (f && f.dia) || 1, mes, anio, pais: 'Perú', proveedor: 'Andes Latam',
          categoria: 'Asesoría legal Andes Latam', concepto: 'Honorarios', moneda: 'USD',
          solicitante: solicitanteResumen(solicitantes),
          montoOrigen: total, clp, archivo,
          detalle: `Minuta de liquidación${factura ? ' ' + factura : ''} (honorarios + gastos + IGV)`,
        }],
        estado: 'supuesto',
        nota: `Total cobro USD ${total.toFixed(2)} × ${params.USD_CLP} = ${fmt(clp)} (incluye IGV)`,
      };
    }
  }
  return parseAndesPorLineas(buf, ruta, params, archivo);
}

// Respaldo: suma línea a línea de las hojas "ok" (formato antiguo sin resumen de cobro).
function parseAndesPorLineas(buf, ruta, params, archivo) {
  const wb = XLSX.read(buf, { type: 'array' });
  const registros = [];
  const notas = [];
  for (const nombre of wb.SheetNames) {
    if (!/ok/i.test(nombre)) continue;
    const filas = filasDeHoja(wb.Sheets[nombre]);
    if (!filas.length) continue;
    const hIdx = filas.findIndex(r => r && r.some(c => /Descripci[oó]n/i.test(String(c ?? ''))));
    if (hIdx < 0) continue;
    const H = filas[hIdx].map(h => h == null ? '' : String(h));
    const cValorUSD = buscaColumna(H, /Valor\s*\(USD\)/i);
    const cTotal = buscaColumna(H, /^Total$/i);
    const cCat = buscaColumna(H, /Categor/i);
    const cDesc = buscaColumna(H, /Descripci/i);
    const cSolic = buscaColumna(H, /Solicitante/i);
    if (cValorUSD >= 0) {
      // Timesheet de abogados en USD (columnas Día/Mes/Año)
      const cDia = buscaColumna(H, /^D[ií]a$/i), cMes = buscaColumna(H, /^Mes$/i), cAnio = buscaColumna(H, /^A[ñn]o$/i);
      for (const r of filas.slice(hIdx + 1)) {
        const usd = numeroUS(r[cValorUSD]);
        if (!isFinite(usd) || usd <= 0) continue;
        const clp = Math.round(usd * params.USD_CLP * (1 + params.IVA_PE));
        registros.push({
          dia: numeroUS(r[cDia]) || 1, mes: numeroUS(r[cMes]) || null, anio: numeroUS(r[cAnio]) || null,
          pais: 'Perú', proveedor: 'Andes Latam',
          categoria: (cCat >= 0 && r[cCat]) ? String(r[cCat]) : 'Asesoría legal por horas',
          solicitante: (cSolic >= 0 && r[cSolic]) ? String(r[cSolic]).trim() : '',
          concepto: 'Gastos', moneda: 'USD', montoOrigen: usd, clp, archivo,
          detalle: cDesc >= 0 ? String(r[cDesc] ?? '').slice(0, 120) : '',
        });
      }
    } else if (cTotal >= 0) {
      // Gastos reembolsables en soles (columna Fecha dd/mm/yy)
      const cFecha = buscaColumna(H, /^Fecha$/i);
      for (const r of filas.slice(hIdx + 1)) {
        const pen = numeroUS(r[cTotal]);
        if (!isFinite(pen) || pen <= 0) continue;
        let dia = 1, mes = null, anio = null;
        const mF = String(r[cFecha] ?? '').match(/(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
        if (mF) { dia = +mF[1]; mes = +mF[2]; anio = +mF[3] < 100 ? 2000 + +mF[3] : +mF[3]; }
        const clp = Math.round(pen * params.PEN_CLP * (1 + params.IVA_PE));
        registros.push({
          dia, mes, anio, pais: 'Perú', proveedor: 'Andes Latam',
          categoria: (cCat >= 0 && r[cCat]) ? String(r[cCat]) : 'Gastos reembolsables',
          solicitante: (cSolic >= 0 && r[cSolic]) ? String(r[cSolic]).trim() : '',
          concepto: 'Gastos', moneda: 'PEN', montoOrigen: pen, clp, archivo,
          detalle: cDesc >= 0 ? String(r[cDesc] ?? '').slice(0, 120) : '',
        });
      }
    }
  }
  if (!registros.length) return { registros, estado: 'error', nota: 'No se encontraron hojas "ok" con datos' };
  // Fallback de fecha por ruta para filas sin mes/año
  const rf = fechaDeRuta(ruta);
  for (const reg of registros) { if (!reg.mes && rf.mes) { reg.mes = rf.mes; reg.anio = reg.anio || rf.anio; } }
  notas.push(`${registros.length} filas (USD→CLP ${params.USD_CLP}, PEN→CLP ${params.PEN_CLP}, +IGV ${params.IVA_PE * 100}%)`);
  return { registros, estado: 'supuesto', nota: notas.join(' · ') };
}

// Planillas tabulares (Colombia, Argentina, Sensus listado, genérico):
// busca una fila de encabezado con Fecha + (Valor|Total|Monto).
export function parseTabularGenerico(buf, ruta, params, archivo, proveedor) {
  const wb = XLSX.read(buf, { type: 'array', cellDates: true });
  const registros = [];
  const monedaPais = { Colombia: ['COP', params.COP_CLP], Argentina: ['ARS', params.ARS_CLP], Chile: ['CLP', 1], 'Perú': ['PEN', params.PEN_CLP] };
  const [moneda, tasa] = monedaPais[proveedor.pais] || ['CLP', 1];
  for (const nombre of wb.SheetNames) {
    const filas = filasDeHoja(wb.Sheets[nombre]);
    const hIdx = filas.findIndex(r => r && r.some(c => /fecha/i.test(String(c ?? ''))) && r.some(c => /(valor|total|monto)/i.test(String(c ?? ''))));
    if (hIdx < 0) continue;
    const H = filas[hIdx].map(h => h == null ? '' : String(h));
    const cFecha = buscaColumna(H, /fecha/i);
    const cMonto = buscaColumna(H, /^total\b/i, /valor/i, /monto/i);
    const cCat = buscaColumna(H, /categor/i);
    const cDesc = buscaColumna(H, /descripci/i, /detalle/i, /materia/i);
    const cSolic = buscaColumna(H, /solicitante/i, /solicitado/i);
    for (const r of filas.slice(hIdx + 1)) {
      const v = numeroUS(r[cMonto]);
      if (!isFinite(v) || v <= 0) continue;
      let dia = 1, mes = null, anio = null;
      const cell = r[cFecha];
      if (cell instanceof Date) { dia = cell.getDate(); mes = cell.getMonth() + 1; anio = cell.getFullYear(); }
      else if (cell != null) {
        const f = parseFechaTexto(String(cell)) || (String(cell).match(/(\d{1,2})\/(\d{1,2})\/(\d{2,4})/) && { dia: +RegExp.$1, mes: +RegExp.$2, anio: +RegExp.$3 < 100 ? 2000 + +RegExp.$3 : +RegExp.$3 });
        if (f) ({ dia, mes, anio } = f);
      }
      registros.push({
        dia, mes, anio, pais: proveedor.pais, proveedor: proveedor.nombre,
        categoria: (cCat >= 0 && r[cCat]) ? String(r[cCat]) : 'Gastos legales',
        solicitante: (cSolic >= 0 && r[cSolic]) ? String(r[cSolic]).trim() : '',
        concepto: /honorario/i.test(String(r[cCat] ?? '') + nombre) ? 'Honorarios' : 'Gastos',
        moneda, montoOrigen: v, clp: Math.round(v * tasa), archivo,
        detalle: cDesc >= 0 ? String(r[cDesc] ?? '').slice(0, 120) : '',
      });
    }
  }
  if (!registros.length) return { registros, estado: 'error', nota: 'No se reconoció una tabla con columnas Fecha + Monto' };
  const rf = fechaDeRuta(ruta);
  for (const reg of registros) { if (!reg.mes && rf.mes) { reg.mes = rf.mes; reg.anio = reg.anio || rf.anio; } }
  return { registros, estado: 'ok', nota: `${registros.length} filas (${moneda}→CLP ${tasa})` };
}

// Base histórica: hoja "Base resumen" del Consolidado Paises.xlsx.
// Replica la lógica del Excel: esta base plana alimenta las tablas resumen.
// Los valores ya reportados ahí se respetan tal cual (no se recalculan).
export function parseBaseResumen(buf, ruta, params, archivo) {
  const wb = XLSX.read(buf, { type: 'array', cellDates: true });
  const hoja = wb.SheetNames.find(n => /base\s*resumen/i.test(n));
  if (!hoja) return { registros: [], estado: 'error', nota: 'No se encontró la hoja "Base resumen"' };
  const filas = filasDeHoja(wb.Sheets[hoja]);
  const H = (filas[0] || []).map(h => String(h ?? ''));
  // Solo el bloque principal (columnas A–H); a la derecha hay bloques auxiliares duplicados
  const col = {
    fecha: H.findIndex(h => /^Fecha$/i.test(h)),
    categoria: H.findIndex(h => /^Categoria$/i.test(h)),
    mes: H.findIndex(h => /^Mes$/i.test(h)),
    pais: H.findIndex(h => /^Pais$/i.test(h)),
    clp: H.findIndex(h => /^CLP IVA$/i.test(h)),
    concepto: H.findIndex(h => /^Concepto$/i.test(h)),
  };
  if (col.pais < 0 || col.clp < 0) return { registros: [], estado: 'error', nota: 'La hoja "Base resumen" no tiene las columnas esperadas (Pais, CLP IVA)' };
  const registros = [];
  for (const r of filas.slice(1)) {
    const clp = numeroUS(r[col.clp]);
    const pais = r[col.pais];
    if (!pais || !isFinite(clp) || clp <= 0) continue;
    const f = r[col.fecha];
    const esFecha = f instanceof Date && !isNaN(f);
    const conceptoRaw = String(r[col.concepto] ?? '').toLowerCase();
    const concepto = conceptoRaw.startsWith('juicio') ? 'Juicios y otros'
      : conceptoRaw.startsWith('honorario') ? 'Honorarios' : 'Gastos';
    registros.push({
      dia: esFecha ? f.getDate() : 1,
      mes: numeroUS(r[col.mes]) || (esFecha ? f.getMonth() + 1 : null),
      anio: esFecha ? f.getFullYear() : null,
      pais: String(pais).trim(), proveedor: 'Reportado (base histórica)',
      categoria: String(r[col.categoria] ?? 'Sin categoría').trim(),
      concepto, moneda: 'CLP', montoOrigen: clp, clp: Math.round(clp),
      archivo, detalle: '', fuente: 'base',
    });
  }
  return { registros, estado: 'ok', nota: `${registros.length} movimientos ya reportados — se mantienen tal cual`, fuente: 'base' };
}

// Lo ya reportado manda: si la base histórica tiene datos para un país+mes,
// los archivos de carpetas no vuelven a sumar ese país+mes (evita doble conteo).
export function aplicarBaseHistorica(resultados) {
  const cubiertos = new Set();
  for (const res of resultados) {
    if (res.fuente !== 'base') continue;
    for (const r of res.registros) if (r.anio && r.mes) cubiertos.add(`${r.pais}|${r.anio}-${r.mes}`);
  }
  if (!cubiertos.size) return resultados;
  for (const res of resultados) {
    if (res.fuente === 'base' || !res.registros.length) continue;
    const antes = res.registros.length;
    res.registros = res.registros.filter(r => !cubiertos.has(`${r.pais}|${r.anio}-${r.mes}`));
    const quitados = antes - res.registros.length;
    if (quitados && !res.registros.length) {
      res.estado = 'omitido';
      res.nota = 'Período ya reportado en la base histórica del Excel — se mantiene el valor reportado';
    } else if (quitados) {
      res.nota = (res.nota || '') + ` · ${quitados} filas omitidas por estar ya reportadas en la base histórica`;
    }
  }
  return resultados;
}

// Plantilla de carga manual: la genera la app y cualquiera puede llenarla a mano.
// Se reconoce por sus encabezados, así que basta subirla a la carpeta del proveedor.
export function parseCargaManual(buf, ruta, params, archivo) {
  const wb = XLSX.read(buf, { type: 'array', cellDates: true });
  const registros = [];
  for (const hoja of wb.SheetNames) {
    const filas = filasDeHoja(wb.Sheets[hoja]);
    const hIdx = filas.findIndex(r => r && r.some(c => /^solicitante$/i.test(String(c ?? '').trim()))
      && r.some(c => /monto/i.test(String(c ?? ''))));
    if (hIdx < 0) continue;
    const H = filas[hIdx].map(h => String(h ?? '').trim());
    const col = {
      fecha: buscaColumna(H, /^fecha$/i), solicitante: buscaColumna(H, /^solicitante$/i),
      proveedor: buscaColumna(H, /^proveedor$/i), pais: buscaColumna(H, /^pa[ií]s$/i),
      concepto: buscaColumna(H, /^concepto$/i), categoria: buscaColumna(H, /^categor/i),
      moneda: buscaColumna(H, /^moneda$/i), origen: buscaColumna(H, /monto\s*origen/i),
      clp: buscaColumna(H, /monto\s*clp/i), usd: buscaColumna(H, /monto\s*usd/i),
      detalle: buscaColumna(H, /^detalle$/i),
    };
    for (const r of filas.slice(hIdx + 1)) {
      const clp = numeroUS(r[col.clp]);
      if (!isFinite(clp) || clp <= 0) continue;
      let dia = 1, mes = null, anio = null;
      const cel = r[col.fecha];
      if (cel instanceof Date && !isNaN(cel)) { dia = cel.getDate(); mes = cel.getMonth() + 1; anio = cel.getFullYear(); }
      else {
        const m = String(cel ?? '').match(/(\d{1,2})[\/-](\d{1,2})[\/-](\d{2,4})/);
        if (m) { dia = +m[1]; mes = +m[2]; anio = +m[3] < 100 ? 2000 + +m[3] : +m[3]; }
      }
      const rf = fechaDeRuta(ruta);
      registros.push({
        dia, mes: mes || rf.mes, anio: anio || rf.anio,
        pais: String(r[col.pais] ?? '').trim() || 'Sin país',
        proveedor: String(r[col.proveedor] ?? '').trim() || 'Sin proveedor',
        solicitante: String(r[col.solicitante] ?? '').trim(),
        categoria: String(r[col.categoria] ?? '').trim() || 'Carga manual',
        concepto: String(r[col.concepto] ?? '').trim() || 'Gastos',
        moneda: String(r[col.moneda] ?? 'CLP').trim(),
        montoOrigen: numeroUS(r[col.origen]) || clp,
        clp: Math.round(clp), usd: numeroUS(r[col.usd]) || undefined,
        archivo, detalle: String(r[col.detalle] ?? '').trim(), fuente: 'plantilla',
      });
    }
  }
  if (!registros.length) return { registros: [], estado: 'error', nota: 'Plantilla de carga manual sin filas con Monto CLP' };
  return { registros, estado: 'ok', nota: `${registros.length} movimiento(s) digitado(s) en plantilla` };
}

// ---------- despachador ----------

export async function procesarArchivo({ nombre, ruta, arrayBuffer }, params) {
  const base = { archivo: nombre, ruta };
  if (/anulad/i.test(nombre)) return { ...base, estado: 'omitido', nota: 'Documento ANULADO — excluido', registros: [] };
  if (/consolidado\s+paises\.xlsx$/i.test(nombre)) return { ...base, ...parseBaseResumen(arrayBuffer, ruta, params, nombre) };
  if (/\.xlsx?$/i.test(nombre) && /carga\s*manual/i.test(nombre)) {
    return { ...base, ...parseCargaManual(arrayBuffer, ruta, params, nombre) };
  }
  const prov = proveedorDeRuta(ruta);
  // Nota de cobro en UF: es de Moraga (Chile) aunque esté archivada en otra carpeta.
  if (/\.pdf$/i.test(nombre) && /^ncc/i.test(nombre)) {
    const r = await parseMoragaNCCL(arrayBuffer, ruta, params, nombre);
    if (!r.registros.length) return { ...base, ...r, proveedor: 'Alvaro Moraga', pais: 'Chile' };
    if (r.registros.length) {
      const fuera = !prov || prov.id !== 'moraga';
      if (fuera) r.nota += ' · ⚠ Documento de Moraga (Chile) archivado fuera de su carpeta';
      return { ...base, ...r, proveedor: 'Alvaro Moraga', pais: 'Chile' };
    }
  }
  if (!prov) return { ...base, estado: 'error', nota: 'Archivo fuera de una carpeta de proveedor', registros: [] };
  const ext = nombre.split('.').pop().toLowerCase();
  // Proveedor nuevo: se procesa igual con los lectores genéricos y se avisa.
  const avisoNuevo = prov.nuevo
    ? ` · 🆕 Proveedor nuevo detectado ("${prov.nombre}")` + (prov.paisDeducido ? `, país deducido: ${prov.pais}` : ' — agrega el país al nombre de la carpeta (ej. "Estudio X - Chile")')
    : '';
  try {
    let r;
    // Anexos de horas sin montos (HH…, "detalle cobro"): el monto lo aporta la boleta del mes
    if (ext === 'pdf' && (/^hh\b/i.test(nombre) || /detalle\s+de?\s*cobro/i.test(nombre)) && !/^bh/i.test(nombre)) {
      // No trae montos, pero sí los solicitantes del mes: se guardan para enriquecer
      // los cobros de la misma carpeta.
      let solicitantesAnexo = [];
      try { solicitantesAnexo = solicitantesDelTexto(await textoPdf(arrayBuffer)); } catch { /* anexo ilegible */ }
      return {
        ...base, estado: 'omitido', registros: [], solicitantesAnexo,
        nota: 'Anexo de detalle de horas (sin montos) — el cobro lo aporta la boleta o factura del mes'
          + (solicitantesAnexo.length ? `; aporta ${solicitantesAnexo.length} solicitante(s)` : ''),
      };
    }
    if (prov.id === 'moraga' && ext === 'pdf') {
      r = /^bh/i.test(nombre) ? await parseMoragaBH(arrayBuffer, ruta, params, nombre)
        : /nccl?/i.test(nombre) ? await parseMoragaNCCL(arrayBuffer, ruta, params, nombre)
        : await parseFacturaExenta(arrayBuffer, ruta, params, nombre, prov);
    } else if (prov.id === 'sensus' && ext === 'pdf' && /^nc/i.test(nombre)) {
      r = await parseSensusNC(arrayBuffer, ruta, params, nombre);
    } else if (prov.id === 'sensus') {
      return { ...base, estado: 'omitido', nota: 'Documento de respaldo — el monto lo aporta la nota de cobro (NC) del mes', registros: [] };
    } else if ((prov.id === 'dentons' || prov.id === 'aninat') && ext === 'pdf') {
      r = await parseFacturaExenta(arrayBuffer, ruta, params, nombre, prov);
    } else if (prov.id === 'andes' && ext === 'xlsx') {
      r = parseAndesLatam(arrayBuffer, ruta, params, nombre);
    } else if (ext === 'xlsx' || ext === 'xlsm') {
      r = parseTabularGenerico(arrayBuffer, ruta, params, nombre, prov);
    } else if (ext === 'pdf') {
      // Facturas en moneda extranjera (SUNAT Perú, facturas argentinas) o formato desconocido
      r = await parseFacturaGenerica(arrayBuffer, ruta, params, nombre, prov);
    } else {
      return { ...base, estado: 'omitido', nota: `Extensión .${ext} no soportada`, registros: [] };
    }
    if (avisoNuevo) {
      r.nota = (r.nota || '') + avisoNuevo;
      if (r.estado === 'ok') r.estado = 'supuesto';
    }
    return { ...base, ...r, proveedor: prov.nombre, pais: prov.pais };
  } catch (e) {
    return { ...base, estado: 'error', nota: 'Error al leer: ' + (e.message || e) + avisoNuevo, registros: [] };
  }
}

// Los anexos de detalle traen los solicitantes del mes pero no los montos; las
// boletas traen el monto pero no el solicitante. Se cruzan por carpeta (= mes).
export function enriquecerSolicitantes(resultados) {
  const porCarpeta = new Map();
  const carpetaDe = res => String(res.ruta || '').split('/').slice(0, -1).join('/');
  for (const res of resultados) {
    if (!res.solicitantesAnexo || !res.solicitantesAnexo.length) continue;
    const set = porCarpeta.get(carpetaDe(res)) || new Set();
    res.solicitantesAnexo.forEach(s => set.add(s));
    porCarpeta.set(carpetaDe(res), set);
  }
  if (!porCarpeta.size) return resultados;
  for (const res of resultados) {
    if (!res.registros || !res.registros.length) continue;
    const lista = [...(porCarpeta.get(carpetaDe(res)) || [])];
    if (!lista.length) continue;
    let n = 0;
    for (const r of res.registros) {
      if (r.solicitante) continue;
      r.solicitante = solicitanteResumen(lista);
      r.solicitantesLista = lista;
      n++;
    }
    if (n) res.nota = (res.nota || '') + ` · solicitante(s) según el anexo del mes: ${lista.join(', ')}`;
  }
  return resultados;
}

// Deduplicación Moraga: si un mes tiene boleta (BH) y nota de cobro (NCCL), la boleta manda.
export function deduplicar(resultados) {
  const bhMeses = new Set();
  for (const res of resultados) {
    if (res.estado === 'error' || res.estado === 'omitido') continue;
    if (/^bh/i.test(res.archivo) && res.registros.length) {
      for (const r of res.registros) bhMeses.add(`${r.anio}-${r.mes}`);
    }
  }
  for (const res of resultados) {
    if (/nccl?/i.test(res.archivo) && res.registros.length) {
      const r0 = res.registros[0];
      if (bhMeses.has(`${r0.anio}-${r0.mes}`)) {
        res.estado = 'omitido';
        res.nota = 'Mes ya cubierto por boleta de honorarios (BH) — se omite para no duplicar';
        res.registros = [];
      }
    }
  }
  // Copias exactas: mismo proveedor+monto+mes en archivos distintos con nombre tipo "(1)" o "Copia de"
  const vistos = new Map();
  for (const res of resultados) {
    for (const r of [...res.registros]) {
      const clave = `${r.proveedor}|${r.anio}-${r.mes}|${r.clp}|${(r.detalle || '').slice(0, 40)}`;
      const previo = vistos.get(clave);
      if (previo && /\(\d\)|copia/i.test(res.archivo)) {
        res.registros = res.registros.filter(x => x !== r);
        res.estado = 'omitido';
        res.nota = `Duplicado de ${previo} — se omite`;
      } else if (!previo) vistos.set(clave, res.archivo);
    }
  }
  return resultados;
}

function fechaReg(f) {
  return f ? { dia: f.dia || 1, mes: f.mes, anio: f.anio, fechaOrigen: f.origen } : { dia: 1, mes: null, anio: null };
}

function fmt(n) { return new Intl.NumberFormat('es-CL').format(Math.round(n)); }
