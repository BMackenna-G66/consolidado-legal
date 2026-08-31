// Lectores por proveedor. Cada parser recibe el contenido del archivo y devuelve
// { registros: [...], estado: 'ok'|'supuesto'|'error', nota }
// Registro: { dia, mes, anio, pais, proveedor, categoria, concepto, moneda, montoOrigen, clp, archivo, detalle }

import { fechaDeRuta, parseFechaTexto, numeroCL, numeroUS, proveedorDeRuta, montosDelTexto, monedaDelTexto } from './config.js';
import { detalleMoraga } from './detalle.js';

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

// Grupos (concepto del trabajo, solicitante) del DETALLE TRABAJOS de un anexo
// de Moraga, con las horas de cada grupo: es la pauta para repartir la boleta
// del mes, que trae un solo monto sin desglose.
async function gruposDelAnexo(arrayBuffer) {
  const grupos = new Map();
  for (const l of await detalleMoraga(arrayBuffer, {})) {
    const k = `${l.concepto}|${l.solicitante || ''}`;
    const g = grupos.get(k) || { concepto: l.concepto, solicitante: l.solicitante || '', minutos: 0, lineas: 0 };
    const m = String(l.horas || '').match(/^(\d{1,4}):([0-5]\d)$/);
    if (m) g.minutos += +m[1] * 60 + +m[2];
    g.lineas++;
    grupos.set(k, g);
  }
  return [...grupos.values()];
}

// Reparte un total en partes proporcionales a los pesos cuidando que la suma
// dé exactamente el total: el residuo del redondeo se carga a la parte mayor.
function repartir(total, pesos, dec) {
  const W = pesos.reduce((s, w) => s + w, 0) || 1;
  const f = 10 ** dec;
  const partes = pesos.map(w => Math.round(total * w / W * f) / f);
  const dif = Math.round((total - partes.reduce((s, p) => s + p, 0)) * f) / f;
  const iMax = partes.indexOf(Math.max(...partes));
  partes[iMax] = Math.round((partes[iMax] + dif) * f) / f;
  return partes;
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
      categoria: 'Asesoría Profesional', concepto: 'Honorarios', moneda: 'CLP',
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
  let gruposAnexo = [];
  try { gruposAnexo = await gruposDelAnexo(buf); } catch { /* anexo sin detalle legible */ }
  if (!uf) return { registros: [], estado: 'omitido', solicitantesAnexo, gruposAnexo,
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
      categoria: 'Asesoría Profesional', concepto: 'Honorarios', moneda: 'UF',
      solicitante: solicitanteResumen(solicitantesDelTexto(texto)),
      montoOrigen: uf, clp, archivo, detalle: `Nota de cobro ${uf} UF`,
    }],
    solicitantesAnexo, gruposAnexo,
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
      const clpTotal = Math.round(total * params.USD_CLP);
      const lineas = parseAndesPorLineas(buf, ruta, params, archivo);
      if (lineas.registros.length) {
        // Detalle real de la minuta, más un ajuste para cuadrar con el total facturado
        for (const r of lineas.registros) {
          if (!r.mes) r.mes = mes;
          if (!r.anio) r.anio = anio || rf.anio || null;
        }
        const sumaLineas = lineas.registros.reduce((sm, r) => sm + r.clp, 0);
        const dif = clpTotal - sumaLineas;
        if (Math.abs(dif) > 1000) {
          lineas.registros.push({
            dia: 1, mes, anio, pais: 'Perú', proveedor: 'Andes Latam',
            categoria: 'Fee mensual / ajuste de minuta', concepto: 'Honorarios', moneda: 'USD',
            solicitante: '', montoOrigen: +(dif / params.USD_CLP).toFixed(2), clp: dif, archivo,
            detalle: `Cuadra el detalle con el total facturado${factura ? ' ' + factura : ''} (retainer / IGV)`,
          });
        }
        return {
          registros: lineas.registros, estado: 'supuesto',
          nota: `${lineas.registros.length} líneas de la minuta; total facturado USD ${total.toFixed(2)} = ${fmt(clpTotal)} (incluye IGV)`,
        };
      }
      return {
        registros: [{
          dia: (f && f.dia) || 1, mes, anio, pais: 'Perú', proveedor: 'Andes Latam',
          categoria: 'Asesoría legal Andes Latam', concepto: 'Honorarios', moneda: 'USD',
          solicitante: solicitanteResumen(solicitantes),
          montoOrigen: total, clp: clpTotal, archivo,
          detalle: `Minuta de liquidación${factura ? ' ' + factura : ''} (honorarios + gastos + IGV)`,
        }],
        estado: 'supuesto',
        nota: `Total cobro USD ${total.toFixed(2)} × ${params.USD_CLP} = ${fmt(clpTotal)} (incluye IGV)`,
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
  const conOk = wb.SheetNames.filter(n => /\bok\b/i.test(n));
  const hojasALeer = conOk.length ? conOk : wb.SheetNames;
  for (const nombre of hojasALeer) {
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
          concepto: 'Honorarios', moneda: 'USD', montoOrigen: usd, clp, archivo,
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
  for (const reg of registros) {
    if (!reg.mes && rf.mes) reg.mes = rf.mes;
    if (!reg.anio) reg.anio = rf.anio || null;
  }
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


// El proveedor de cada fila histórica se recupera cruzando su monto contra las
// hojas por proveedor del MISMO libro (de donde nació la Base resumen): si un
// monto aparece en la hoja de Moraga, esa fila es de Moraga. Para Argentina y
// Colombia el proveedor viene por fila en su propia hoja.
function atribuidorDeProveedores(wb) {
  const HOJAS = [
    { re: /moraga/i,            pais: 'Chile',    prov: 'Alvaro Moraga' },
    { re: /sensus/i,            pais: 'Chile',    prov: 'Sensus Legis' },
    { re: /dento/i,             pais: 'Chile',    prov: 'Dentons' },
    { re: /aninat/i,            pais: 'Chile',    prov: 'Aninat' },
    { re: /chile\s*-\s*otros/i, pais: 'Chile',    prov: 'Juicios y contingencias' },
    { re: /andes/i,             pais: 'Perú',     prov: 'Andes Latam' },
    { re: /niubox/i,            pais: 'Perú',     prov: 'NiuBox' },
    { re: /per[úu]\s*-\s*otros/i, pais: 'Perú',  prov: 'Investigaciones y juicios' },
    { re: /colombia\s*-\s*externo/i, pais: 'Colombia', prov: 'Requerimientos externos' },
  ];
  const porMonto = new Map();
  const add = (pais, v, prov) => {
    if (!isFinite(v) || v < 900 || v > 5e8) return;
    const k = `${pais}|${Math.round(v)}`;
    if (!porMonto.has(k)) porMonto.set(k, new Set());
    porMonto.get(k).add(prov);
  };
  const arregla = (t) => {
    t = String(t).trim().split(/\s+-\s+/)[0].trim();
    if (!t) return '';
    return t === t.toUpperCase()
      ? t.toLowerCase().replace(/(^|\s)\S/g, c => c.toUpperCase())
      : t;
  };
  for (const nombre of wb.SheetNames) {
    const filas = filasDeHoja(wb.Sheets[nombre]);
    const cfg = HOJAS.find(h => h.re.test(nombre));
    if (cfg) {
      for (const fila of filas) for (const c of (fila || [])) {
        if (typeof c === 'number') add(cfg.pais, c, cfg.prov);
      }
      continue;
    }
    if (/argentina/i.test(nombre) || /colombia\s*-\s*interno/i.test(nombre)) {
      const pais = /argentina/i.test(nombre) ? 'Argentina' : 'Colombia';
      const hIdx = filas.findIndex(r => r && r.some(c => /proveedor|carlos gomez/i.test(String(c ?? ''))));
      if (hIdx < 0) continue;
      const provCol = filas[hIdx].findIndex(c => /proveedor|carlos gomez/i.test(String(c ?? '')));
      for (const fila of filas.slice(hIdx + 1)) {
        const prov = arregla(fila?.[provCol] ?? '');
        if (!prov) continue;
        for (const c of fila) if (typeof c === 'number') add(pais, c, prov);
      }
    }
  }
  return (pais, clp, categoria) => {
    const set = porMonto.get(`${pais}|${Math.round(clp)}`);
    if (set && set.size === 1) return [...set][0];
    const cat = String(categoria || '');
    if (pais === 'Chile') {
      if (/20\.?009/i.test(cat)) return 'Sensus Legis';
      if (/stablecoin|cuenta remunerada|licencia fintech/i.test(cat)) return 'Dentons';
      if (/asesor[ií]a (profesional|legal)|galleguillos|moraga/i.test(cat)) return 'Alvaro Moraga';
    }
    if (set && set.size > 1) return null;
    if (pais === 'Perú' && !/juicio|investigaci/i.test(cat)) return 'Andes Latam';
    return null;
  };
}

// Colombia se lleva en sus propias hojas ("Colombia - Interno" = Garrigues,
// "Colombia - Externo" = requerimientos), y NO siempre se traspasa a la hoja
// "Base resumen": ahí Colombia se corta en octubre 2025 aunque las hojas de
// origen llegan hasta junio 2026. Se leen directo, solo en los periodos que la
// Base resumen no cubre, para no duplicar lo ya reportado.
function parseHojasColombia(wb, cubiertos, archivo) {
  const registros = [];
  const per = d => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
  const nombreHoja = re => wb.SheetNames.find(n => re.test(n));

  const interno = nombreHoja(/colombia\s*-\s*interno/i);
  if (interno) {
    const filas = filasDeHoja(wb.Sheets[interno]);
    const H = (filas[0] || []).map(h => String(h ?? ''));
    const c = {
      fecha: buscaColumna(H, /fecha/i), prov: 1,
      desc: buscaColumna(H, /descripci/i), cat: buscaColumna(H, /categor/i),
      valor: buscaColumna(H, /valor/i), tipo: buscaColumna(H, /^tipo$/i),
    };
    for (const f of filas.slice(1)) {
      const d = f?.[c.fecha];
      if (!(d instanceof Date) || isNaN(d)) continue;
      const v = numeroUS(f[c.valor]);
      if (!isFinite(v) || v <= 0) continue;
      if (cubiertos.has(per(d))) continue;
      const prov = String(f[c.prov] ?? '').trim();
      registros.push({
        dia: d.getDate(), mes: d.getMonth() + 1, anio: d.getFullYear(), pais: 'Colombia',
        proveedor: prov ? prov.charAt(0).toUpperCase() + prov.slice(1).toLowerCase() : 'Garrigues',
        categoria: String(f[c.cat] ?? 'Asesoría legal Colombia').trim(),
        concepto: /honorario|asesor/i.test(String(f[c.tipo] ?? '')) ? 'Honorarios' : 'Gastos',
        moneda: 'CLP', montoOrigen: v, clp: Math.round(v), archivo,
        detalle: String(f[c.desc] ?? '').replace(/\s+/g, ' ').trim().slice(0, 160),
        solicitante: '', fuente: 'base',
      });
    }
  }

  const externo = nombreHoja(/colombia\s*-\s*externo/i);
  if (externo) {
    const filas = filasDeHoja(wb.Sheets[externo]);
    const hIdx = filas.findIndex(f => f && f.some(x => /cop\s*a\s*clp/i.test(String(x ?? ''))));
    if (hIdx >= 0) {
      const H = filas[hIdx].map(h => String(h ?? ''));
      const c = {
        fecha: buscaColumna(H, /fecha de requerimiento/i, /^fecha/i),
        cat: buscaColumna(H, /categor/i), desc: buscaColumna(H, /descripci/i),
        clp: buscaColumna(H, /cop\s*a\s*clp/i),
      };
      for (const f of filas.slice(hIdx + 1)) {
        const d = f?.[c.fecha];
        if (!(d instanceof Date) || isNaN(d)) continue;
        const v = numeroUS(f[c.clp]);
        if (!isFinite(v) || v <= 0) continue;
        if (cubiertos.has(per(d))) continue;
        registros.push({
          dia: d.getDate(), mes: d.getMonth() + 1, anio: d.getFullYear(), pais: 'Colombia',
          proveedor: 'Requerimientos externos',
          categoria: String(f[c.cat] ?? 'Requerimientos').trim(),
          concepto: 'Gastos', moneda: 'CLP', montoOrigen: v, clp: Math.round(v), archivo,
          detalle: String(f[c.desc] ?? '').replace(/\s+/g, ' ').trim().slice(0, 160),
          solicitante: '', fuente: 'base',
        });
      }
    }
  }
  return registros;
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
  const atribuir = atribuidorDeProveedores(wb);
  let atribuidos = 0;
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
      pais: String(pais).trim(),
      proveedor: (() => {
        const p = atribuir(String(pais).trim(), clp, r[col.categoria]);
        if (p) atribuidos++;
        return p || 'Base histórica (Excel)';
      })(),
      categoria: String(r[col.categoria] ?? 'Sin categoría').trim(),
      concepto, moneda: 'CLP', montoOrigen: clp, clp: Math.round(clp),
      archivo, detalle: '', fuente: 'base',
    });
  }
  // Colombia: completar los periodos que la Base resumen no alcanzó a recoger
  const cubiertosCol = new Set(registros.filter(r => r.pais === 'Colombia')
    .map(r => `${r.anio}-${String(r.mes).padStart(2, '0')}`));
  const colombia = parseHojasColombia(wb, cubiertosCol, archivo);
  registros.push(...colombia);
  const perCol = [...new Set(colombia.map(r => `${r.anio}-${String(r.mes).padStart(2, '0')}`))].sort();

  return { registros, estado: 'ok', fuente: 'base',
    nota: `${registros.length} movimientos históricos; proveedor identificado en ${atribuidos} cruzando las hojas del propio Excel`
      + (colombia.length ? ` · ${colombia.length} movimientos de Colombia recuperados de sus hojas de origen (${perCol.join(', ')}), que no estaban en la hoja "Base resumen"` : '') };
}

// Los documentos de las carpetas son la fuente principal: traen proveedor,
// solicitante y detalle. La base histórica del Excel solo rellena los países+meses
// donde NO hay documentos (evita doble conteo sin perder detalle).
export function aplicarBaseHistorica(resultados) {
  const cubiertos = new Set();
  for (const res of resultados) {
    if (res.fuente === 'base' || res.fuente === 'manual') continue;
    for (const r of (res.registros || [])) {
      if (r.anio && r.mes && r.clp > 0) cubiertos.add(`${r.pais}|${r.anio}-${r.mes}`);
    }
  }
  for (const res of resultados) {
    if (res.fuente !== 'base') continue;
    const antes = res.registros.length;
    res.registros = res.registros.filter(r => !(r.anio && r.mes) || !cubiertos.has(`${r.pais}|${r.anio}-${r.mes}`));
    const quitados = antes - res.registros.length;
    res.nota = `${antes} movimientos históricos; se usan ${res.registros.length} (en ${quitados} los documentos de las carpetas mandan y aportan el detalle)`;
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
  if (/\.xlsx?$/i.test(nombre) && /andes\s*latam/i.test(nombre)) {
    const r = parseAndesLatam(arrayBuffer, ruta, params, nombre);
    if ((!prov || prov.id !== 'andes') && r.registros.length) {
      r.nota = (r.nota || '') + ' · ⚠ Minuta de Andes Latam archivada fuera de su carpeta';
      if (r.estado === 'ok') r.estado = 'supuesto';
    }
    return { ...base, ...r, proveedor: 'Andes Latam', pais: 'Perú' };
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
      // No trae montos, pero sí el detalle del mes (conceptos, solicitantes y
      // horas): se guarda para repartir los cobros de la misma carpeta.
      let solicitantesAnexo = [], gruposAnexo = [];
      try { solicitantesAnexo = solicitantesDelTexto(await textoPdf(arrayBuffer)); } catch { /* anexo ilegible */ }
      try { gruposAnexo = await gruposDelAnexo(arrayBuffer); } catch { /* anexo sin detalle legible */ }
      return {
        ...base, estado: 'omitido', registros: [], solicitantesAnexo, gruposAnexo,
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

// Los anexos de detalle traen el desglose del mes (concepto del trabajo,
// solicitante y horas) pero no los montos; las boletas traen el monto pero sin
// desglose. Se cruzan por carpeta (= mes): el monto de la boleta se reparte
// entre los grupos del anexo en proporción a las horas trabajadas.
export function enriquecerSolicitantes(resultados) {
  const carpetaDe = res => String(res.ruta || '').split('/').slice(0, -1).join('/');
  const porCarpeta = new Map();
  const gruposPorCarpeta = new Map();
  for (const res of resultados) {
    const carpeta = carpetaDe(res);
    if (res.solicitantesAnexo && res.solicitantesAnexo.length) {
      const set = porCarpeta.get(carpeta) || new Set();
      res.solicitantesAnexo.forEach(s => set.add(s));
      porCarpeta.set(carpeta, set);
    }
    if (res.gruposAnexo && res.gruposAnexo.length) {
      const m = gruposPorCarpeta.get(carpeta) || new Map();
      for (const g of res.gruposAnexo) {
        const k = `${g.concepto}|${g.solicitante}`;
        const acc = m.get(k) || { concepto: g.concepto, solicitante: g.solicitante, minutos: 0, lineas: 0 };
        acc.minutos += g.minutos;
        acc.lineas += g.lineas;
        m.set(k, acc);
      }
      gruposPorCarpeta.set(carpeta, m);
    }
  }

  // 1) Prorrateo de las boletas/notas de cobro de Moraga por las horas del anexo:
  //    cada grupo (concepto, solicitante) recibe su parte proporcional del monto.
  const hh = min => `${Math.floor(min / 60)}:${String(min % 60).padStart(2, '0')}`;
  for (const res of resultados) {
    if (!res.registros || !res.registros.length) continue;
    if (!/^bh|ncc/i.test(res.archivo || '')) continue; // solo la boleta o nota de cobro del mes
    const grupos = [...(gruposPorCarpeta.get(carpetaDe(res)) || new Map()).values()];
    if (!grupos.length) continue;
    const pesos = grupos.map(g => g.minutos || 1); // una línea sin horas pesa 1 minuto
    const totalMin = pesos.reduce((s, w) => s + w, 0);
    let repartidos = 0;
    res.registros = res.registros.flatMap(r => {
      if (r.proveedor !== 'Alvaro Moraga' || r.concepto !== 'Honorarios') return [r];
      repartidos++;
      const clps = repartir(r.clp, pesos, 0);
      const origenes = r.moneda === 'CLP' ? clps : repartir(r.montoOrigen, pesos, 2);
      return grupos.map((g, i) => ({
        ...r,
        categoria: g.concepto,
        solicitante: g.solicitante,
        montoOrigen: origenes[i],
        clp: clps[i],
        detalle: `${r.detalle} — prorrateo por horas del anexo (${hh(pesos[i])} de ${hh(totalMin)} h)`,
        prorrateo: true,
      }));
    });
    if (repartidos) {
      res.nota = (res.nota || '')
        + ` · monto repartido por horas del anexo en ${grupos.length} grupo(s) de concepto y solicitante`;
    }
  }

  // 2) Respaldo: registros sin desglose heredan los nombres sueltos del anexo.
  for (const res of resultados) {
    if (!res.registros || !res.registros.length) continue;
    const lista = [...(porCarpeta.get(carpetaDe(res)) || [])];
    if (!lista.length) continue;
    let n = 0;
    for (const r of res.registros) {
      if (r.prorrateo || r.solicitante) continue;
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
      if (previo && (previo === res.archivo || /\(\d\)|copia/i.test(res.archivo))) {
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
