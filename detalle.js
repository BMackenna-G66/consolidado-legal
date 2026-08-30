// Extracción de DETALLE línea por línea.
// Los montos viven a nivel de documento (boleta/factura), pero varios documentos
// traen adentro el desglose real del trabajo: fecha, solicitante, profesional,
// descripción y horas. Esto lo saca sin agrupar nada.

const pdfjsLib = () => window.pdfjsLib;
const XLSX = () => window.XLSX;

// Reconstruye las líneas visuales de un PDF agrupando los fragmentos por su
// coordenada vertical (pdf.js entrega trozos sueltos, no líneas).
export async function lineasPdf(arrayBuffer) {
  const pdf = await pdfjsLib().getDocument({ data: arrayBuffer }).promise;
  const out = [];
  for (let p = 1; p <= pdf.numPages; p++) {
    const page = await pdf.getPage(p);
    const tc = await page.getTextContent();
    const porY = new Map();
    for (const it of tc.items) {
      if (!it.str || !it.str.trim()) continue;
      const y = Math.round(it.transform[5] / 2.5) * 2.5;
      if (!porY.has(y)) porY.set(y, []);
      porY.get(y).push({ x: it.transform[4], s: it.str });
    }
    for (const y of [...porY.keys()].sort((a, b) => b - a)) {
      const texto = porY.get(y).sort((a, b) => a.x - b.x).map(o => o.s).join(' ').replace(/\s+/g, ' ').trim();
      if (texto) out.push({ pagina: p, y, texto });
    }
  }
  return out;
}

const RE_HORAS = /(\d{1,4}):([0-5]\d)$/;

// Concepto del trabajo según la descripción de la línea. Los anexos de Moraga
// agrupan todo bajo "Asuntos generales", así que el tipo real de trabajo
// (laboral, societario, juicios…) hay que deducirlo del texto de cada línea.
// El orden importa: gana el primer patrón que calza.
export const CONCEPTOS_TRABAJO = [
  ['Asesoría migratoria', /\bvisas?\b|migraci|migrator|nacionalizaci/i],
  ['Marcas y dominios', /\bmarcas?\b|dominios?\b|inapi|propiedad intelectual/i],
  ['Asesoría tributaria', /\bdj\s?\d{3,4}\b|declaraci[oó]n jurada|tributari|impuesto|\btgr\b|declaraci[oó]n de renta/i],
  ['Trámites notariales y publicaciones', /notar[ií]a|notarial|protocolizaci|legalizaci|apostilla|diario oficial|zofri/i],
  ['Societario y directorios', /directorio|\bsod\b|\bsed\b|\bjoa\b|\bjea\b|junta|accionist|accionari|\bacciones\b|\bactas?\b|memoria anual|gerente general|societari|corporativ|cap table|escritura p[uú]blica|estatuto|\bpoder\b|sociedad/i],
  ['Regulatorio (CMF / UAF)', /\bcmf\b|\buaf\b|sernac|regulatori|circular|hecho esencial|oficio reservado|gambling|industrias? (de riesgo|grises)|fintech|derecho (de )?petici/i],
  ['Juicios y litigios', /juicios?\b|\bjpl\b|tribunal|audiencia|expediente|alegato|demanda|querell|sentencia|fraude|\bcausas?\b|monitorio|avenimiento|\breceptor\b|fiscal[ií]a|litigio|absoluci[oó]n|procesal|transacci|judicial/i],
  ['Asesoría laboral', /despidos?\b|desvinculaci|finiquito|carta oferta|ley karin|trabajador|postulante|reclutamiento|previsional|cotizaci|laboral|\bdicom\b/i],
  ['Proyectos y contratos', /m&a|everest|contratos?\b|contrapropuesta|opini[oó]n legal|due diligence|cambio de control|\bnda\b|side letter|itau|nevasa|estructura b2b|grow or go|recaudaci[oó]n/i],
];

export function conceptoDeTrabajo(texto) {
  const t = String(texto || '');
  for (const [nombre, re] of CONCEPTOS_TRABAJO) if (re.test(t)) return nombre;
  return 'Asesoría general';
}

// Las duraciones de Excel llegan como Date (0.0104 de día = 00:15); se formatean
function limpiar(t) {
  return String(t ?? '').replace(/_x000d_/gi, ' ').replace(/[\r\n]+/g, ' ').replace(/\s+/g, ' ').trim();
}

function comoHoras(v) {
  if (v == null || v === '') return '';
  if (v instanceof Date && !isNaN(v)) {
    return `${String(v.getHours()).padStart(2, '0')}:${String(v.getMinutes()).padStart(2, '0')}`;
  }
  if (typeof v === 'number' && v > 0 && v < 1) {
    const min = Math.round(v * 24 * 60);
    return `${String(Math.floor(min / 60)).padStart(2, '0')}:${String(min % 60).padStart(2, '0')}`;
  }
  return String(v);
}
const RE_FECHA = /^(\d{1,2})\/(\d{1,2})\/(\d{4})\s+(.*)$/;

// Anexos de Moraga (HH…, NCCL … detalle cobro). El bloque "DETALLE TRABAJOS"
// tiene la jerarquía categoría → "Solicitado por X" → profesional → líneas con
// fecha. Las horas aparecen en el formato 2026 y no en el 2025; ambos se leen.
export async function detalleMoraga(arrayBuffer, meta) {
  const lineas = await lineasPdf(arrayBuffer);
  const inicio = lineas.findIndex(l => /DETALLE TRABAJOS/i.test(l.texto));
  if (inicio < 0) return [];               // solo trae el resumen, no hay detalle
  const filas = [];
  let categoria = '', solicitante = '', profesional = '', vistoSolicitante = false, pendiente = null;
  const cerrar = () => {
    if (pendiente && pendiente.descripcion) {
      pendiente.concepto = conceptoDeTrabajo(pendiente.descripcion);
      filas.push(pendiente);
    }
    pendiente = null;
  };

  for (const { texto } of lineas.slice(inicio + 1)) {
    if (/^Total general/i.test(texto) || /^P[áa]gina/i.test(texto) || /moragaycia\.cl/i.test(texto)
        || /^Grupo Cliente/i.test(texto) || /Suma de Horas/i.test(texto)) { cerrar(); continue; }
    if (/^Etiquetas de fila/i.test(texto)) { cerrar(); categoria = ''; solicitante = ''; profesional = ''; vistoSolicitante = false; continue; }

    const mHoras = texto.match(RE_HORAS);
    if (/^\d{1,4}:[0-5]\d$/.test(texto)) { if (pendiente && !pendiente.horas) pendiente.horas = texto; continue; }
    const sinHoras = mHoras ? texto.slice(0, texto.length - mHoras[0].length).trim() : texto;
    if (!sinHoras) continue;

    if (/^Solicitad[oa]\s+por\s+/i.test(sinHoras) || /^Sin\s+Solicitante$/i.test(sinHoras)) {
      cerrar();
      // "AMF - Álvaro Moraga Fritz" → "Álvaro Moraga Fritz" (iniciales de cortesía)
      solicitante = /^Sin/i.test(sinHoras) ? ''
        : sinHoras.replace(/^Solicitad[oa]\s+por\s+/i, '').replace(/^[A-ZÁÉÍÓÚÑ]{2,4}\s*-\s*/, '').trim();
      profesional = ''; vistoSolicitante = true;
      continue;
    }

    const mFecha = sinHoras.match(RE_FECHA);
    if (mFecha) {
      cerrar();
      pendiente = {
        ...meta, categoria, solicitante, profesional,
        fecha: `${mFecha[1].padStart(2, '0')}/${mFecha[2].padStart(2, '0')}/${mFecha[3]}`,
        dia: +mFecha[1], mes: +mFecha[2], anio: +mFecha[3],
        descripcion: mFecha[4].trim(), horas: mHoras ? mHoras[0] : '', unidad: 'horas trabajadas',
      };
      continue;
    }

    // Encabezado de agrupación: antes del primer "Solicitado por" es la categoría;
    // después, es el profesional que ejecutó las líneas siguientes.
    const palabras = sinHoras.split(/\s+/);
    if (/^[A-ZÁÉÍÓÚÑ]/.test(sinHoras) && palabras.length <= 6 && !/[.;:]$/.test(sinHoras)) {
      cerrar();
      if (vistoSolicitante) profesional = sinHoras; else categoria = sinHoras;
      continue;
    }

    if (pendiente) pendiente.descripcion = (pendiente.descripcion + ' ' + sinHoras).replace(/\s+/g, ' ').trim();
  }
  cerrar();
  return filas;
}

// Minutas de Andes Latam: hojas de timesheet (horas en USD) y de gastos (soles).
export function detalleAndes(arrayBuffer, meta) {
  const wb = XLSX().read(arrayBuffer, { type: 'array', cellDates: true });
  const filas = [];
  const conOk = wb.SheetNames.filter(n => /\bok\b/i.test(n));
  const hojas = conOk.length ? conOk : wb.SheetNames;
  for (const hoja of hojas) {
    const rows = XLSX().utils.sheet_to_json(wb.Sheets[hoja], { header: 1, raw: true, defval: null });
    const encabezados = rows.map((r, i) => (r && r.some(c => /descripci/i.test(String(c ?? '')))) ? i : -1).filter(i => i >= 0);
    for (let bloque = 0; bloque < encabezados.length; bloque++) {
    const hIdx = encabezados[bloque];
    const fin = bloque + 1 < encabezados.length ? encabezados[bloque + 1] : rows.length;
    const H = rows[hIdx].map(h => String(h ?? '').trim());
    const col = (...pats) => { for (const p of pats) { const i = H.findIndex(h => p.test(h)); if (i >= 0) return i; } return -1; };
    const c = {
      n: col(/^N°$/i), dia: col(/^d[ií]a$/i), mes: col(/^mes$/i), anio: col(/^a[ñn]o$/i),
      abogado: col(/^abogado$/i, /^nombre$/i), solicitante: col(/^solicitante$/i),
      desc: col(/descripci/i), dur: col(/^duraci[oó]n$/i), tarifa: col(/tarifa/i),
      valor: col(/valor\s*\(usd\)/i), total: col(/^total$/i), fecha: col(/^fecha$/i),
      cat: col(/categor/i), tipoDoc: col(/tipo\s*documento/i),
    };
    for (const r of rows.slice(hIdx + 1, fin)) {
      const desc = limpiar(r[c.desc]);
      if (!desc || /^total/i.test(desc)) continue;
      const valor = c.valor >= 0 ? Number(r[c.valor]) : (c.total >= 0 ? Number(r[c.total]) : NaN);
      if (!isFinite(valor) || valor <= 0) continue;
      let dia = c.dia >= 0 ? Number(r[c.dia]) : null;
      let mes = c.mes >= 0 ? Number(r[c.mes]) : null;
      let anio = c.anio >= 0 ? Number(r[c.anio]) : null;
      if (!mes && c.fecha >= 0) {
        const cel = r[c.fecha];
        if (cel instanceof Date && !isNaN(cel)) { dia = cel.getDate(); mes = cel.getMonth() + 1; anio = cel.getFullYear(); }
        else {
          const m = String(cel ?? '').match(/(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
          if (m) { dia = +m[1]; mes = +m[2]; anio = +m[3] < 100 ? 2000 + +m[3] : +m[3]; }
        }
      }
      filas.push({
        ...meta,
        documento: c.n >= 0 && r[c.n] != null ? String(r[c.n]) : '',
        hoja, dia: dia || null, mes: mes || null, anio: anio || null,
        fecha: (dia && mes && anio) ? `${String(dia).padStart(2, '0')}/${String(mes).padStart(2, '0')}/${anio}` : '',
        categoria: c.cat >= 0 ? limpiar(r[c.cat]) : '',
        solicitante: c.solicitante >= 0 ? limpiar(r[c.solicitante]) : '',
        profesional: c.abogado >= 0 ? limpiar(r[c.abogado]) : '',
        descripcion: desc,
        horas: c.dur >= 0 ? comoHoras(r[c.dur]) : '',
        tarifa: c.tarifa >= 0 ? Number(r[c.tarifa]) || '' : '',
        valor, monedaValor: c.valor >= 0 ? 'USD' : 'PEN',
        unidad: c.valor >= 0 ? 'honorario por hora' : 'gasto reembolsable',
        tipoDocumento: c.tipoDoc >= 0 ? limpiar(r[c.tipoDoc]) : '',
      });
    }
    }
  }
  return filas;
}

// Nota de cobro Sensus: hitos de la Ley 20.009 con cantidad de unidades.
export async function detalleSensus(arrayBuffer, meta) {
  const lineas = await lineasPdf(arrayBuffer);
  const filas = [];
  for (const { texto } of lineas) {
    const m = texto.match(/^([a-e])\)\s*(.+?)\s+(\d{1,3})$/);
    if (!m) continue;
    const descripcion = `${m[1]}) ${m[2].replace(/\s+/g, ' ').trim()}`;
    if (filas.some(f => f.descripcion === descripcion && f.valor === +m[3])) continue;
    filas.push({ ...meta, categoria: 'Hito Ley 20.009', descripcion, valor: +m[3], unidad: 'unidades', monedaValor: 'UF' });
  }
  return filas;
}

// Despachador: devuelve las líneas de detalle de un archivo, o [] si no tiene.
export async function extraerDetalle({ nombre, ruta, arrayBuffer }, meta) {
  const ext = nombre.split('.').pop().toLowerCase();
  const base = { ...meta, archivo: nombre, ruta };
  try {
    if (/andes latam/i.test(ruta) && ext === 'xlsx') return detalleAndes(arrayBuffer, base);
    if (ext === 'pdf' && (/^hh\b/i.test(nombre) || /detalle/i.test(nombre) || /^ncc/i.test(nombre))) {
      return await detalleMoraga(arrayBuffer, base);
    }
    if (/sensus/i.test(ruta) && ext === 'pdf' && /^nc/i.test(nombre)) return await detalleSensus(arrayBuffer, base);
  } catch { /* documento sin detalle legible */ }
  return [];
}
