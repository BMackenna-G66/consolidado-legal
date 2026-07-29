// Orquestador: carga archivos, los procesa con los parsers y pinta el reporte.

import { DEFAULT_PARAMS, MESES } from './config.js';
import { procesarArchivo, deduplicar, aplicarBaseHistorica } from './parsers.js';
import { leerCarpetaLocal, leerSharePoint, graphDisponible, soportaFSA, elegirCarpetaFSA, leerCarpetaRecordada, carpetaGuardada } from './fuentes.js';
import { construirPivot, renderPivot, renderResumenTarjetas, exportarCSV } from './reporte.js';
import { clasificar, overrides, setOverride, CONCEPTOS, clasificarConIA, apiKeyGemini, setApiKeyGemini, normalizaCat } from './clasificador.js';
import { resultadosDemo } from './demo.js';

window.pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.worker.min.js';

const $id = id => document.getElementById(id);
const PARAMS_KEY = 'consolidado-legal-params';

let resultados = [];   // salida por archivo (estado, nota, registros)
let archivosCrudos = null; // para reprocesar al cambiar parámetros

function paramsActuales() {
  try { return { ...DEFAULT_PARAMS, ...JSON.parse(localStorage.getItem(PARAMS_KEY) || '{}') }; }
  catch { return { ...DEFAULT_PARAMS }; }
}

// ---------- carga ----------

$id('btn-local').addEventListener('click', async () => {
  if (!soportaFSA()) { $id('input-carpeta').click(); return; }
  try {
    progreso('Leyendo carpeta…');
    const { archivos } = await elegirCarpetaFSA(progreso);
    archivosCrudos = archivos;
    await procesarYRender();
  } catch (e) {
    progreso(e.name === 'AbortError' ? '' : '⚠ ' + (e.message || e));
  }
});
$id('input-carpeta').addEventListener('change', async e => {
  if (!e.target.files.length) return;
  progreso('Leyendo carpeta…');
  archivosCrudos = await leerCarpetaLocal(e.target.files, progreso);
  await procesarYRender();
});

// Si hay una carpeta recordada de una visita anterior, ofrecer actualización en un clic
(async () => {
  const dir = await carpetaGuardada();
  if (!dir) return;
  const btn = $id('btn-actualizar');
  btn.hidden = false;
  btn.textContent = `↻ Actualizar desde "${dir.name}"`;
  btn.addEventListener('click', async () => {
    try {
      progreso('Releyendo carpeta…');
      const { archivos } = await leerCarpetaRecordada(progreso);
      archivosCrudos = archivos;
      await procesarYRender();
    } catch (e) { progreso('⚠ ' + (e.message || e)); }
  });
})();

$id('btn-sharepoint').addEventListener('click', async () => {
  if (!graphDisponible()) {
    progreso('⚠ Falta configurar clientId de Azure AD en config.js (ver README). Mientras tanto usa la opción de carpeta local.');
    return;
  }
  try {
    progreso('Conectando con Microsoft…');
    archivosCrudos = await leerSharePoint(progreso);
    await procesarYRender();
  } catch (e) {
    progreso('⚠ ' + (e.message || e));
  }
});

$id('btn-recargar').addEventListener('click', () => location.reload());
$id('btn-csv').addEventListener('click', () => exportarCSV(registrosFiltrados()));
$id('btn-params').addEventListener('click', () => $id('panel-params').classList.toggle('visible'));
$id('f-gastos').addEventListener('change', renderTodo);
$id('f-honorarios').addEventListener('change', renderTodo);
$id('f-juicios').addEventListener('change', renderTodo);
$id('f-anio').addEventListener('change', renderTodo);
$id('f-mes').addEventListener('change', renderTodo);
$id('btn-mantenedor').addEventListener('click', () => $id('mantenedor').classList.toggle('seccion-oculta'));

function progreso(msg) { $id('progreso').textContent = msg; }

// ---------- pipeline ----------

async function procesarYRender() {
  const params = paramsActuales();
  progreso(`Interpretando ${archivosCrudos.length} archivos…`);
  resultados = [];
  for (const a of archivosCrudos) resultados.push(await procesarArchivo(a, params));
  deduplicar(resultados);
  aplicarBaseHistorica(resultados);
  finalizarRender(params);
}

function finalizarRender(params) {
  progreso('');
  $id('portada').style.display = 'none';
  $id('vista-reporte').style.display = 'block';
  for (const b of ['btn-params', 'btn-mantenedor', 'btn-csv', 'btn-recargar']) $id(b).hidden = false;
  montarPanelParams(params);
  montarFiltrosFecha();
  renderTodo();
}

// Aplica la clasificación (manual > IA > regla > parser) a cada registro
function registrosClasificados() {
  return resultados.flatMap(r => r.registros).map(r => {
    const { concepto, origen } = clasificar(r);
    return { ...r, concepto, clasifOrigen: origen };
  });
}

function registrosFiltrados() {
  const ver = {
    'Gastos': $id('f-gastos').checked,
    'Honorarios': $id('f-honorarios').checked,
    'Juicios y otros': $id('f-juicios').checked,
  };
  const fAnio = $id('f-anio').value ? +$id('f-anio').value : null;
  const fMes = $id('f-mes').value ? +$id('f-mes').value : null;
  return registrosClasificados().filter(r => ver[r.concepto]
    && (!fAnio || r.anio === fAnio)
    && (!fMes || r.mes === fMes));
}

// Puebla los combos de año y mes con lo que exista en los datos, conservando la selección
function montarFiltrosFecha() {
  const regs = resultados.flatMap(r => r.registros);
  const anios = [...new Set(regs.map(r => r.anio).filter(Boolean))].sort();
  const selA = $id('f-anio'), selM = $id('f-mes');
  const prevA = selA.value, prevM = selM.value;
  selA.innerHTML = '<option value="">Todos</option>' + anios.map(a => `<option value="${a}">${a}</option>`).join('');
  selM.innerHTML = '<option value="">Todos</option>' + MESES.map((m, i) => `<option value="${i + 1}">${m}</option>`).join('');
  selA.value = prevA; selM.value = prevM;
}

function renderTodo() {
  const regs = registrosFiltrados();
  renderResumenTarjetas($id('tarjetas'), regs);
  renderPivot($id('pivot-principal'), construirPivot(regs), 'Resumen de gastos por conceptos legales / administrativos (CLP)');
  renderMantenedor();
  renderArchivos();
}

// ---------- mantenedor de clasificación ----------

function renderMantenedor() {
  const cont = $id('tabla-mantenedor');
  if (!cont) return;
  const porCat = new Map();
  for (const r of registrosClasificados()) {
    const k = normalizaCat(r.categoria);
    if (!porCat.has(k)) porCat.set(k, { categoria: r.categoria, total: 0, n: 0, concepto: r.concepto, origen: r.clasifOrigen, provs: new Set() });
    const e = porCat.get(k);
    e.total += r.clp; e.n++; e.provs.add(r.proveedor);
    if (r.clasifOrigen === 'Manual') { e.concepto = r.concepto; e.origen = 'Manual'; }
  }
  const fmt = n => '$' + new Intl.NumberFormat('es-CL').format(Math.round(n));
  const filas = [...porCat.values()].sort((a, b) => b.total - a.total).map(e => {
    const opts = CONCEPTOS.map(c => `<option ${c === e.concepto ? 'selected' : ''}>${c}</option>`).join('');
    return `<tr><td class="etiqueta" title="${[...e.provs].join(', ')}">${e.categoria}</td>
      <td class="num">${e.n}</td><td class="num">${fmt(e.total)}</td>
      <td><select data-cat="${e.categoria.replace(/"/g, '&quot;')}">${opts}</select></td>
      <td><span class="origen origen-${e.origen.toLowerCase()}">${e.origen}</span></td></tr>`;
  }).join('');
  cont.innerHTML = `<table class="pivot"><thead><tr><th class="etiqueta">Partida / categoría</th><th>Movs.</th><th>Total CLP</th><th>Clasificación</th><th>Origen</th></tr></thead><tbody>${filas}</tbody></table>`;
  cont.querySelectorAll('select').forEach(sel => sel.addEventListener('change', () => {
    setOverride(sel.dataset.cat, sel.value);
    renderTodo();
  }));
}

$id('btn-ia').addEventListener('click', async () => {
  if (!apiKeyGemini()) {
    const k = prompt('Pega tu API key de Gemini (queda solo en esta sesión del navegador):');
    if (!k) return;
    setApiKeyGemini(k);
  }
  const btn = $id('btn-ia');
  btn.disabled = true; btn.textContent = 'Clasificando…';
  try {
    const cats = resultados.flatMap(r => r.registros).filter(r => r.fuente !== 'base').map(r => r.categoria);
    const { nuevas } = await clasificarConIA(cats);
    $id('nota-ia').textContent = nuevas ? `IA clasificó ${nuevas} partidas nuevas (revisable abajo).` : 'No había partidas nuevas por clasificar.';
    renderTodo();
  } catch (e) {
    $id('nota-ia').textContent = '⚠ ' + (e.message || e);
  } finally {
    btn.disabled = false; btn.textContent = '🤖 Clasificar partidas con IA';
  }
});

function renderArchivos() {
  const orden = { error: 0, supuesto: 1, omitido: 2, ok: 3 };
  const etiqueta = { ok: '✅ Procesado', supuesto: '⚠️ Con supuestos', error: '❌ No leído', omitido: '⏭ Omitido' };
  const html = [...resultados]
    .sort((a, b) => orden[a.estado] - orden[b.estado])
    .map(r => `<div class="archivo-fila"><span class="estado ${r.estado}">${etiqueta[r.estado]}</span><span class="nombre">${r.ruta || r.archivo}</span><span class="nota">${r.nota || ''}</span></div>`)
    .join('');
  $id('lista-archivos').innerHTML = html || '<div class="archivo-fila">Sin archivos.</div>';
}

// ---------- parámetros ----------

const PARAM_LABELS = {
  USD_CLP: 'USD → CLP', PEN_CLP: 'PEN (sol) → CLP', COP_CLP: 'COP → CLP', ARS_CLP: 'ARS → CLP',
  UF_CLP: 'Valor UF (CLP)', IVA_PE: 'IGV Perú (fracción)', IVA_CL: 'IVA Chile (fracción)',
};

function montarPanelParams(params) {
  const grid = $id('params-grid');
  grid.innerHTML = Object.keys(PARAM_LABELS).map(k =>
    `<label>${PARAM_LABELS[k]}<input type="number" step="any" data-param="${k}" value="${params[k]}"></label>`).join('');
  grid.querySelectorAll('input').forEach(inp => {
    inp.addEventListener('change', async () => {
      const nuevos = paramsActuales();
      nuevos[inp.dataset.param] = parseFloat(inp.value) || DEFAULT_PARAMS[inp.dataset.param];
      localStorage.setItem(PARAMS_KEY, JSON.stringify(nuevos));
      await procesarYRender(); // reprocesa con los nuevos tipos de cambio
    });
  });
}

// ---------- modo demo ----------
// Con carpeta muestras/ presente (desarrollo local) procesa esos archivos reales;
// si no existe (versión publicada), muestra el reporte con datos ficticios.

function demoSintetica() {
  resultados = resultadosDemo();
  $id('banner-demo').hidden = false;
  finalizarRender(paramsActuales());
}

async function cargarDemo() {
  try {
    progreso('Cargando muestras de prueba…');
    const resp = await fetch('muestras/manifest.json');
    if (!resp.ok) throw new Error('sin muestras');
    const manifest = await resp.json();
    archivosCrudos = [];
    for (const m of manifest) {
      const buf = await (await fetch('muestras/' + encodeURIComponent(m.archivo))).arrayBuffer();
      archivosCrudos.push({ nombre: m.archivo, ruta: m.ruta, arrayBuffer: buf });
    }
    await procesarYRender();
  } catch {
    demoSintetica();
  }
}

$id('btn-demo').addEventListener('click', cargarDemo);
if (new URLSearchParams(location.search).has('demo')) cargarDemo();
