// Orquestador: carga archivos, los procesa con los parsers y pinta el reporte.

import { DEFAULT_PARAMS, MESES } from './config.js';
import { procesarArchivo, deduplicar, aplicarBaseHistorica, enriquecerSolicitantes } from './parsers.js';
import { leerCarpetaLocal, leerSharePoint, graphDisponible, soportaFSA, elegirCarpetaFSA, leerCarpetaRecordada, carpetaGuardada, setClientId } from './fuentes.js';
import { construirPivot, renderPivot, renderResumenTarjetas, exportarCSV, AGRUPACIONES } from './reporte.js';
import { clasificar, overrides, setOverride, CONCEPTOS, clasificarConIA, apiKeyGemini, setApiKeyGemini, normalizaCat } from './clasificador.js';
import { fichas, guardarFicha, borrarFicha, nuevoId, calcularMontos, resultadoManual,
         exportarFichasJSON, importarFichasJSON, exportarFichasXLSX, MONEDAS, PAISES, CONCEPTOS_FICHA } from './manual.js';

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

// Carpeta recordada de una visita anterior: si el permiso sigue vigente se relee
// sola al abrir; si no, queda el botón para reactivarla con un clic.
(async () => {
  const dir = await carpetaGuardada();
  if (!dir) return;
  const btn = $id('btn-actualizar');
  btn.hidden = false;
  btn.textContent = `↻ Actualizar desde "${dir.name}"`;
  const releer = async () => {
    progreso('Releyendo carpeta…');
    const { archivos } = await leerCarpetaRecordada(progreso);
    archivosCrudos = archivos;
    await procesarYRender();
  };
  btn.addEventListener('click', () => releer().catch(e => progreso('⚠ ' + (e.message || e))));
  try {
    if ((await dir.queryPermission({ mode: 'read' })) === 'granted') await releer();
  } catch { /* sin permiso vigente: el usuario usa el botón */ }
})();

$id('btn-sharepoint').addEventListener('click', async () => {
  if (!graphDisponible()) {
    const id = prompt(
      'Modo en línea: pega el "Application (client) ID" del registro de Azure AD.\n\n' +
      'Es el identificador que entrega TI (ver SOLICITUD-TI.md). No es una contraseña:\n' +
      'tu sesión de Microsoft se inicia en la ventana oficial de Microsoft y esta app\n' +
      'nunca ve tus credenciales.'
    );
    if (!id || !id.trim()) { progreso('Modo en línea no configurado — puedes seguir usando la carpeta local.'); return; }
    setClientId(id);
  }
  try {
    progreso('Conectando con Microsoft…');
    archivosCrudos = await leerSharePoint(progreso);
    await procesarYRender();
  } catch (e) {
    progreso('⚠ ' + (e.message || e));
  }
});

// Vuelve a la portada para elegir otra carpeta (sin recaer en la carga automática)
$id('btn-recargar').addEventListener('click', () => {
  $id('vista-reporte').style.display = 'none';
  $id('portada').style.display = 'block';
  for (const b of ['btn-params', 'btn-mantenedor', 'btn-csv', 'btn-recargar']) $id(b).hidden = true;
  progreso('');
});
$id('btn-csv').addEventListener('click', () => exportarCSV(registrosFiltrados()));
$id('btn-params').addEventListener('click', () => $id('panel-params').classList.toggle('visible'));
$id('f-gastos').addEventListener('change', renderTodo);
$id('f-honorarios').addEventListener('change', renderTodo);
$id('f-juicios').addEventListener('change', renderTodo);
$id('f-anio').addEventListener('change', renderTodo);
$id('f-agrupacion').addEventListener('change', renderTodo);
$id('btn-expandir').addEventListener('click', () => {
  const filas = [...document.querySelectorAll('#pivot-principal tr[data-ruta]')];
  const cerrado = filas.some(f => f.classList.contains('oculto'));
  filas.forEach(f => {
    f.classList.toggle('oculto', !cerrado && f.dataset.nivel !== '0');
    f.classList.toggle('abierto', cerrado && f.classList.contains('expandible'));
  });
  $id('btn-expandir').textContent = cerrado ? '⌃ Colapsar todo' : '⌄ Expandir todo';
});
$id('f-mes').addEventListener('change', renderTodo);
$id('btn-mantenedor').addEventListener('click', () => $id('mantenedor').classList.toggle('seccion-oculta'));
$id('btn-fichas').addEventListener('click', () => {
  const s = $id('fichas');
  s.classList.toggle('seccion-oculta');
  if (!s.classList.contains('seccion-oculta')) s.scrollIntoView({ behavior: 'smooth', block: 'start' });
});

function progreso(msg) { $id('progreso').textContent = msg; }

// ---------- pipeline ----------

let resultadosArchivos = []; // solo lo leído de documentos, sin las fichas manuales

async function procesarYRender() {
  const params = paramsActuales();
  progreso(`Interpretando ${archivosCrudos.length} archivos…`);
  resultadosArchivos = [];
  for (const a of archivosCrudos) resultadosArchivos.push(await procesarArchivo(a, params));
  deduplicar(resultadosArchivos);
  enriquecerSolicitantes(resultadosArchivos);
  aplicarBaseHistorica(resultadosArchivos);
  finalizarRender(params);
}

// Une los documentos leídos con las fichas de carga manual
function componerResultados(params) {
  const manual = resultadoManual(params);
  resultados = manual ? [...resultadosArchivos, manual] : [...resultadosArchivos];
}

function finalizarRender(params) {
  progreso('');
  componerResultados(params);
  $id('portada').style.display = 'none';
  $id('vista-reporte').style.display = 'block';
  for (const b of ['btn-params', 'btn-mantenedor', 'btn-fichas', 'btn-csv', 'btn-recargar']) $id(b).hidden = false;
  montarPanelParams(params);
  montarFormularioFicha(params);
  montarSelectorAgrupacion();
  montarFiltrosFecha();
  renderTodo();
}

// Recalcula el reporte tras cambiar las fichas manuales, sin releer los documentos
function refrescarConFichas() {
  const params = paramsActuales();
  componerResultados(params);
  montarFiltrosFecha();
  renderTodo();
  renderFichas(params);
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
function montarSelectorAgrupacion() {
  const sel = $id('f-agrupacion');
  if (sel.options.length) return;
  sel.innerHTML = AGRUPACIONES.map(a => `<option value="${a.id}">${a.nombre}</option>`).join('');
}

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
  const agr = AGRUPACIONES.find(a => a.id === $id('f-agrupacion').value) || AGRUPACIONES[0];
  renderPivot($id('pivot-principal'), construirPivot(regs, agr.dims), 'Resumen de gastos por conceptos legales / administrativos (CLP)');
  renderMantenedor();
  renderArchivos();
}

// ---------- ficha de carga manual ----------

const fmtCLP = n => '$' + new Intl.NumberFormat('es-CL').format(Math.round(n || 0));

function montarFormularioFicha(params) {
  const opciones = (sel, valores) => { $id(sel).innerHTML = valores.map(v => `<option>${v}</option>`).join(''); };
  $id('fi-mes').innerHTML = MESES.map((m, i) => `<option value="${i + 1}">${m}</option>`).join('');
  opciones('fi-pais', PAISES);
  opciones('fi-concepto', CONCEPTOS_FICHA);
  opciones('fi-moneda', MONEDAS);
  const hoy = new Date();
  if (!$id('fi-anio').value) { $id('fi-anio').value = hoy.getFullYear(); $id('fi-mes').value = hoy.getMonth() + 1; }

  // Sugerencias tomadas de lo que ya existe en el reporte
  const regs = resultados.flatMap(r => r.registros);
  $id('lista-proveedores').innerHTML = [...new Set(regs.map(r => r.proveedor))].filter(Boolean).map(p => `<option value="${p}">`).join('');
  $id('lista-categorias').innerHTML = [...new Set(regs.map(r => r.categoria))].filter(Boolean).slice(0, 60).map(c => `<option value="${c}">`).join('');

  const recalcular = () => {
    const p = paramsActuales();
    const { clp, usd } = calcularMontos({
      moneda: $id('fi-moneda').value,
      montoOrigen: $id('fi-monto').value,
      clpManual: $id('fi-clp').value,
    }, p);
    $id('fi-usd').value = usd ? 'US$ ' + new Intl.NumberFormat('es-CL', { minimumFractionDigits: 2 }).format(usd) : '';
    const tasa = $id('fi-moneda').value === 'CLP' ? null : calcularMontos({ moneda: $id('fi-moneda').value, montoOrigen: 1 }, p).clp;
    $id('fi-ayuda-clp').textContent = tasa ? `sugerido: ${fmtCLP(Number($id('fi-monto').value || 0) * tasa)} (1 ${$id('fi-moneda').value} = ${tasa})` : 'igual al monto de origen';
    if (!$id('fi-clp').value) $id('fi-clp').placeholder = String(clp || 0);
  };
  ['fi-monto', 'fi-moneda', 'fi-clp'].forEach(id => $id(id).addEventListener('input', recalcular));
  $id('fi-moneda').addEventListener('change', recalcular);
  recalcular();
  renderFichas(params);
}

function limpiarFicha() {
  for (const id of ['fi-id', 'fi-solicitante', 'fi-proveedor', 'fi-categoria', 'fi-monto', 'fi-clp', 'fi-detalle']) $id(id).value = '';
  $id('fi-usd').value = '';
  $id('fi-guardar').textContent = 'Agregar al reporte';
  $id('fi-msg').textContent = '';
}

$id('form-ficha').addEventListener('submit', e => {
  e.preventDefault();
  const ficha = {
    id: $id('fi-id').value || nuevoId(),
    mes: +$id('fi-mes').value, anio: +$id('fi-anio').value, dia: 1,
    solicitante: $id('fi-solicitante').value.trim(),
    proveedor: $id('fi-proveedor').value.trim(),
    pais: $id('fi-pais').value,
    concepto: $id('fi-concepto').value,
    categoria: $id('fi-categoria').value.trim() || 'Carga manual',
    moneda: $id('fi-moneda').value,
    montoOrigen: Number($id('fi-monto').value),
    clpManual: $id('fi-clp').value === '' ? null : Number($id('fi-clp').value),
    detalle: $id('fi-detalle').value.trim(),
    registrado: new Date().toISOString(),
  };
  guardarFicha(ficha);
  limpiarFicha();
  $id('fi-msg').textContent = '✓ Ficha guardada y sumada al reporte';
  setTimeout(() => { $id('fi-msg').textContent = ''; }, 4000);
  refrescarConFichas();
});

$id('fi-limpiar').addEventListener('click', limpiarFicha);
$id('fi-exportar').addEventListener('click', exportarFichasJSON);
$id('fi-excel').addEventListener('click', () => {
  try {
    const n = exportarFichasXLSX(paramsActuales());
    $id('fi-msg').textContent = `✓ Excel con ${n} movimiento(s) — súbelo a la carpeta del proveedor en SharePoint`;
  } catch (e) { $id('fi-msg').textContent = '⚠ ' + (e.message || e); }
});
$id('fi-importar-btn').addEventListener('click', () => $id('fi-importar').click());
$id('fi-importar').addEventListener('change', async e => {
  if (!e.target.files.length) return;
  try {
    const { nuevas, actualizadas } = await importarFichasJSON(e.target.files[0]);
    $id('fi-msg').textContent = `✓ ${nuevas} nueva(s), ${actualizadas} actualizada(s)`;
    refrescarConFichas();
  } catch (err) { $id('fi-msg').textContent = '⚠ ' + (err.message || err); }
  e.target.value = '';
});

function renderFichas(params) {
  const cont = $id('tabla-fichas');
  if (!cont) return;
  const lista = fichas();
  if (!lista.length) { cont.innerHTML = '<div class="archivo-fila">Aún no hay fichas cargadas manualmente.</div>'; return; }
  const filas = lista
    .slice()
    .sort((a, b) => (b.anio - a.anio) || (b.mes - a.mes))
    .map(f => {
      const { clp, usd } = calcularMontos(f, params);
      return `<tr>
        <td class="etiqueta">${MESES[(f.mes || 1) - 1]} ${f.anio}</td>
        <td>${f.solicitante || ''}</td>
        <td>${f.proveedor || ''}</td>
        <td>${f.pais}</td>
        <td>${f.concepto}</td>
        <td class="num">${f.moneda} ${new Intl.NumberFormat('es-CL').format(f.montoOrigen)}</td>
        <td class="num">${fmtCLP(clp)}</td>
        <td class="num">US$ ${new Intl.NumberFormat('es-CL', { minimumFractionDigits: 2 }).format(usd)}</td>
        <td><div class="acciones-fila"><button data-editar="${f.id}">Editar</button><button class="borrar" data-borrar="${f.id}">Borrar</button></div></td>
      </tr>`;
    }).join('');
  cont.innerHTML = `<table class="pivot"><thead><tr>
    <th class="etiqueta">Período</th><th>Solicitante</th><th>Proveedor</th><th>País</th>
    <th>Concepto</th><th>Monto origen</th><th>CLP</th><th>USD</th><th></th>
  </tr></thead><tbody>${filas}</tbody></table>`;

  cont.querySelectorAll('[data-borrar]').forEach(b => b.addEventListener('click', () => {
    if (!confirm('¿Borrar esta ficha del reporte?')) return;
    borrarFicha(b.dataset.borrar);
    refrescarConFichas();
  }));
  cont.querySelectorAll('[data-editar]').forEach(b => b.addEventListener('click', () => {
    const f = fichas().find(x => x.id === b.dataset.editar);
    if (!f) return;
    $id('fi-id').value = f.id; $id('fi-mes').value = f.mes; $id('fi-anio').value = f.anio;
    $id('fi-solicitante').value = f.solicitante || ''; $id('fi-proveedor').value = f.proveedor || '';
    $id('fi-pais').value = f.pais; $id('fi-concepto').value = f.concepto;
    $id('fi-categoria').value = f.categoria || ''; $id('fi-moneda').value = f.moneda;
    $id('fi-monto').value = f.montoOrigen; $id('fi-clp').value = f.clpManual ?? '';
    $id('fi-detalle').value = f.detalle || '';
    $id('fi-guardar').textContent = 'Guardar cambios';
    $id('fi-monto').dispatchEvent(new Event('input'));
    $id('fichas').scrollIntoView({ behavior: 'smooth', block: 'start' });
  }));
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

// ---------- carga automática de la carpeta de datos local ----------
// Si existe datos/manifest.json (copia de la carpeta de SharePoint sincronizada),
// el reporte se genera solo al abrir la app, sin ningún clic.

async function cargarDatosLocales() {
  const resp = await fetch('datos/manifest.json', { cache: 'no-store' });
  if (!resp.ok) throw new Error('sin carpeta de datos');
  const manifest = await resp.json();
  archivosCrudos = [];
  let i = 0;
  for (const m of manifest) {
    progreso(`Leyendo ${++i}/${manifest.length}: ${m.archivo.split('/').pop()}`);
    const r = await fetch('datos/' + m.archivo.split('/').map(encodeURIComponent).join('/'));
    if (!r.ok) continue;
    archivosCrudos.push({ nombre: m.archivo.split('/').pop(), ruta: m.ruta, arrayBuffer: await r.arrayBuffer() });
  }
  await procesarYRender();
}

(async () => {
  try {
    await cargarDatosLocales();
  } catch {
    // Sin carpeta datos/ (típico de la versión publicada): hay que decirlo, o la
    // portada queda muda y parece que la app no hace nada.
    const aviso = $id('aviso-sin-datos');
    if (aviso) aviso.hidden = false;
    $id('btn-local').classList.add('destacado');
  }
})();
