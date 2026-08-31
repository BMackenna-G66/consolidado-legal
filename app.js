// Orquestador: carga archivos, los procesa con los parsers y pinta el reporte.

import { DEFAULT_PARAMS, MESES } from './config.js';
import { procesarArchivo, deduplicar, aplicarBaseHistorica, enriquecerSolicitantes, parseCargaManual } from './parsers.js';
import { leerCarpetaLocal, leerSharePoint, graphDisponible, soportaFSA, elegirCarpetaFSA, leerCarpetaRecordada, carpetaGuardada, setClientId, leerZip } from './fuentes.js';
import { generarHTMLCompartir } from './compartir.js';
import { construirPivot, renderPivot, renderResumenTarjetas, exportarCSV, exportarBaseMaestra, AGRUPACIONES } from './reporte.js';
import { clasificar, overrides, setOverride, CONCEPTOS, clasificarConIA, apiKeyGemini, setApiKeyGemini, normalizaCat } from './clasificador.js';
import { fichas, guardarFicha, borrarFicha, nuevoId, calcularMontos, resultadoManual,
         exportarFichasJSON, importarFichasJSON, exportarFichasXLSX, MONEDAS, PAISES, CONCEPTOS_FICHA } from './manual.js';

window.pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.worker.min.js';

const $id = id => document.getElementById(id);
const PARAMS_KEY = 'consolidado-legal-params';
const CARGA_KEY = 'consolidado-ultima-carga';

let resultados = [];   // salida por archivo (estado, nota, registros)
let archivosCrudos = null; // para reprocesar al cambiar parámetros
let origenCarga = '';  // de dónde salió la carga vigente (ZIP, carpeta, SharePoint)

// La última carga procesada queda guardada en el navegador: al volver a abrir
// la app el reporte aparece solo, sin re-subir el ZIP.
function guardarCarga() {
  try {
    localStorage.setItem(CARGA_KEY, JSON.stringify({
      fecha: new Date().toISOString(), origen: origenCarga, resultados: resultadosArchivos,
    }));
    return true;
  } catch {
    localStorage.removeItem(CARGA_KEY); // no cabe en el navegador: se sigue sin guardar
    return false;
  }
}

function cargaGuardada() {
  try {
    const d = JSON.parse(localStorage.getItem(CARGA_KEY) || 'null');
    return d && Array.isArray(d.resultados) && d.resultados.length ? d : null;
  } catch { return null; }
}

function renderBannerCarga(fecha, guardada) {
  const b = $id('banner-carga');
  b.hidden = false;
  const cuando = new Date(fecha).toLocaleString('es-CL', { dateStyle: 'short', timeStyle: 'short' });
  b.innerHTML = guardada
    ? `<span>💾 <b>${origenCarga || 'Carga'}</b> del ${cuando}, guardada en este navegador: al volver a abrir la app, el reporte aparece solo. Para actualizar, carga el nuevo ZIP.</span>
       <button class="boton secundario" id="btn-borrar-carga" type="button">Borrar datos guardados</button>`
    : `<span>⚠ <b>${origenCarga || 'Carga'}</b> del ${cuando} — no se pudo guardar en este navegador (datos muy grandes); al recargar la página habrá que subir el ZIP de nuevo.</span>`;
  const btn = $id('btn-borrar-carga');
  if (btn) btn.addEventListener('click', () => {
    if (!confirm('¿Borrar los datos guardados en este navegador? La app volverá a la portada vacía.')) return;
    localStorage.removeItem(CARGA_KEY);
    location.reload();
  });
}

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
    origenCarga = 'Carpeta local';
    archivosCrudos = archivos;
    await procesarYRender();
  } catch (e) {
    progreso(e.name === 'AbortError' ? '' : '⚠ ' + (e.message || e));
  }
});
$id('btn-zip').addEventListener('click', () => $id('input-zip').click());
$id('input-zip').addEventListener('change', async e => {
  if (!e.target.files.length) return;
  try {
    origenCarga = `ZIP "${e.target.files[0].name}"`;
    archivosCrudos = await leerZip(e.target.files[0], progreso);
    await procesarYRender();
  } catch (err) { progreso('⚠ ' + (err.message || err)); }
  e.target.value = '';
});

$id('input-carpeta').addEventListener('change', async e => {
  if (!e.target.files.length) return;
  progreso('Leyendo carpeta…');
  origenCarga = 'Carpeta local';
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
    origenCarga = `Carpeta "${dir.name}"`;
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
    origenCarga = 'SharePoint en vivo';
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
$id('btn-compartir').addEventListener('click', () => {
  const params = paramsActuales();
  const registros = [];
  for (const res of resultados) {
    const partes = String(res.ruta || '').split('/').filter(Boolean);
    for (const r of (res.registros || [])) {
      registros.push({
        anio: r.anio, mes: r.mes, dia: r.dia, pais: r.pais, proveedor: r.proveedor,
        solicitante: /^Varios \(/.test(r.solicitante || '') ? '' : (r.solicitante || ''),
        categoria: r.categoria, concepto: clasificar(r).concepto, moneda: r.moneda,
        montoOrigen: r.montoOrigen, clp: Math.round(r.clp),
        usd: r.usd != null ? r.usd : +(r.clp / params.USD_CLP).toFixed(2),
        detalle: r.detalle || '', archivo: r.archivo || '',
        carpeta: partes[0] || '(raíz)', fuente: r.fuente || 'documento',
      });
    }
  }
  const archivos = resultados.map(res => ({
    estado: res.estado, archivo: res.archivo, ruta: res.ruta || '',
    n: (res.registros || []).length,
    clp: Math.round((res.registros || []).reduce((s2, r) => s2 + r.clp, 0)),
    nota: res.nota || '',
  }));
  // Vínculo de la carpeta donde el equipo sube las planillas de carga manual:
  // se incrusta en el HTML generado (que vive en SharePoint), no en el repo.
  let carpetaUrl = localStorage.getItem('consolidado-share-url') || '';
  if (!carpetaUrl) {
    carpetaUrl = (prompt(
      'Pega el vínculo de la carpeta de SharePoint "Consolidado Cobros - Pagos [Compliance]".\n\n' +
      'Se incrusta en el reporte para que quien cargue un gasto pueda abrir la carpeta y subir\n' +
      'su planilla de un clic. Puedes dejarlo en blanco: el reporte igual se genera.'
    ) || '').trim();
    if (carpetaUrl) localStorage.setItem('consolidado-share-url', carpetaUrl);
  }
  const html = generarHTMLCompartir({ registros, archivos, params, carpetaUrl });
  window.__ultimoCompartir = { kb: Math.round(html.length / 1024), movimientos: registros.length };
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([html], { type: 'text/html;charset=utf-8' }));
  a.download = 'Reporte Consolidado Legal.html';
  a.click();
  URL.revokeObjectURL(a.href);
  progreso('Reporte generado: súbelo a la carpeta de SharePoint y comparte su vínculo.');
});
$id('btn-maestra').addEventListener('click', () => {
  const { movimientos } = exportarBaseMaestra(resultados, paramsActuales());
  progreso(`Base maestra generada con ${movimientos} movimientos.`);
});
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
  if (!archivosCrudos) {
    // Los datos vigentes vienen de la carga guardada: no hay archivos crudos que
    // reprocesar (p. ej. tras cambiar parámetros). Hay que volver a cargar el ZIP.
    progreso('Estos datos vienen de la carga guardada — para reprocesar (nuevos parámetros), vuelve a cargar el ZIP.');
    return;
  }
  const params = paramsActuales();
  progreso(`Interpretando ${archivosCrudos.length} archivos…`);
  resultadosArchivos = [];
  for (const a of archivosCrudos) resultadosArchivos.push(await procesarArchivo(a, params));
  deduplicar(resultadosArchivos);
  enriquecerSolicitantes(resultadosArchivos);
  aplicarBaseHistorica(resultadosArchivos);
  renderBannerCarga(new Date().toISOString(), guardarCarga());
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
  for (const b of ['btn-params', 'btn-mantenedor', 'btn-fichas', 'btn-maestra', 'btn-compartir', 'btn-csv', 'btn-recargar']) $id(b).hidden = false;
  window.__consolidado = { resultados, params }; // punto de acceso para exportaciones
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
  const archivo = e.target.files[0];
  try {
    if (/\.xlsx?$/i.test(archivo.name)) {
      // Planilla de carga manual que alguien envió por Teams o correo en vez de
      // subirla a SharePoint: entra igual, como fichas de este navegador.
      const { registros } = parseCargaManual(await archivo.arrayBuffer(), '', paramsActuales(), archivo.name);
      for (const r of registros) {
        guardarFicha({
          id: nuevoId(), mes: r.mes, anio: r.anio, dia: r.dia || 1,
          solicitante: r.solicitante || '', proveedor: r.proveedor, pais: r.pais,
          concepto: r.concepto, categoria: r.categoria, moneda: r.moneda,
          montoOrigen: r.montoOrigen, clpManual: r.clp,
          detalle: (r.detalle ? r.detalle + ' · ' : '') + `planilla ${archivo.name}`,
          registrado: new Date().toISOString(),
        });
      }
      $id('fi-msg').textContent = `✓ ${registros.length} movimiento(s) importados de la planilla`;
    } else {
      const { nuevas, actualizadas } = await importarFichasJSON(archivo);
      $id('fi-msg').textContent = `✓ ${nuevas} nueva(s), ${actualizadas} actualizada(s)`;
    }
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

// Inventario de archivos plegado tras un botón y agrupado por estado, cada
// estado en su propia tabla desplegable.
const ESTADOS_ARCHIVO = {
  error: ['❌ No leído', 'No se pudo extraer texto (escaneos sin capa de texto) — registrar a mano o pedir el PDF electrónico.'],
  supuesto: ['⚠️ Con supuestos', 'Leídos correctamente; el monto usa una conversión configurable (USD, UF, IGV).'],
  ok: ['✅ Procesado', 'Leídos sin supuestos.'],
  omitido: ['⏭ Omitido', 'Sin monto propio: anexos de horas, respaldos y duplicados que ya aporta otro documento.'],
};

function renderArchivos() {
  const cont = $id('lista-archivos');
  const abiertos = new Set([...cont.querySelectorAll('details[open]')].map(d => d.dataset.estado));
  const fmt = n => '$' + new Intl.NumberFormat('es-CL').format(Math.round(n));
  const clpDe = r => (r.registros || []).reduce((s, x) => s + x.clp, 0);
  let html = '';
  for (const est of Object.keys(ESTADOS_ARCHIVO)) {
    const lista = resultados.filter(r => r.estado === est);
    if (!lista.length) continue;
    const clp = lista.reduce((s, r) => s + clpDe(r), 0);
    const filas = lista.map(r => `<tr>
      <td class="etiqueta">${r.archivo}</td>
      <td class="etiqueta">${String(r.ruta || '').split('/').slice(0, -1).join('/') || '(raíz)'}</td>
      <td class="num">${(r.registros || []).length || ''}</td>
      <td class="num">${clpDe(r) ? fmt(clpDe(r)) : ''}</td>
      <td class="nota-archivo">${r.nota || ''}</td></tr>`).join('');
    html += `<details class="grupo-archivos" data-estado="${est}"${abiertos.has(est) ? ' open' : ''}>
      <summary><span class="estado ${est}">${ESTADOS_ARCHIVO[est][0]}</span><span class="conteo">${lista.length}</span>
        <span class="resumen-grupo">${clp ? fmt(clp) + ' aportados · ' : ''}${ESTADOS_ARCHIVO[est][1]}</span></summary>
      <div class="pivot-scroll"><table class="pivot"><thead><tr>
        <th class="etiqueta">Archivo</th><th class="etiqueta">Carpeta</th><th>Movs.</th><th>CLP</th><th class="etiqueta">Nota</th>
      </tr></thead><tbody>${filas}</tbody></table></div></details>`;
  }
  cont.innerHTML = html || '<div class="archivo-fila">Sin archivos.</div>';
  const nErr = resultados.filter(r => r.estado === 'error').length;
  const nSup = resultados.filter(r => r.estado === 'supuesto').length;
  $id('btn-archivos').textContent = `📄 Archivos leídos (${resultados.length})`
    + (nErr ? ` — ${nErr} sin leer` : '') + (nSup ? ` · ${nSup} con supuestos` : '');
}

$id('btn-archivos').addEventListener('click', () => {
  const l = $id('lista-archivos');
  l.hidden = !l.hidden;
});

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

// ---------- arranque ----------
// 1) Si hay una carga guardada de una visita anterior, el reporte aparece solo.
// 2) Si no, y el modo SharePoint en vivo está configurado, se lee de ahí.
// 3) Si no hay nada, la portada explica cómo cargar el ZIP.

(async () => {
  const previa = cargaGuardada();
  if (previa) {
    resultadosArchivos = previa.resultados;
    origenCarga = previa.origen || 'Carga anterior';
    renderBannerCarga(previa.fecha, true);
    finalizarRender(paramsActuales());
    return;
  }
  if (!graphDisponible()) {
    $id('aviso-sin-datos').hidden = false;
    $id('btn-zip').classList.add('destacado');
    $id('card-manual').classList.add('principal');
    $id('card-manual').classList.remove('secundaria');
    $id('card-sp').classList.remove('principal');
    $id('card-sp').classList.add('secundaria');
    return;
  }
  try {
    progreso('Conectando con SharePoint…');
    archivosCrudos = await leerSharePoint(progreso);
    await procesarYRender();
  } catch (e) {
    progreso('');
    $id('aviso-sin-datos').hidden = false;
    $id('btn-sharepoint').classList.add('destacado');
    if (!/popup|cancel|user_cancelled/i.test(String(e.message || e))) {
      $id('detalle-error').textContent = 'Último intento: ' + (e.message || e);
      $id('detalle-error').hidden = false;
    }
  }
})();
