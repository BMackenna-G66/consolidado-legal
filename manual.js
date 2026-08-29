// Carga manual tipo ficha (ERP): movimientos que no vienen de un documento en las
// carpetas — o que vienen de un PDF que no se puede leer automáticamente.
// Se guardan en el navegador y se suman al reporte como una fuente más.

import { DEFAULT_PARAMS, MESES } from './config.js';

const KEY = 'consolidado-manual-v1';

export const MONEDAS = ['CLP', 'USD', 'PEN', 'COP', 'ARS', 'UF'];
export const PAISES = ['Chile', 'Colombia', 'Perú', 'Argentina'];
export const CONCEPTOS_FICHA = ['Gastos', 'Honorarios', 'Juicios y otros'];

export function fichas() {
  try { return JSON.parse(localStorage.getItem(KEY) || '[]'); } catch { return []; }
}

function guardar(lista) {
  localStorage.setItem(KEY, JSON.stringify(lista));
}

export function nuevoId() {
  return 'M' + Date.now().toString(36) + Math.floor(Math.random() * 1e4).toString(36);
}

export function guardarFicha(ficha) {
  const lista = fichas();
  const i = lista.findIndex(f => f.id === ficha.id);
  if (i >= 0) lista[i] = ficha; else lista.push(ficha);
  guardar(lista);
  return ficha;
}

export function borrarFicha(id) {
  guardar(fichas().filter(f => f.id !== id));
}

// Tasa a CLP de cada moneda según los parámetros vigentes.
export function tasaCLP(moneda, params) {
  const p = { ...DEFAULT_PARAMS, ...params };
  return { CLP: 1, USD: p.USD_CLP, PEN: p.PEN_CLP, COP: p.COP_CLP, ARS: p.ARS_CLP, UF: p.UF_CLP }[moneda] ?? 1;
}

// Cálculo derivado: CLP desde la moneda de origen y su equivalente en USD.
export function calcularMontos({ moneda, montoOrigen, clpManual }, params) {
  const p = { ...DEFAULT_PARAMS, ...params };
  const origen = Number(montoOrigen) || 0;
  const clp = clpManual != null && clpManual !== '' ? Number(clpManual) : Math.round(origen * tasaCLP(moneda, p));
  const usd = p.USD_CLP ? +(clp / p.USD_CLP).toFixed(2) : 0;
  return { clp, usd };
}

// Convierte las fichas en registros con la misma forma que los de los parsers,
// para que entren al pivot, al mantenedor y a la exportación sin casos especiales.
export function registrosDeFichas(params) {
  return fichas().map(f => {
    const { clp, usd } = calcularMontos(f, params);
    return {
      dia: f.dia || 1, mes: f.mes, anio: f.anio,
      pais: f.pais, proveedor: f.proveedor || 'Sin proveedor',
      categoria: f.categoria || 'Carga manual',
      concepto: f.concepto || 'Gastos',
      moneda: f.moneda, montoOrigen: Number(f.montoOrigen) || 0,
      clp, usd, archivo: '(carga manual)',
      detalle: [f.solicitante ? 'Solicita: ' + f.solicitante : '', f.detalle || ''].filter(Boolean).join(' — '),
      fuente: 'manual', idFicha: f.id, solicitante: f.solicitante || '',
    };
  });
}

// Resultado con forma de "archivo procesado" para el panel de archivos.
export function resultadoManual(params) {
  const regs = registrosDeFichas(params);
  if (!regs.length) return null;
  return {
    archivo: 'Carga manual', ruta: '(fichas ingresadas a mano)',
    estado: 'ok', fuente: 'manual',
    nota: `${regs.length} ficha(s) ingresada(s) manualmente`,
    registros: regs,
  };
}

export function exportarFichasJSON() {
  const blob = new Blob([JSON.stringify(fichas(), null, 1)], { type: 'application/json' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = `fichas_manuales_${new Date().toISOString().slice(0, 10)}.json`;
  a.click();
  URL.revokeObjectURL(a.href);
}

export async function importarFichasJSON(file) {
  const texto = await file.text();
  const entrantes = JSON.parse(texto);
  if (!Array.isArray(entrantes)) throw new Error('El archivo no contiene una lista de fichas');
  const lista = fichas();
  let nuevas = 0, actualizadas = 0;
  for (const f of entrantes) {
    if (!f || !f.id) continue;
    const i = lista.findIndex(x => x.id === f.id);
    if (i >= 0) { lista[i] = f; actualizadas++; } else { lista.push(f); nuevas++; }
  }
  guardar(lista);
  return { nuevas, actualizadas };
}

export { MESES };
