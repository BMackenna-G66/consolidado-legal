// Motor del reporte: pivot país × mes (columnas agrupadas por año, totales por año
// y total general), con filas expandibles por proveedor y categoría.

import { MESES } from './config.js';

const fmtCLP = new Intl.NumberFormat('es-CL', { maximumFractionDigits: 0 });
const $ = v => (v == null || v === 0) ? '' : '$' + fmtCLP.format(Math.round(v));

export function construirPivot(registros) {
  // Estructura: paises → { total, meses{ 'anio-mes': v }, hijos: proveedor → categoria }
  const anios = [...new Set(registros.map(r => r.anio).filter(Boolean))].sort();
  const columnas = [];
  for (const a of anios) {
    const mesesConDatos = [...new Set(registros.filter(r => r.anio === a).map(r => r.mes).filter(Boolean))];
    const desde = Math.min(...mesesConDatos), hasta = Math.max(...mesesConDatos);
    for (let m = desde; m <= hasta; m++) columnas.push({ anio: a, mes: m });
    columnas.push({ anio: a, total: true });
  }

  const arbol = new Map();
  for (const r of registros) {
    if (!r.anio || !r.mes) continue;
    const kMes = `${r.anio}-${r.mes}`;
    const pais = nodo(arbol, r.pais);
    suma(pais, kMes, r.clp, r.anio);
    const prov = nodo(pais.hijos, r.proveedor);
    suma(prov, kMes, r.clp, r.anio);
    const cat = nodo(prov.hijos, r.categoria || 'Sin categoría');
    suma(cat, kMes, r.clp, r.anio);
  }
  return { anios, columnas, arbol };
}

function nodo(mapa, clave) {
  if (!mapa.has(clave)) mapa.set(clave, { meses: {}, porAnio: {}, total: 0, hijos: new Map() });
  return mapa.get(clave);
}
function suma(n, kMes, v, anio) {
  n.meses[kMes] = (n.meses[kMes] || 0) + v;
  n.porAnio[anio] = (n.porAnio[anio] || 0) + v;
  n.total += v;
}

export function renderPivot(contenedor, pivot, titulo) {
  const { anios, columnas, arbol } = pivot;
  const totalGeneral = { meses: {}, porAnio: {}, total: 0 };
  for (const [, p] of arbol) {
    for (const k in p.meses) totalGeneral.meses[k] = (totalGeneral.meses[k] || 0) + p.meses[k];
    for (const a in p.porAnio) totalGeneral.porAnio[a] = (totalGeneral.porAnio[a] || 0) + p.porAnio[a];
    totalGeneral.total += p.total;
  }

  let html = `<div class="pivot-titulo">${titulo}</div><div class="pivot-scroll"><table class="pivot">`;
  // Fila de años
  html += '<thead><tr><th class="etiqueta"></th>';
  for (const a of anios) {
    const nMeses = columnas.filter(c => c.anio === a && !c.total).length;
    html += `<th colspan="${nMeses}" class="anio">${a}</th><th class="total-anio">Total ${a}</th>`;
  }
  html += '<th class="total-general">Total general</th></tr>';
  // Fila de meses
  html += '<tr><th class="etiqueta">Etiquetas de fila</th>';
  for (const c of columnas) html += c.total ? '<th class="total-anio"></th>' : `<th>${MESES[c.mes - 1]}</th>`;
  html += '<th class="total-general"></th></tr></thead><tbody>';

  const filaCeldas = (n) => {
    let s = '';
    for (const c of columnas) {
      s += c.total
        ? `<td class="total-anio num">${$(n.porAnio[c.anio])}</td>`
        : `<td class="num">${$(n.meses[`${c.anio}-${c.mes}`])}</td>`;
    }
    s += `<td class="total-general num">${$(n.total)}</td>`;
    return s;
  };

  const paises = [...arbol.entries()].sort((a, b) => b[1].total - a[1].total);
  let uid = 0;
  for (const [pais, nP] of paises) {
    const idP = 'g' + (uid++);
    html += `<tr class="nivel-pais expandible" data-grupo="${idP}"><td class="etiqueta"><span class="flecha">▸</span> ${pais}</td>${filaCeldas(nP)}</tr>`;
    for (const [prov, nProv] of [...nP.hijos.entries()].sort((a, b) => b[1].total - a[1].total)) {
      const idProv = 'g' + (uid++);
      html += `<tr class="nivel-prov oculto expandible" data-padre="${idP}" data-grupo="${idProv}"><td class="etiqueta"><span class="flecha">▸</span> ${prov}</td>${filaCeldas(nProv)}</tr>`;
      for (const [cat, nCat] of [...nProv.hijos.entries()].sort((a, b) => b[1].total - a[1].total)) {
        html += `<tr class="nivel-cat oculto" data-padre="${idProv}" data-ancestro="${idP}"><td class="etiqueta">${cat}</td>${filaCeldas(nCat)}</tr>`;
      }
    }
  }
  html += `<tr class="fila-total"><td class="etiqueta">Total general</td>${filaCeldas(totalGeneral)}</tr>`;
  html += '</tbody></table></div>';
  contenedor.innerHTML = html;

  // Expandir/colapsar
  contenedor.querySelectorAll('tr.expandible').forEach(tr => {
    tr.addEventListener('click', () => {
      const grupo = tr.dataset.grupo;
      const abierto = tr.classList.toggle('abierto');
      contenedor.querySelectorAll(`tr[data-padre="${grupo}"]`).forEach(h => {
        h.classList.toggle('oculto', !abierto);
        if (!abierto && h.dataset.grupo) {
          h.classList.remove('abierto');
          contenedor.querySelectorAll(`tr[data-padre="${h.dataset.grupo}"]`).forEach(n => n.classList.add('oculto'));
        }
      });
      if (!abierto) contenedor.querySelectorAll(`tr[data-ancestro="${grupo}"]`).forEach(n => n.classList.add('oculto'));
    });
  });
}

export function renderResumenTarjetas(contenedor, registros) {
  const total = registros.reduce((s, r) => s + r.clp, 0);
  const porAnio = {};
  const porPais = {};
  for (const r of registros) {
    if (r.anio) porAnio[r.anio] = (porAnio[r.anio] || 0) + r.clp;
    porPais[r.pais] = (porPais[r.pais] || 0) + r.clp;
  }
  const tarjeta = (t, v, sub = '') => `<div class="tarjeta"><div class="tarjeta-titulo">${t}</div><div class="tarjeta-valor">${$(v) || '$0'}</div><div class="tarjeta-sub">${sub}</div></div>`;
  let html = tarjeta('Total general', total, `${registros.length} movimientos`);
  for (const a of Object.keys(porAnio).sort()) html += tarjeta(`Total ${a}`, porAnio[a]);
  const lider = Object.entries(porPais).sort((x, y) => y[1] - x[1])[0];
  if (lider) html += tarjeta('País con mayor gasto', lider[1], lider[0]);
  contenedor.innerHTML = html;
}

export function exportarCSV(registros) {
  const enc = ['Año', 'Mes', 'Día', 'País', 'Proveedor', 'Categoría', 'Concepto', 'Moneda', 'Monto origen', 'CLP', 'Archivo', 'Detalle'];
  const filas = registros.map(r => [r.anio, r.mes, r.dia, r.pais, r.proveedor, r.categoria, r.concepto, r.moneda, r.montoOrigen, r.clp, r.archivo, (r.detalle || '').replace(/[\r\n;]+/g, ' ')]);
  const csv = [enc, ...filas].map(f => f.map(c => `"${String(c ?? '').replace(/"/g, '""')}"`).join(';')).join('\n');
  const blob = new Blob(['﻿' + csv], { type: 'text/csv;charset=utf-8' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = `consolidado_legal_${new Date().toISOString().slice(0, 10)}.csv`;
  a.click();
  URL.revokeObjectURL(a.href);
}
