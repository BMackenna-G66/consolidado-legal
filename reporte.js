// Motor del reporte: tabla dinámica con columnas por mes agrupadas por año y
// filas jerárquicas de profundidad configurable (país → proveedor → concepto → solicitante).

import { MESES } from './config.js';

const fmtCLP = new Intl.NumberFormat('es-CL', { maximumFractionDigits: 0 });
const $ = v => (v == null || v === 0) ? '' : '$' + fmtCLP.format(Math.round(v));

// Agrupaciones disponibles. La primera es la vista por defecto.
export const AGRUPACIONES = [
  { id: 'pais', nombre: 'País → Proveedor → Concepto → Solicitante', dims: ['pais', 'proveedor', 'categoria', 'solicitante'] },
  { id: 'solicitante', nombre: 'País → Solicitante → Proveedor → Concepto', dims: ['pais', 'solicitante', 'proveedor', 'categoria'] },
  { id: 'proveedor', nombre: 'Proveedor → Concepto → Solicitante', dims: ['proveedor', 'categoria', 'solicitante'] },
  { id: 'concepto', nombre: 'Concepto → País → Proveedor → Solicitante', dims: ['categoria', 'pais', 'proveedor', 'solicitante'] },
];

const VACIOS = {
  solicitante: 'Sin solicitante registrado',
  categoria: 'Sin concepto',
  proveedor: 'Sin proveedor',
  pais: 'Sin país',
};

const valorDim = (r, dim) => {
  const v = r[dim];
  return (v == null || String(v).trim() === '') ? (VACIOS[dim] || '(sin dato)') : String(v);
};

function nodoVacio() {
  return { meses: {}, porAnio: {}, total: 0, n: 0, hijos: new Map() };
}

function acumular(nodo, r) {
  const k = `${r.anio}-${r.mes}`;
  nodo.meses[k] = (nodo.meses[k] || 0) + r.clp;
  nodo.porAnio[r.anio] = (nodo.porAnio[r.anio] || 0) + r.clp;
  nodo.total += r.clp;
  nodo.n++;
}

export function construirPivot(registros, dims) {
  const dimensiones = dims && dims.length ? dims : AGRUPACIONES[0].dims;
  const anios = [...new Set(registros.map(r => r.anio).filter(Boolean))].sort();
  const columnas = [];
  for (const a of anios) {
    const meses = [...new Set(registros.filter(r => r.anio === a).map(r => r.mes).filter(Boolean))];
    if (!meses.length) continue;
    for (let m = Math.min(...meses); m <= Math.max(...meses); m++) columnas.push({ anio: a, mes: m });
    columnas.push({ anio: a, total: true });
  }

  const raiz = nodoVacio();
  for (const r of registros) {
    if (!r.anio || !r.mes) continue;
    acumular(raiz, r);
    let actual = raiz;
    for (const dim of dimensiones) {
      const clave = valorDim(r, dim);
      if (!actual.hijos.has(clave)) actual.hijos.set(clave, nodoVacio());
      actual = actual.hijos.get(clave);
      acumular(actual, r);
    }
  }
  return { anios, columnas, raiz, dimensiones };
}

export function renderPivot(contenedor, pivot, titulo) {
  const { anios, columnas, raiz, dimensiones } = pivot;

  const celdas = (n) => {
    let s = '';
    for (const c of columnas) {
      s += c.total
        ? `<td class="total-anio num">${$(n.porAnio[c.anio])}</td>`
        : `<td class="num">${$(n.meses[`${c.anio}-${c.mes}`])}</td>`;
    }
    return s + `<td class="total-general num">${$(n.total)}</td>`;
  };

  let html = `<div class="pivot-titulo">${titulo}</div><div class="pivot-scroll"><table class="pivot">`;
  html += '<thead><tr><th class="etiqueta"></th>';
  for (const a of anios) {
    const nMeses = columnas.filter(c => c.anio === a && !c.total).length;
    html += `<th colspan="${nMeses}" class="anio">${a}</th><th class="total-anio">Total ${a}</th>`;
  }
  html += '<th class="total-general">Total general</th></tr>';
  html += '<tr><th class="etiqueta">Etiquetas de fila</th>';
  for (const c of columnas) html += c.total ? '<th class="total-anio"></th>' : `<th>${MESES[c.mes - 1]}</th>`;
  html += '<th class="total-general"></th></tr></thead><tbody>';

  // Recorre el árbol en profundidad emitiendo una fila por nodo.
  // data-ruta ("0.2.1") permite plegar toda la descendencia por prefijo.
  const filas = [];
  (function recorrer(nodo, nivel, ruta) {
    const hijos = [...nodo.hijos.entries()].sort((a, b) => b[1].total - a[1].total);
    hijos.forEach(([clave, hijo], i) => {
      const rutaHijo = ruta ? `${ruta}.${i}` : String(i);
      const tieneHijos = hijo.hijos.size > 0;
      const flecha = tieneHijos ? '<span class="flecha">▸</span> ' : '';
      const dim = dimensiones[nivel];
      filas.push(
        `<tr class="lvl-${nivel} ${tieneHijos ? 'expandible' : ''} ${nivel > 0 ? 'oculto' : ''}" data-ruta="${rutaHijo}" data-nivel="${nivel}">` +
        `<td class="etiqueta" title="${dim}: ${clave} · ${hijo.n} movimiento(s)">${flecha}${clave}` +
        `<span class="conteo">${hijo.n}</span></td>${celdas(hijo)}</tr>`
      );
      if (tieneHijos) recorrer(hijo, nivel + 1, rutaHijo);
    });
  })(raiz, 0, '');

  html += filas.join('');
  html += `<tr class="fila-total"><td class="etiqueta">Total general</td>${celdas(raiz)}</tr>`;
  html += '</tbody></table></div>';
  contenedor.innerHTML = html;

  contenedor.querySelectorAll('tr.expandible').forEach(tr => {
    tr.addEventListener('click', () => {
      const ruta = tr.dataset.ruta;
      const abierto = tr.classList.toggle('abierto');
      const hijosDirectos = contenedor.querySelectorAll(`tr[data-ruta^="${ruta}."]`);
      hijosDirectos.forEach(h => {
        const profundidad = h.dataset.ruta.split('.').length - ruta.split('.').length;
        if (abierto && profundidad === 1) h.classList.remove('oculto');
        else if (!abierto) { h.classList.add('oculto'); h.classList.remove('abierto'); }
      });
    });
  });
}

export function renderResumenTarjetas(contenedor, registros) {
  const total = registros.reduce((s, r) => s + r.clp, 0);
  const porAnio = {}, porPais = {};
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

// Base maestra: un Excel con todo lo que la app leyó y calculó — cada movimiento
// con su trazabilidad al archivo de origen, más las tablas de resumen y el
// inventario de archivos. Es la foto completa del consolidado en un solo libro.
export function exportarBaseMaestra(resultados, params) {
  const XLSX = window.XLSX;
  const filas = [];
  for (const res of resultados) {
    const carpeta = String(res.ruta || '').split('/').slice(0, -1).join('/') || '(raíz)';
    for (const r of (res.registros || [])) {
      filas.push({
        'Año': r.anio, 'Mes': r.mes, 'Mes nombre': r.mes ? MESES[r.mes - 1] : '', 'Día': r.dia,
        'País': r.pais, 'Proveedor / abogado': r.proveedor, 'Solicitante': r.solicitante || '',
        'Concepto cobrado': r.categoria, 'Tipo': r.concepto,
        'Moneda origen': r.moneda, 'Monto origen': r.montoOrigen,
        'Monto CLP': r.clp,
        'Monto USD': r.usd != null ? r.usd : (params.USD_CLP ? +(r.clp / params.USD_CLP).toFixed(2) : ''),
        'Detalle': r.detalle || '',
        'Origen del dato': r.fuente === 'base' ? 'Base histórica Excel'
          : r.fuente === 'manual' ? 'Ficha manual'
          : r.fuente === 'plantilla' ? 'Plantilla de carga manual' : 'Documento leído',
        'Archivo': r.archivo || '', 'Carpeta': carpeta, 'Ruta completa': res.ruta || '',
        'Estado lectura': res.estado || '', 'Nota de lectura': res.nota || '',
        'Solicitantes detectados': (r.solicitantesLista || []).join(', '),
      });
    }
  }
  filas.sort((a, b) => (a['Año'] - b['Año']) || (a['Mes'] - b['Mes']) || String(a['País']).localeCompare(b['País']));

  const wb = XLSX.utils.book_new();
  const hoja = (nombre, datos, cols) => {
    const ws = XLSX.utils.json_to_sheet(datos);
    ws['!cols'] = (cols || Object.keys(datos[0] || {})).map(c => ({ wch: Math.min(46, Math.max(11, String(c).length + 4)) }));
    ws['!autofilter'] = { ref: ws['!ref'] };
    ws['!freeze'] = { xSplit: 0, ySplit: 1 };
    XLSX.utils.book_append_sheet(wb, ws, nombre);
  };

  hoja('Base maestra', filas);

  // Resúmenes: los mismos cortes que muestra el reporte en pantalla
  const suma = (clave) => {
    const m = new Map();
    for (const f of filas) {
      const k = f[clave] || '(sin dato)';
      const e = m.get(k) || { n: 0, clp: 0 };
      e.n++; e.clp += f['Monto CLP'] || 0;
      m.set(k, e);
    }
    return [...m.entries()].sort((a, b) => b[1].clp - a[1].clp)
      .map(([k, v]) => ({ [clave]: k, 'Movimientos': v.n, 'Total CLP': Math.round(v.clp) }));
  };
  hoja('Por país', suma('País'));
  hoja('Por proveedor', suma('Proveedor / abogado'));
  hoja('Por concepto', suma('Concepto cobrado'));
  hoja('Por solicitante', suma('Solicitante'));

  // País × mes, la tabla del resumen
  const periodos = [...new Set(filas.map(f => `${f['Año']}-${String(f['Mes']).padStart(2, '0')}`))].sort();
  const paises = [...new Set(filas.map(f => f['País']))];
  const pivot = paises.map(p => {
    const fila = { 'País': p };
    let total = 0;
    for (const per of periodos) {
      const v = filas.filter(f => f['País'] === p && `${f['Año']}-${String(f['Mes']).padStart(2, '0')}` === per)
        .reduce((s, f) => s + (f['Monto CLP'] || 0), 0);
      fila[per] = Math.round(v) || '';
      total += v;
    }
    fila['Total general'] = Math.round(total);
    return fila;
  });
  const totalFila = { 'País': 'Total general' };
  for (const per of periodos) totalFila[per] = Math.round(pivot.reduce((s, r) => s + (Number(r[per]) || 0), 0));
  totalFila['Total general'] = Math.round(pivot.reduce((s, r) => s + (Number(r['Total general']) || 0), 0));
  hoja('Resumen país x mes', [...pivot, totalFila]);

  // Inventario de archivos: qué se leyó, con qué supuestos y qué quedó fuera
  hoja('Archivos', resultados.map(res => ({
    'Estado': res.estado, 'Archivo': res.archivo, 'Ruta': res.ruta || '',
    'Movimientos aportados': (res.registros || []).length,
    'Monto CLP aportado': Math.round((res.registros || []).reduce((s, r) => s + r.clp, 0)),
    'Nota': res.nota || '',
  })));

  hoja('Parámetros', Object.entries(params).map(([k, v]) => ({ 'Parámetro': k, 'Valor': v })));

  const fecha = new Date().toISOString().slice(0, 10);
  XLSX.writeFile(wb, `Base maestra Consolidado Legal ${fecha}.xlsx`);
  return { movimientos: filas.length, archivos: resultados.length };
}

export function exportarCSV(registros) {
  const enc = ['Año', 'Mes', 'Día', 'País', 'Proveedor', 'Solicitante', 'Concepto cobrado', 'Tipo', 'Moneda', 'Monto origen', 'CLP', 'USD', 'Origen', 'Archivo', 'Detalle'];
  const filas = registros.map(r => [
    r.anio, r.mes, r.dia, r.pais, r.proveedor, r.solicitante || '', r.categoria, r.concepto,
    r.moneda, r.montoOrigen, r.clp, r.usd ?? '', r.fuente || 'documento', r.archivo,
    (r.detalle || '').replace(/[\r\n;]+/g, ' '),
  ]);
  const csv = [enc, ...filas].map(f => f.map(c => `"${String(c ?? '').replace(/"/g, '""')}"`).join(';')).join('\n');
  const blob = new Blob(['﻿' + csv], { type: 'text/csv;charset=utf-8' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = `consolidado_legal_${new Date().toISOString().slice(0, 10)}.csv`;
  a.click();
  URL.revokeObjectURL(a.href);
}
