// Genera la base maestra en Excel ejecutando los MISMOS lectores de la app,
// pero en Node, para no depender del navegador.
import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';

const require = createRequire(import.meta.url);
const XLSX = require('xlsx');
const pdfjs = require('pdfjs-dist/legacy/build/pdf.js');
pdfjs.GlobalWorkerOptions.workerSrc = require.resolve('pdfjs-dist/legacy/build/pdf.worker.js');

// Raíz del proyecto (este script vive en scripts/)
const APP = path.resolve(path.dirname(new URL(import.meta.url).pathname), '..');

// Shims: los módulos de la app esperan un navegador
globalThis.window = { XLSX, pdfjsLib: pdfjs };
globalThis.localStorage = { getItem: () => null, setItem() {}, removeItem() {} };

const { procesarArchivo, deduplicar, aplicarBaseHistorica, enriquecerSolicitantes } = await import(path.join(APP, 'parsers.js'));
const { DEFAULT_PARAMS, MESES } = await import(path.join(APP, 'config.js'));
const { extraerDetalle } = await import(path.join(APP, 'detalle.js'));

const manifest = JSON.parse(fs.readFileSync(path.join(APP, 'datos/manifest.json'), 'utf8'));
const params = { ...DEFAULT_PARAMS };

const resultados = [];
const detalle = [];
for (const m of manifest) {
  const full = path.join(APP, 'datos', m.archivo);
  const buf = fs.readFileSync(full);
  const arrayBuffer = buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
  const nombre = m.archivo.split('/').pop();
  const partes = m.ruta.split('/').filter(Boolean);
  resultados.push(await procesarArchivo({ nombre, ruta: m.ruta, arrayBuffer }, params));
  const lineas = await extraerDetalle({ nombre, ruta: m.ruta, arrayBuffer }, {
    carpetaProveedor: partes[0] || '(raíz)',
    carpetaMes: partes.length > 2 ? partes[partes.length - 2] : '',
  });
  detalle.push(...lineas);
}
deduplicar(resultados);
enriquecerSolicitantes(resultados);
aplicarBaseHistorica(resultados);

// ---- armado del libro ----
const filas = [];
for (const res of resultados) {
  const carpeta = String(res.ruta || '').split('/').slice(0, -1).join('/') || '(raíz)';
  const partes = String(res.ruta || '').split('/').filter(Boolean);
  for (const r of (res.registros || [])) {
    filas.push({
      'Año': r.anio ?? '', 'Mes': r.mes ?? '', 'Mes nombre': r.mes ? MESES[r.mes - 1] : '', 'Día': r.dia ?? '',
      'Periodo': (r.anio && r.mes) ? `${r.anio}-${String(r.mes).padStart(2, '0')}` : '',
      'País': r.pais, 'Proveedor / abogado': r.proveedor,
      'Solicitante': /^Varios \(/.test(r.solicitante || '') ? '' : (r.solicitante || ''),
      'Concepto cobrado': r.categoria, 'Tipo': r.concepto,
      'Moneda origen': r.moneda, 'Monto origen': Number(r.montoOrigen) || 0,
      'Monto CLP': Math.round(r.clp),
      'Monto USD': r.usd != null ? r.usd : +(r.clp / params.USD_CLP).toFixed(2),
      'Detalle': r.detalle || '',
      'Origen del dato': r.fuente === 'base' ? 'Base histórica Excel'
        : r.fuente === 'manual' ? 'Ficha manual'
        : r.fuente === 'plantilla' ? 'Plantilla de carga manual' : 'Documento leído',
      'Carpeta proveedor': partes[0] || '(raíz)',
      'Carpeta mes': partes.length > 2 ? partes[partes.length - 2] : '',
      'Carpeta completa': carpeta,
      'Archivo': r.archivo || '', 'Ruta completa': res.ruta || '',
      'Estado lectura': res.estado || '', 'Nota de lectura': res.nota || '',
      'Solicitantes detectados': (r.solicitantesLista || []).join(', '),
    });
  }
}
filas.sort((a, b) => String(a['Periodo']).localeCompare(String(b['Periodo']))
  || String(a['País']).localeCompare(String(b['País']))
  || String(a['Proveedor / abogado']).localeCompare(String(b['Proveedor / abogado'])));

const wb = XLSX.utils.book_new();
const agregar = (nombre, datos) => {
  const ws = XLSX.utils.json_to_sheet(datos);
  const cols = Object.keys(datos[0] || {});
  ws['!cols'] = cols.map(c => ({ wch: Math.min(48, Math.max(11, c.length + 3)) }));
  if (ws['!ref']) ws['!autofilter'] = { ref: ws['!ref'] };
  ws['!freeze'] = { xSplit: 0, ySplit: 1, topLeftCell: 'A2', activePane: 'bottomLeft', state: 'frozen' };
  XLSX.utils.book_append_sheet(wb, ws, nombre);
};

agregar('Base maestra', filas);

const suma = (clave) => {
  const m = new Map();
  for (const f of filas) {
    const k = f[clave] || '(sin dato)';
    const e = m.get(k) || { n: 0, clp: 0, usd: 0 };
    e.n++; e.clp += f['Monto CLP']; e.usd += f['Monto USD'];
    m.set(k, e);
  }
  return [...m.entries()].sort((a, b) => b[1].clp - a[1].clp)
    .map(([k, v]) => ({ [clave]: k, 'Movimientos': v.n, 'Total CLP': Math.round(v.clp), 'Total USD': +v.usd.toFixed(2) }));
};
// Detalle línea por línea de los documentos que lo traen adentro
const filasDetalle = detalle.map(d => ({
  'Periodo': (d.anio && d.mes) ? `${d.anio}-${String(d.mes).padStart(2, '0')}` : '',
  'Fecha': d.fecha || '', 'Año': d.anio ?? '', 'Mes': d.mes ?? '', 'Día': d.dia ?? '',
  'Carpeta proveedor': d.carpetaProveedor || '', 'Carpeta mes': d.carpetaMes || '',
  'Categoría': d.categoria || '', 'Solicitante': d.solicitante || '',
  'Profesional / abogado': d.profesional || '',
  'Descripción del trabajo': d.descripcion || '',
  'Horas': d.horas || '', 'Tarifa': d.tarifa ?? '',
  'Valor': d.valor ?? '', 'Moneda del valor': d.monedaValor || '', 'Unidad': d.unidad || '',
  'N° documento': d.documento || '', 'Hoja': d.hoja || '', 'Tipo documento': d.tipoDocumento || '',
  'Archivo': d.archivo || '', 'Ruta completa': d.ruta || '',
}));
filasDetalle.sort((a, b) => String(a['Periodo']).localeCompare(String(b['Periodo']))
  || String(a['Carpeta proveedor']).localeCompare(String(b['Carpeta proveedor'])));
if (filasDetalle.length) agregar('Detalle línea por línea', filasDetalle);

agregar('Por país', suma('País'));
agregar('Por proveedor', suma('Proveedor / abogado'));
agregar('Por concepto', suma('Concepto cobrado'));
agregar('Por solicitante', suma('Solicitante'));
agregar('Por tipo', suma('Tipo'));
agregar('Por origen del dato', suma('Origen del dato'));

// País × mes, igual que el resumen en pantalla
const periodos = [...new Set(filas.map(f => f['Periodo']).filter(Boolean))].sort();
const paises = [...new Set(filas.map(f => f['País']))].sort();
const pivot = paises.map(p => {
  const fila = { 'País': p };
  let total = 0;
  for (const per of periodos) {
    const v = filas.filter(f => f['País'] === p && f['Periodo'] === per).reduce((s, f) => s + f['Monto CLP'], 0);
    fila[per] = v ? Math.round(v) : '';
    total += v;
  }
  fila['Total general'] = Math.round(total);
  return fila;
});
const totalGeneral = { 'País': 'TOTAL GENERAL' };
for (const per of periodos) totalGeneral[per] = Math.round(pivot.reduce((s, r) => s + (Number(r[per]) || 0), 0));
totalGeneral['Total general'] = Math.round(pivot.reduce((s, r) => s + r['Total general'], 0));
agregar('Resumen país x mes', [...pivot, totalGeneral]);

agregar('Archivos leídos', resultados.map(res => ({
  'Estado': res.estado, 'Archivo': res.archivo, 'Ruta': res.ruta || '',
  'Movimientos': (res.registros || []).length,
  'Monto CLP aportado': Math.round((res.registros || []).reduce((s, r) => s + r.clp, 0)),
  'Nota': res.nota || '',
})));

agregar('Parámetros', [
  ...Object.entries(params).map(([k, v]) => ({ 'Parámetro': k, 'Valor': v, 'Descripción': '' })),
  { 'Parámetro': 'Generado', 'Valor': new Date().toISOString().slice(0, 19).replace('T', ' '), 'Descripción': 'Fecha de generación de esta base' },
  { 'Parámetro': 'Archivos procesados', 'Valor': resultados.length, 'Descripción': 'Documentos leídos desde las carpetas' },
  { 'Parámetro': 'Movimientos', 'Valor': filas.length, 'Descripción': 'Filas de la hoja Base maestra' },
]);

const salida = process.argv[2] || path.join(APP, 'Base maestra Consolidado Legal.xlsx');
XLSX.writeFile(wb, salida);
console.log(JSON.stringify({
  archivo: salida, movimientos: filas.length, lineasDetalle: detalle.length, archivos: resultados.length,
  totalCLP: Math.round(filas.reduce((s, f) => s + f['Monto CLP'], 0)),
  hojas: wb.SheetNames,
}, null, 1));
