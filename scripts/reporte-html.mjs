// Genera un reporte HTML autocontenido: un solo archivo con los datos incrustados,
// que se abre con doble clic y no necesita servidor, login ni conexión.
// Pensado para subirlo a SharePoint: quien tenga acceso a la carpeta lo abre y ve
// el consolidado; los permisos son los de SharePoint, no hay nada publicado afuera.
import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath, pathToFileURL } from 'node:url';

const require = createRequire(import.meta.url);
const XLSX = require('xlsx');
const pdfjs = require('pdfjs-dist/legacy/build/pdf.js');
pdfjs.GlobalWorkerOptions.workerSrc = require.resolve('pdfjs-dist/legacy/build/pdf.worker.js');

// fileURLToPath aguanta espacios en la ruta, .pathname los dejaría percent-encoded
const APP = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const mod = f => pathToFileURL(path.join(APP, f)).href;
globalThis.window = { XLSX, pdfjsLib: pdfjs };
globalThis.localStorage = { getItem: () => null, setItem() {}, removeItem() {} };

const { procesarArchivo, deduplicar, aplicarBaseHistorica, enriquecerSolicitantes } = await import(mod('parsers.js'));
const { DEFAULT_PARAMS, MESES } = await import(mod('config.js'));

const DATOS = process.env.DATOS || path.join(APP, 'datos');
const manifest = JSON.parse(fs.readFileSync(path.join(DATOS, 'manifest.json'), 'utf8'));
const params = { ...DEFAULT_PARAMS };

const resultados = [];
for (const m of manifest) {
  const buf = fs.readFileSync(path.join(DATOS, m.archivo));
  const arrayBuffer = buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
  resultados.push(await procesarArchivo({ nombre: m.archivo.split('/').pop(), ruta: m.ruta, arrayBuffer }, params));
}
deduplicar(resultados);
enriquecerSolicitantes(resultados);
aplicarBaseHistorica(resultados);

// Se incrustan solo los movimientos y el inventario: lo mínimo para reconstruir
// el reporte en el navegador de quien lo abra.
const registros = [];
for (const res of resultados) {
  const partes = String(res.ruta || '').split('/').filter(Boolean);
  for (const r of (res.registros || [])) {
    registros.push({
      anio: r.anio, mes: r.mes, dia: r.dia, pais: r.pais, proveedor: r.proveedor,
      solicitante: /^Varios \(/.test(r.solicitante || '') ? '' : (r.solicitante || ''),
      categoria: r.categoria, concepto: r.concepto, moneda: r.moneda,
      montoOrigen: r.montoOrigen, clp: Math.round(r.clp),
      usd: r.usd != null ? r.usd : +(r.clp / params.USD_CLP).toFixed(2),
      detalle: r.detalle || '', archivo: r.archivo || '',
      carpeta: partes[0] || '(raíz)', mesCarpeta: partes.length > 2 ? partes[partes.length - 2] : '',
      fuente: r.fuente || 'documento',
    });
  }
}
const archivos = resultados.map(res => ({
  estado: res.estado, archivo: res.archivo, ruta: res.ruta || '',
  n: (res.registros || []).length,
  clp: Math.round((res.registros || []).reduce((s, r) => s + r.clp, 0)),
  nota: res.nota || '',
}));

const { generarHTMLCompartir } = await import(mod('compartir.js'));
const html = generarHTMLCompartir({ registros, archivos, params });

const salida = process.argv[2] || path.join(APP, 'Reporte Consolidado Legal.html');
fs.writeFileSync(salida, html);
console.log(JSON.stringify({ archivo: salida, movimientos: registros.length, archivos: archivos.length, kb: Math.round(html.length / 1024) }, null, 1));
