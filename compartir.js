// Genera el "Reporte para compartir": un HTML autocontenido con los datos
// incrustados, que cualquiera abre con doble clic — sin login, sin servidor.
// Se sube a la carpeta de SharePoint y los permisos son los de SharePoint.

import { MESES } from './config.js';

export function generarHTMLCompartir({ registros, archivos, params }) {
  const generado = new Date().toLocaleString('es-CL');
  const datos = JSON.stringify({ registros, archivos, params, generado, MESES });

  return `<!DOCTYPE html>
<html lang="es"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Consolidado Legal — Global66</title>
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
:root{--azul:#3458b1;--navy:#0a1b44;--borde:#e2e8f0}
*{box-sizing:border-box}body{margin:0;font-family:Inter,system-ui,sans-serif;background:#f8fafc;color:#1e293b}
header{background:var(--navy);color:#fff;padding:14px 28px;display:flex;gap:14px;align-items:baseline;flex-wrap:wrap}
header h1{font-size:17px;margin:0;font-weight:600}header .sub{font-size:12px;opacity:.75}
main{padding:20px 28px 60px;max-width:1500px;margin:0 auto}
#tarjetas{display:flex;gap:14px;flex-wrap:wrap;margin-bottom:22px}
.tarjeta{background:#fff;border:1px solid var(--borde);border-radius:12px;padding:14px 20px;min-width:180px}
.tarjeta-titulo{font-size:11.5px;text-transform:uppercase;letter-spacing:.4px;color:#64748b}
.tarjeta-valor{font-size:21px;font-weight:700;color:var(--navy);margin-top:3px}
.tarjeta-sub{font-size:11.5px;color:#94a3b8}
#filtros{display:flex;gap:16px;align-items:center;flex-wrap:wrap;font-size:13px;margin-bottom:14px}
#filtros select{border:1px solid var(--borde);border-radius:6px;padding:5px 8px;font:13px Inter;background:#fff}
#filtros label{display:flex;gap:5px;align-items:center;cursor:pointer}
.titulo{font-size:15px;font-weight:600;color:var(--navy);margin:26px 0 8px}
.scroll{overflow-x:auto;background:#fff;border:1px solid var(--borde);border-radius:12px}
table{border-collapse:collapse;width:100%;font-size:12px;white-space:nowrap}
th,td{padding:6px 10px;border-bottom:1px solid #f1f5f9}
thead th{background:var(--navy);color:#fff;font-weight:500;text-align:right;position:sticky;top:0}
thead th.etiqueta,td.etiqueta{text-align:left;min-width:230px}
thead th.anio{text-align:center;border-left:2px solid #33406b}
td.num{text-align:right;font-variant-numeric:tabular-nums}
.total-anio{background:#eef2ff;font-weight:600;border-left:2px solid var(--borde)}
thead .total-anio,thead .total-general{background:#22315e}
.total-general{background:#e8edfb;font-weight:700}
tr.lvl-0{font-weight:600;background:#fbfcff}tr.lvl-0 .etiqueta{padding-left:10px}
tr.lvl-1 .etiqueta{padding-left:30px;color:#334155}
tr.lvl-2 .etiqueta{padding-left:50px;color:#475569}tr.lvl-2{font-size:11.5px}
tr.lvl-3 .etiqueta{padding-left:70px;color:#64748b;font-style:italic}tr.lvl-3{font-size:11px}
tr.expandible{cursor:pointer}tr.expandible:hover{background:#f1f5f9}
tr.oculto{display:none}.flecha{display:inline-block;transition:transform .15s;color:var(--azul)}
tr.abierto .flecha{transform:rotate(90deg)}
tr.fila-total{font-weight:700;background:#eef2ff;border-top:2px solid var(--azul)}
.conteo{display:inline-block;margin-left:8px;font-size:10px;color:#94a3b8;background:#f1f5f9;border-radius:9px;padding:1px 7px}
.nota{font-size:12px;color:#64748b;margin:6px 0 0}
button{background:#fff;color:var(--navy);border:1px solid var(--borde);border-radius:8px;padding:7px 14px;font:500 13px Inter;cursor:pointer}
</style></head><body>
<header><h1>Consolidado Legal por País</h1><span class="sub">Compliance · Global66 — reporte generado el ${generado}</span></header>
<main>
<div id="tarjetas"></div>
<div id="filtros">
  <b>Concepto:</b>
  <label><input type="checkbox" id="f-g" checked> Gastos</label>
  <label><input type="checkbox" id="f-h" checked> Honorarios</label>
  <label><input type="checkbox" id="f-j"> Juicios y otros</label>
  <b>Año:</b><select id="f-anio"><option value="">Todos</option></select>
  <b>Mes:</b><select id="f-mes"><option value="">Todos</option></select>
  <b>Abrir por:</b><select id="f-agr"></select>
  <button id="btn-exp">⌄ Expandir todo</button>
  <button id="btn-csv">⬇ Exportar CSV</button>
</div>
<div class="titulo">Resumen de gastos por conceptos legales / administrativos (CLP)</div>
<div id="pivot"></div>
<div class="titulo">Archivos leídos</div>
<div class="scroll" id="archivos"></div>
<p class="nota">Reporte estático: refleja los documentos existentes al momento de generarlo. Para una versión al día, regenerarlo desde la app y reemplazar este archivo.</p>
</main>
<script>
const D = ${datos};
const MESES = D.MESES, fmt = new Intl.NumberFormat('es-CL');
const $ = v => (v==null||v===0)?'':'$'+fmt.format(Math.round(v));
const AGR = [
 {id:'pais',n:'País → Proveedor → Concepto → Solicitante',d:['pais','proveedor','categoria','solicitante']},
 {id:'sol',n:'País → Solicitante → Proveedor → Concepto',d:['pais','solicitante','proveedor','categoria']},
 {id:'prov',n:'Proveedor → Concepto → Solicitante',d:['proveedor','categoria','solicitante']},
 {id:'con',n:'Concepto → País → Proveedor → Solicitante',d:['categoria','pais','proveedor','solicitante']}];
const VACIO={solicitante:'Sin solicitante registrado',categoria:'Sin concepto',proveedor:'Sin proveedor',pais:'Sin país'};
const g=(r,d)=>{const v=r[d];return (v==null||String(v).trim()==='')?(VACIO[d]||'(sin dato)'):String(v)};
const el=i=>document.getElementById(i);
el('f-agr').innerHTML=AGR.map(a=>'<option value="'+a.id+'">'+a.n+'</option>').join('');
el('f-mes').innerHTML='<option value="">Todos</option>'+MESES.map((m,i)=>'<option value="'+(i+1)+'">'+m+'</option>').join('');
el('f-anio').innerHTML='<option value="">Todos</option>'+[...new Set(D.registros.map(r=>r.anio).filter(Boolean))].sort().map(a=>'<option>'+a+'</option>').join('');
function filtrados(){
  const ver={'Gastos':el('f-g').checked,'Honorarios':el('f-h').checked,'Juicios y otros':el('f-j').checked};
  const fa=el('f-anio').value?+el('f-anio').value:null, fm=el('f-mes').value?+el('f-mes').value:null;
  return D.registros.filter(r=>ver[r.concepto]&&(!fa||r.anio===fa)&&(!fm||r.mes===fm));
}
function nodo(){return{m:{},a:{},t:0,n:0,h:new Map()}}
function acum(x,r){const k=r.anio+'-'+r.mes;x.m[k]=(x.m[k]||0)+r.clp;x.a[r.anio]=(x.a[r.anio]||0)+r.clp;x.t+=r.clp;x.n++}
function render(){
  const regs=filtrados();
  const anios=[...new Set(regs.map(r=>r.anio).filter(Boolean))].sort();
  const cols=[];
  for(const a of anios){const ms=[...new Set(regs.filter(r=>r.anio===a).map(r=>r.mes).filter(Boolean))];
    if(!ms.length)continue; for(let m=Math.min(...ms);m<=Math.max(...ms);m++)cols.push({a,m}); cols.push({a,tot:1});}
  const dims=(AGR.find(x=>x.id===el('f-agr').value)||AGR[0]).d;
  const raiz=nodo();
  for(const r of regs){if(!r.anio||!r.mes)continue;acum(raiz,r);let c=raiz;
    for(const d of dims){const k=g(r,d);if(!c.h.has(k))c.h.set(k,nodo());c=c.h.get(k);acum(c,r)}}
  const celdas=x=>cols.map(c=>c.tot?'<td class="total-anio num">'+$(x.a[c.a])+'</td>':'<td class="num">'+$(x.m[c.a+'-'+c.m])+'</td>').join('')+'<td class="total-general num">'+$(x.t)+'</td>';
  let h='<div class="scroll"><table><thead><tr><th class="etiqueta"></th>';
  for(const a of anios){const n=cols.filter(c=>c.a===a&&!c.tot).length;h+='<th colspan="'+n+'" class="anio">'+a+'</th><th class="total-anio">Total '+a+'</th>'}
  h+='<th class="total-general">Total general</th></tr><tr><th class="etiqueta">Etiquetas de fila</th>';
  for(const c of cols)h+=c.tot?'<th class="total-anio"></th>':'<th>'+MESES[c.m-1]+'</th>';
  h+='<th class="total-general"></th></tr></thead><tbody>';
  const filas=[];
  (function rec(x,niv,ruta){[...x.h.entries()].sort((a,b)=>b[1].t-a[1].t).forEach(([k,hij],i)=>{
    const rr=ruta?ruta+'.'+i:''+i, tiene=hij.h.size>0;
    filas.push('<tr class="lvl-'+niv+(tiene?' expandible':'')+(niv>0?' oculto':'')+'" data-ruta="'+rr+'" data-nivel="'+niv+'"><td class="etiqueta">'+(tiene?'<span class="flecha">▸</span> ':'')+k+'<span class="conteo">'+hij.n+'</span></td>'+celdas(hij)+'</tr>');
    if(tiene)rec(hij,niv+1,rr)})})(raiz,0,'');
  h+=filas.join('')+'<tr class="fila-total"><td class="etiqueta">Total general</td>'+celdas(raiz)+'</tr></tbody></table></div>';
  el('pivot').innerHTML=h;
  el('pivot').querySelectorAll('tr.expandible').forEach(tr=>tr.addEventListener('click',()=>{
    const r=tr.dataset.ruta, ab=tr.classList.toggle('abierto');
    el('pivot').querySelectorAll('tr[data-ruta^="'+r+'."]').forEach(x=>{
      const p=x.dataset.ruta.split('.').length-r.split('.').length;
      if(ab&&p===1)x.classList.remove('oculto');else if(!ab){x.classList.add('oculto');x.classList.remove('abierto')}})}));
  const tot=regs.reduce((s,r)=>s+r.clp,0),pa={},pp={};
  for(const r of regs){if(r.anio)pa[r.anio]=(pa[r.anio]||0)+r.clp;pp[r.pais]=(pp[r.pais]||0)+r.clp}
  const t=(a,b,c)=>'<div class="tarjeta"><div class="tarjeta-titulo">'+a+'</div><div class="tarjeta-valor">'+($(b)||'$0')+'</div><div class="tarjeta-sub">'+(c||'')+'</div></div>';
  let ht=t('Total general',tot,regs.length+' movimientos');
  for(const a of Object.keys(pa).sort())ht+=t('Total '+a,pa[a]);
  const li=Object.entries(pp).sort((x,y)=>y[1]-x[1])[0];
  if(li)ht+=t('País con mayor gasto',li[1],li[0]);
  el('tarjetas').innerHTML=ht;
}
el('archivos').innerHTML='<table><thead><tr><th class="etiqueta">Archivo</th><th>Estado</th><th>Movs.</th><th>CLP</th><th class="etiqueta">Nota</th></tr></thead><tbody>'+
 D.archivos.map(a=>'<tr><td class="etiqueta">'+a.ruta+'</td><td>'+a.estado+'</td><td class="num">'+a.n+'</td><td class="num">'+$(a.clp)+'</td><td class="etiqueta">'+a.nota+'</td></tr>').join('')+'</tbody></table>';
['f-g','f-h','f-j','f-anio','f-mes','f-agr'].forEach(i=>el(i).addEventListener('change',render));
el('btn-exp').addEventListener('click',()=>{
  const fs=[...el('pivot').querySelectorAll('tr[data-ruta]')], cer=fs.some(f=>f.classList.contains('oculto'));
  fs.forEach(f=>{f.classList.toggle('oculto',!cer&&f.dataset.nivel!=='0');f.classList.toggle('abierto',cer&&f.classList.contains('expandible'))});
  el('btn-exp').textContent=cer?'⌃ Colapsar todo':'⌄ Expandir todo'});
el('btn-csv').addEventListener('click',()=>{
  const enc=['Año','Mes','País','Proveedor','Solicitante','Concepto cobrado','Tipo','Moneda','Monto origen','CLP','USD','Detalle','Carpeta','Archivo'];
  const f=filtrados().map(r=>[r.anio,r.mes,r.pais,r.proveedor,r.solicitante,r.categoria,r.concepto,r.moneda,r.montoOrigen,r.clp,r.usd,r.detalle,r.carpeta,r.archivo]);
  const csv=[enc,...f].map(x=>x.map(c=>'"'+String(c==null?'':c).replace(/"/g,'""')+'"').join(';')).join('\\n');
  const a=document.createElement('a');a.href=URL.createObjectURL(new Blob(['\\ufeff'+csv],{type:'text/csv;charset=utf-8'}));
  a.download='consolidado_legal.csv';a.click()});
render();
</script></body></html>`;
}
