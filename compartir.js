// Genera el "Reporte para compartir": un HTML autocontenido con los datos
// incrustados, que cualquiera abre con doble clic — sin login, sin servidor.
// Se sube a la carpeta de SharePoint y los permisos son los de SharePoint.

import { MESES } from './config.js';

// carpetaUrl: vínculo de la carpeta de SharePoint donde se suben las plantillas.
// Va incrustado en el HTML generado (que vive en SharePoint), nunca en el repo.
export function generarHTMLCompartir({ registros, archivos, params, carpetaUrl }) {
  const generado = new Date().toLocaleString('es-CL');
  const datos = JSON.stringify({ registros, archivos, params, generado, MESES, carpetaUrl: carpetaUrl || '' });

  return `<!DOCTYPE html>
<html lang="es"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Consolidado Legal — Global66</title>
<script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
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
details.grupo{background:#fff;border:1px solid var(--borde);border-radius:10px;margin:8px 0}
details.grupo summary{cursor:pointer;padding:9px 14px;font-size:13px;font-weight:600;display:flex;gap:10px;align-items:baseline;flex-wrap:wrap}
details.grupo summary::-webkit-details-marker{display:none}
details.grupo summary::before{content:'▸';color:var(--azul);transition:transform .15s;display:inline-block}
details.grupo[open] summary::before{transform:rotate(90deg)}
details.grupo .g-desc{color:#64748b;font-weight:400;font-size:12px}
details.grupo .g-scroll{overflow-x:auto}
details.grupo td.g-nota{white-space:normal;min-width:260px;color:#64748b}
.est-ok{color:#16a34a}.est-supuesto{color:#b45309}.est-error{color:#dc2626}.est-omitido{color:#94a3b8}
.pasos{background:#fff;border:1px solid var(--borde);border-radius:12px;padding:16px 18px;margin-bottom:10px}
.pasos>b{font-size:13.5px;color:var(--navy)}
.paso{display:flex;gap:11px;align-items:flex-start;margin-top:12px;font-size:12.5px;line-height:1.5}
.paso .num{background:var(--azul);color:#fff;border-radius:50%;width:21px;height:21px;flex-shrink:0;display:flex;align-items:center;justify-content:center;font-size:11.5px;font-weight:700}
.aviso-pend{background:#fffbeb;border:1px solid #fcd34d;color:#92400e;border-radius:10px;padding:10px 14px;font-size:12.5px;margin-bottom:10px}
.badge-pend{background:#fef3c7;color:#92400e;border-radius:9px;padding:1px 8px;font-size:11px;font-weight:600;white-space:nowrap}
.badge-env{background:#dcfce7;color:#166534;border-radius:9px;padding:1px 8px;font-size:11px;font-weight:600;white-space:nowrap}
label.fx{display:flex;flex-direction:column;gap:4px;font-size:11.5px;font-weight:500;color:#475569}
label.fx input,label.fx select{border:1px solid var(--borde);border-radius:6px;padding:7px 9px;font:12.5px Inter;background:#fff}
code{background:#f1f5f9;padding:1px 5px;border-radius:4px;font-size:11px}
#fx-tabla button{padding:3px 9px;font-size:11.5px}
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
  <button id="btn-ficha" style="background:var(--azul);color:#fff;border:0">➕ Agregar gasto/cobro</button>
</div>
<div id="zona-ficha" style="display:none">
  <div class="titulo">Carga manual de gastos y cobros</div>
  <div style="background:#fff;border:1px solid var(--borde);border-radius:12px;padding:16px 18px;margin-bottom:10px">
    <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:10px 14px">
      <label class="fx">Mes<select id="fx-mes"></select></label>
      <label class="fx">Año<input type="number" id="fx-anio" min="2020" max="2100"></label>
      <label class="fx">Solicitante<input id="fx-sol" placeholder="Quién lo pidió"></label>
      <label class="fx">Proveedor / abogado<input id="fx-prov" placeholder="Ej: Aninat"></label>
      <label class="fx">País<select id="fx-pais"><option>Chile</option><option>Colombia</option><option>Perú</option><option>Argentina</option></select></label>
      <label class="fx">Tipo<select id="fx-con"><option>Gastos</option><option>Honorarios</option><option>Juicios y otros</option></select></label>
      <label class="fx">Concepto / partida<input id="fx-cat" placeholder="Ej: Honorarios varios"></label>
      <label class="fx">Moneda origen<select id="fx-mon"><option>CLP</option><option>USD</option><option>PEN</option><option>COP</option><option>ARS</option><option>UF</option></select></label>
      <label class="fx">Monto origen<input type="number" id="fx-monto" step="any" min="0"></label>
      <label class="fx">Monto CLP<input type="number" id="fx-clp" step="1" min="0" placeholder="auto"></label>
      <label class="fx" style="grid-column:1/-1">Detalle<input id="fx-det" placeholder="Referencia, factura, descripción"></label>
    </div>
    <div style="display:flex;gap:8px;align-items:center;margin-top:14px;flex-wrap:wrap">
      <button id="fx-add" style="background:var(--azul);color:#fff;border:0">Agregar al reporte</button>
      <span id="fx-msg" style="font-size:12.5px;color:#16a34a;font-weight:500"></span>
    </div>
  </div>
  <div class="pasos">
    <b>Para que estos movimientos lleguen al consolidado oficial</b>
    <p class="nota">Esta página no puede enviar datos por sí sola: lo que agregas queda <b>solo en este navegador</b>.
    Son dos pasos y quedan registrados para el equipo de Compliance.</p>
    <div class="paso"><span class="num">1</span>
      <div><b>Descarga la planilla</b> con lo que agregaste.
        <div style="margin-top:6px"><button id="fx-exp" style="background:var(--azul);color:#fff;border:0">⬇ Descargar planilla</button></div></div></div>
    <div class="paso"><span class="num">2</span>
      <div><b>Súbela a la carpeta de SharePoint</b> <code>Consolidado Cobros - Pagos [Compliance]</code>
        (dentro de la carpeta del proveedor, o en la raíz). Compliance la incorpora al consolidado en la siguiente actualización.
        <div id="fx-link-zona" style="margin-top:6px"></div></div></div>
  </div>
  <div id="fx-pendiente" class="aviso-pend" style="display:none"></div>
  <div class="scroll" id="fx-tabla"></div>
</div>
<div class="titulo">Resumen de gastos por conceptos legales / administrativos (CLP)</div>
<div id="pivot"></div>
<div style="margin-top:26px"><button id="btn-arch">📄 Archivos leídos</button></div>
<div id="archivos" style="display:none"></div>
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
const FK='consolidado-html-fichas';
const fichasL=()=>{try{return JSON.parse(localStorage.getItem(FK)||'[]')}catch(e){return[]}};
const TASA={CLP:1,USD:D.params.USD_CLP,PEN:D.params.PEN_CLP,COP:D.params.COP_CLP,ARS:D.params.ARS_CLP,UF:D.params.UF_CLP};
const fichaAReg=f=>({anio:+f.anio,mes:+f.mes,dia:1,pais:f.pais,proveedor:f.prov||'Sin proveedor',
  solicitante:f.sol||'',categoria:f.cat||'Carga manual',concepto:f.con,moneda:f.mon,
  montoOrigen:+f.monto||0,clp:Math.round(f.clp!=null&&f.clp!==''?+f.clp:(+f.monto||0)*(TASA[f.mon]||1)),
  usd:0,detalle:f.det||'',archivo:'(carga manual)',carpeta:'(carga manual)',fuente:'manual'});
function todos(){return [...D.registros,...fichasL().map(f=>{const r=fichaAReg(f);r.usd=+(r.clp/(D.params.USD_CLP||1)).toFixed(2);return r})]}
function filtrados(){
  const ver={'Gastos':el('f-g').checked,'Honorarios':el('f-h').checked,'Juicios y otros':el('f-j').checked};
  const fa=el('f-anio').value?+el('f-anio').value:null, fm=el('f-mes').value?+el('f-mes').value:null;
  return todos().filter(r=>ver[r.concepto]&&(!fa||r.anio===fa)&&(!fm||r.mes===fm));
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
const EST={error:['❌ No leído','No se pudo extraer texto (escaneos sin capa de texto) — registrar a mano o pedir el PDF electrónico.'],
 supuesto:['⚠️ Con supuestos','Leídos correctamente; el monto usa una conversión configurable (USD, UF, IGV).'],
 ok:['✅ Procesado','Leídos sin supuestos.'],
 omitido:['⏭ Omitido','Sin monto propio: anexos de horas, respaldos y duplicados que ya aporta otro documento.']};
el('archivos').innerHTML=Object.keys(EST).map(est=>{
  const l=D.archivos.filter(a=>a.estado===est);
  if(!l.length)return '';
  const clp=l.reduce((s,a)=>s+(a.clp||0),0);
  return '<details class="grupo"><summary><span class="est-'+est+'">'+EST[est][0]+'</span><span class="conteo">'+l.length+'</span><span class="g-desc">'+(clp?$(clp)+' aportados · ':'')+EST[est][1]+'</span></summary>'+
    '<div class="g-scroll"><table><thead><tr><th class="etiqueta">Archivo</th><th>Movs.</th><th>CLP</th><th class="etiqueta">Nota</th></tr></thead><tbody>'+
    l.map(a=>'<tr><td class="etiqueta">'+a.ruta+'</td><td class="num">'+(a.n||'')+'</td><td class="num">'+$(a.clp)+'</td><td class="g-nota">'+(a.nota||'')+'</td></tr>').join('')+
    '</tbody></table></div></details>'}).join('');
const nErr=D.archivos.filter(a=>a.estado==='error').length;
const nSup=D.archivos.filter(a=>a.estado==='supuesto').length;
el('btn-arch').textContent='📄 Archivos leídos ('+D.archivos.length+')'+(nErr?' — '+nErr+' sin leer':'')+(nSup?' · '+nSup+' con supuestos':'');
el('btn-arch').addEventListener('click',()=>{const z=el('archivos');z.style.display=z.style.display==='none'?'block':'none'});
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
el('fx-mes').innerHTML=MESES.map((m,i)=>'<option value="'+(i+1)+'">'+m+'</option>').join('');
const hoyF=new Date();el('fx-anio').value=hoyF.getFullYear();el('fx-mes').value=hoyF.getMonth()+1;
el('btn-ficha').addEventListener('click',()=>{const z=el('zona-ficha');z.style.display=z.style.display==='none'?'block':'none';if(z.style.display==='block')z.scrollIntoView({behavior:'smooth'})});
// El paso 2 (subir a SharePoint) es manual: se marca cuándo se descargó cada
// movimiento y se avisa mientras queden pendientes de enviar.
if(D.carpetaUrl)el('fx-link-zona').innerHTML='<a href="'+D.carpetaUrl+'" target="_blank" rel="noopener"><button>📂 Abrir carpeta en SharePoint</button></a>';
else el('fx-link-zona').innerHTML='<span class="nota">Búscala en tu SharePoint de Compliance, o pídele el vínculo a quien te compartió este reporte.</span>';
function renderFichas(){
  const l=fichasL();
  const pend=l.filter(f=>!f.env);
  const av=el('fx-pendiente');
  if(pend.length){av.style.display='block';
    av.innerHTML='⚠ <b>'+pend.length+' movimiento(s) sin enviar.</b> Están guardados solo en este navegador: nadie más los ve todavía. Completa el paso 1 y el paso 2 de arriba.'}
  else av.style.display='none';
  el('fx-tabla').innerHTML=l.length?'<table><thead><tr><th class="etiqueta">Período</th><th class="etiqueta">Solicitante</th><th class="etiqueta">Proveedor</th><th>País</th><th>Tipo</th><th>Origen</th><th>CLP</th><th>Envío</th><th></th></tr></thead><tbody>'+
    l.map(f=>{const r=fichaAReg(f);
      const est=f.env?'<span class="badge-env">✓ Descargado '+f.env+'</span>':'<span class="badge-pend">Sin enviar</span>';
      return '<tr><td class="etiqueta">'+MESES[f.mes-1]+' '+f.anio+'</td><td class="etiqueta">'+(f.sol||'')+'</td><td class="etiqueta">'+(f.prov||'')+'</td><td>'+f.pais+'</td><td>'+f.con+'</td><td class="num">'+f.mon+' '+fmt.format(+f.monto||0)+'</td><td class="num">'+$(r.clp)+'</td><td>'+est+'</td><td><button data-del="'+f.id+'">Borrar</button></td></tr>'}).join('')+'</tbody></table>'
    :'<table><tbody><tr><td class="etiqueta">Aún no has agregado movimientos en este navegador.</td></tr></tbody></table>';
  el('fx-tabla').querySelectorAll('[data-del]').forEach(b=>b.addEventListener('click',()=>{
    localStorage.setItem(FK,JSON.stringify(fichasL().filter(x=>String(x.id)!==b.dataset.del)));renderFichas();render()}));
}
el('fx-add').addEventListener('click',()=>{
  const f={id:Date.now(),mes:+el('fx-mes').value,anio:+el('fx-anio').value,sol:el('fx-sol').value.trim(),
    prov:el('fx-prov').value.trim(),pais:el('fx-pais').value,con:el('fx-con').value,cat:el('fx-cat').value.trim(),
    mon:el('fx-mon').value,monto:el('fx-monto').value,clp:el('fx-clp').value,det:el('fx-det').value.trim()};
  if(!f.monto&&!f.clp){el('fx-msg').textContent='⚠ Ingresa un monto';return}
  if(!f.prov){el('fx-msg').textContent='⚠ Ingresa el proveedor';return}
  localStorage.setItem(FK,JSON.stringify([...fichasL(),f]));
  ['fx-sol','fx-prov','fx-cat','fx-monto','fx-clp','fx-det'].forEach(i=>el(i).value='');
  el('fx-msg').textContent='✓ Agregado al reporte de esta página';setTimeout(()=>el('fx-msg').textContent='',3500);
  renderFichas();render();
});
el('fx-exp').addEventListener('click',()=>{
  const l=fichasL();
  if(!l.length){el('fx-msg').textContent='⚠ No hay movimientos que exportar';return}
  if(typeof XLSX==='undefined'){el('fx-msg').textContent='⚠ Sin conexión: no se pudo cargar el generador de Excel';return}
  const filas=l.map(f=>{const r=fichaAReg(f);return{'Fecha':'01/'+String(f.mes).padStart(2,'0')+'/'+f.anio,
    'Solicitante':f.sol||'','Proveedor':f.prov||'','Pais':f.pais,'Concepto':f.con,'Categoria':f.cat||'Carga manual',
    'Moneda':f.mon,'Monto origen':+f.monto||0,'Monto CLP':r.clp,'Monto USD':+(r.clp/(D.params.USD_CLP||1)).toFixed(2),'Detalle':f.det||''}});
  const ws=XLSX.utils.json_to_sheet(filas);const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,'Carga manual');
  const quien=(l[0].sol||'equipo').replace(/[^\\wÁÉÍÓÚáéíóúñÑ]+/g,'-').replace(/^-|-$/g,'');
  XLSX.writeFile(wb,'Carga manual - '+quien+' - '+l[0].anio+'-'+String(l[0].mes).padStart(2,'0')+'.xlsx');
  const hoy=new Date().toLocaleDateString('es-CL');
  localStorage.setItem(FK,JSON.stringify(fichasL().map(f=>({...f,env:f.env||hoy}))));
  renderFichas();
  el('fx-msg').textContent='';
  el('fx-pendiente').style.display='block';
  el('fx-pendiente').innerHTML='📥 <b>Planilla descargada.</b> Falta el paso 2: súbela a la carpeta de SharePoint para que llegue a Compliance.'
    +(D.carpetaUrl?' <a href="'+D.carpetaUrl+'" target="_blank" rel="noopener">Abrir carpeta</a>':'');
});
renderFichas();
render();
</script></body></html>`;
}
