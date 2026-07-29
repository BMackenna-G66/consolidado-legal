// Fuentes de datos: SharePoint (Microsoft Graph con login corporativo) o carpeta local
// (la misma carpeta sincronizada con OneDrive, o una copia descargada).
// Ambas devuelven la misma forma: [{ nombre, ruta, arrayBuffer }]

import { AZURE } from './config.js';

const EXTENSIONES = /\.(pdf|xlsx|xlsm)$/i;

// ---------- Carpeta local (input webkitdirectory) ----------

export async function leerCarpetaLocal(fileList, onProgreso) {
  const archivos = [];
  const lista = [...fileList].filter(f => EXTENSIONES.test(f.name) && !f.name.startsWith('~$'));
  let i = 0;
  for (const f of lista) {
    onProgreso?.(`Leyendo ${++i}/${lista.length}: ${f.name}`);
    const rel = '/' + (f.webkitRelativePath || f.name);
    archivos.push({ nombre: f.name, ruta: rel, arrayBuffer: await f.arrayBuffer() });
  }
  return archivos;
}

// ---------- SharePoint vía Microsoft Graph ----------

let msalApp = null;

export function graphDisponible() {
  return Boolean(AZURE.clientId);
}

async function tokenGraph() {
  if (!msalApp) {
    msalApp = new window.msal.PublicClientApplication({
      auth: {
        clientId: AZURE.clientId,
        authority: `https://login.microsoftonline.com/${AZURE.tenantId}`,
        redirectUri: window.location.origin + window.location.pathname,
      },
      cache: { cacheLocation: 'sessionStorage' },
    });
    await msalApp.initialize();
  }
  const scopes = ['Files.Read.All'];
  const cuentas = msalApp.getAllAccounts();
  if (cuentas.length) {
    try {
      const r = await msalApp.acquireTokenSilent({ scopes, account: cuentas[0] });
      return r.accessToken;
    } catch { /* cae al popup */ }
  }
  const r = await msalApp.loginPopup({ scopes });
  return r.accessToken;
}

function shareId(url) {
  const b64 = btoa(url).replace(/=+$/, '').replace(/\//g, '_').replace(/\+/g, '-');
  return 'u!' + b64;
}

async function gGET(token, url) {
  const r = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (!r.ok) throw new Error(`Graph ${r.status} en ${url.split('?')[0]}`);
  return r.json();
}

function urlCarpeta() {
  let url = AZURE.shareUrl || localStorage.getItem('consolidado-share-url') || '';
  if (!url) {
    url = (prompt('Pega el enlace compartido de la carpeta "Consolidado Cobros - Pagos [Compliance]" en SharePoint:') || '').trim();
    if (url) localStorage.setItem('consolidado-share-url', url);
  }
  if (!url) throw new Error('Falta el enlace de la carpeta de SharePoint');
  return url;
}

export async function leerSharePoint(onProgreso) {
  const carpeta = urlCarpeta();
  const token = await tokenGraph();
  onProgreso?.('Resolviendo carpeta compartida…');
  const raiz = await gGET(token, `https://graph.microsoft.com/v1.0/shares/${shareId(carpeta)}/driveItem`);
  const driveId = raiz.parentReference.driveId;
  const archivos = [];

  async function recorrer(itemId, rutaBase) {
    let url = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/children?$top=200&$select=id,name,folder,file,size,@microsoft.graph.downloadUrl`;
    while (url) {
      const pagina = await gGET(token, url);
      for (const it of pagina.value) {
        if (it.folder) {
          await recorrer(it.id, `${rutaBase}/${it.name}`);
        } else if (EXTENSIONES.test(it.name)) {
          onProgreso?.(`Descargando ${it.name}…`);
          const resp = await fetch(it['@microsoft.graph.downloadUrl']);
          archivos.push({ nombre: it.name, ruta: `${rutaBase}/${it.name}`, arrayBuffer: await resp.arrayBuffer() });
        }
      }
      url = pagina['@odata.nextLink'] || null;
    }
  }

  await recorrer(raiz.id, '');
  return archivos;
}
