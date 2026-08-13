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

// ---------- Carpeta local con memoria (File System Access API) ----------
// Chrome/Edge permiten guardar el "handle" de la carpeta elegida: la primera vez
// se selecciona, y en las visitas siguientes basta un clic en "Actualizar".

const DB = 'consolidado-legal';

function idb() {
  return new Promise((res, rej) => {
    const req = indexedDB.open(DB, 1);
    req.onupgradeneeded = () => req.result.createObjectStore('handles');
    req.onsuccess = () => res(req.result);
    req.onerror = () => rej(req.error);
  });
}

async function guardarHandle(handle) {
  const db = await idb();
  await new Promise((res, rej) => {
    const tx = db.transaction('handles', 'readwrite');
    tx.objectStore('handles').put(handle, 'carpeta');
    tx.oncomplete = res; tx.onerror = () => rej(tx.error);
  });
}

export async function carpetaGuardada() {
  try {
    const db = await idb();
    return await new Promise((res) => {
      const tx = db.transaction('handles', 'readonly');
      const rq = tx.objectStore('handles').get('carpeta');
      rq.onsuccess = () => res(rq.result || null);
      rq.onerror = () => res(null);
    });
  } catch { return null; }
}

export function soportaFSA() { return typeof window.showDirectoryPicker === 'function'; }

async function leerDirectorio(dir, ruta, archivos, onProgreso) {
  for await (const [nombre, handle] of dir.entries()) {
    if (handle.kind === 'directory') {
      await leerDirectorio(handle, `${ruta}/${nombre}`, archivos, onProgreso);
    } else if (EXTENSIONES.test(nombre) && !nombre.startsWith('~$')) {
      onProgreso?.(`Leyendo ${nombre}…`);
      const f = await handle.getFile();
      archivos.push({ nombre, ruta: `${ruta}/${nombre}`, arrayBuffer: await f.arrayBuffer() });
    }
  }
}

export async function elegirCarpetaFSA(onProgreso) {
  const dir = await window.showDirectoryPicker({ id: 'consolidado-legal', mode: 'read' });
  await guardarHandle(dir).catch(() => {});
  const archivos = [];
  await leerDirectorio(dir, '', archivos, onProgreso);
  return { archivos, nombre: dir.name };
}

export async function leerCarpetaRecordada(onProgreso) {
  const dir = await carpetaGuardada();
  if (!dir) throw new Error('No hay carpeta guardada');
  if ((await dir.queryPermission({ mode: 'read' })) !== 'granted') {
    if ((await dir.requestPermission({ mode: 'read' })) !== 'granted') throw new Error('Permiso de lectura denegado');
  }
  const archivos = [];
  await leerDirectorio(dir, '', archivos, onProgreso);
  return { archivos, nombre: dir.name };
}

// ---------- SharePoint vía Microsoft Graph ----------

let msalApp = null;

const CLIENT_ID_KEY = 'consolidado-azure-client-id';

// El Client ID del registro de Azure AD puede venir del código o guardarse en el
// navegador, para habilitar el modo en línea sin editar archivos.
export function clientId() {
  return AZURE.clientId || localStorage.getItem(CLIENT_ID_KEY) || '';
}

export function setClientId(id) {
  const limpio = String(id || '').trim();
  if (limpio) localStorage.setItem(CLIENT_ID_KEY, limpio);
  else localStorage.removeItem(CLIENT_ID_KEY);
  msalApp = null; // fuerza recrear el cliente con el nuevo id
}

export function graphDisponible() {
  return Boolean(clientId());
}

async function tokenGraph() {
  if (!msalApp) {
    msalApp = new window.msal.PublicClientApplication({
      auth: {
        clientId: clientId(),
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
