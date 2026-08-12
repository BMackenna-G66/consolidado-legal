#!/bin/bash
# Doble clic para abrir el Consolidado Legal.
# 1) Sincroniza los archivos desde la carpeta de SharePoint (OneDrive) a datos/
# 2) Levanta el servidor local y abre el navegador con el reporte ya generado.
cd "$(dirname "$0")" || exit 1

SP="$HOME/Library/CloudStorage/OneDrive-Global81SPA/Consolidado Cobros - Pagos [Compliance]"
PORT=4173

if [ -d "$SP" ]; then
  echo "Sincronizando archivos desde SharePoint/OneDrive…"
  mkdir -p datos
  CAMBIOS=$(rsync -ai --delete \
    --include="*/" --include="*.pdf" --include="*.xlsx" --include="*.xlsm" \
    --exclude="*" --prune-empty-dirs "$SP/" datos/ | grep -E '^(>f|\*deleting)' )
  find datos -name "~\$*" -delete 2>/dev/null

  if [ -z "$CAMBIOS" ]; then
    echo ""
    echo "⚠️  SIN CAMBIOS: no llegó ningún archivo nuevo ni modificado."
    echo "    Si esperabas cargas nuevas, revisa en OneDrive (icono de la nube en la"
    echo "    barra de menú) que la sincronización esté al día y no en pausa."
  else
    echo ""
    echo "Cambios detectados:"
    echo "$CAMBIOS" | sed -e 's/^>f[^ ]* /  + /' -e 's/^\*deleting  */  - /'
  fi
  echo ""

  python3 - <<'PY'
import json, os
items = []
for root, _, files in os.walk('datos'):
    for f in sorted(files):
        if f.startswith('~$') or not f.lower().endswith(('.pdf', '.xlsx', '.xlsm')):
            continue
        rel = os.path.relpath(os.path.join(root, f), 'datos').replace(os.sep, '/')
        items.append({'archivo': rel, 'ruta': '/' + rel})
items.sort(key=lambda x: (0 if 'Consolidado Paises' in x['archivo'] else 1, x['archivo']))
json.dump(items, open('datos/manifest.json', 'w'), ensure_ascii=False, indent=1)
print(f'{len(items)} archivos listos.')
PY
else
  echo "AVISO: no encuentro la carpeta sincronizada de OneDrive."
  echo "Se usarán los archivos que ya estén en datos/."
fi

if ! lsof -i :$PORT >/dev/null 2>&1; then
  python3 servidor.py $PORT >/dev/null 2>&1 &
  sleep 1
fi

open "http://localhost:$PORT"
echo ""
echo "Consolidado Legal abierto en http://localhost:$PORT"
echo "Puedes cerrar esta ventana."
