#!/bin/bash
# Doble clic para abrir el Consolidado Legal.
# 1) Sincroniza los archivos desde la carpeta de SharePoint (OneDrive) a datos/
# 2) Levanta el servidor local y abre el navegador con el reporte ya generado.
cd "$(dirname "$0")" || exit 1

SP="$HOME/Library/CloudStorage/OneDrive-Global81SPA/Consolidado Cobros - Pagos [Compliance]"
PORT=4173

mkdir -p datos
N_SP=$(find "$SP" -type f \( -iname "*.pdf" -o -iname "*.xlsx" -o -iname "*.xlsm" \) ! -name "~\$*" 2>/dev/null | wc -l | tr -d ' ')
N_LOCAL=$(find datos -type f \( -iname "*.pdf" -o -iname "*.xlsx" -o -iname "*.xlsm" \) 2>/dev/null | wc -l | tr -d ' ')

# Protección: si OneDrive está desactualizado (tiene menos archivos que los ya
# cargados), NO se sincroniza — de lo contrario borraría datos buenos.
if [ -d "$SP" ] && [ "$N_SP" -lt "$N_LOCAL" ]; then
  echo ""
  echo "⚠️  OneDrive está DESACTUALIZADO: tiene $N_SP archivos y el reporte ya carga $N_LOCAL."
  echo "    Se conservan los $N_LOCAL archivos actuales (no se sincroniza para no perderlos)."
  echo "    Revisa que OneDrive haya iniciado sesión y terminado de sincronizar."
  echo ""
elif [ -d "$SP" ]; then
  echo "Sincronizando archivos desde SharePoint/OneDrive…"
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
else
  echo "AVISO: no encuentro la carpeta sincronizada de OneDrive."
  echo "Se usarán los archivos que ya estén en datos/."
fi

# El manifiesto se regenera siempre, con lo que haya quedado en datos/
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

if ! lsof -i :$PORT >/dev/null 2>&1; then
  python3 servidor.py $PORT >/dev/null 2>&1 &
  sleep 1
fi

open "http://localhost:$PORT"
echo ""
echo "Consolidado Legal abierto en http://localhost:$PORT"
echo "Puedes cerrar esta ventana."
