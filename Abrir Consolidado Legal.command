#!/bin/bash
# Abre la app en el navegador. Ya no copia ni sincroniza carpetas: la app lee
# los documentos directamente desde SharePoint.
# Para el uso normal basta el sitio publicado:
#   https://bmackenna-g66.github.io/consolidado-legal/
cd "$(dirname "$0")" || exit 1
PORT=4173
if ! lsof -i :$PORT >/dev/null 2>&1; then
  python3 servidor.py $PORT >/dev/null 2>&1 &
  sleep 1
fi
open "http://localhost:$PORT"
echo "Consolidado Legal abierto en http://localhost:$PORT"
echo "La app lee desde SharePoint; no hay carpeta local que mantener."
