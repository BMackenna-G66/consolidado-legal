#!/bin/bash
# Doble clic en este archivo para abrir el Consolidado Legal.
# Levanta un servidor local en el puerto 4173 y abre el navegador.
cd "$(dirname "$0")" || exit 1

PORT=4173
if ! lsof -i :$PORT >/dev/null 2>&1; then
  python3 -m http.server $PORT >/dev/null 2>&1 &
  sleep 1
fi

open "http://localhost:$PORT"
echo "Consolidado Legal abierto en http://localhost:$PORT"
echo "Puedes cerrar esta ventana; el servidor sigue corriendo hasta que reinicies el Mac."
