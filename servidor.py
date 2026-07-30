#!/usr/bin/env python3
"""Servidor local del Consolidado Legal. Igual que http.server pero sin caché,
para que el reporte siempre refleje los archivos y el código actuales."""
import sys
from functools import partial
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer


class SinCache(SimpleHTTPRequestHandler):
    def end_headers(self):
        self.send_header('Cache-Control', 'no-store, must-revalidate')
        self.send_header('Pragma', 'no-cache')
        super().end_headers()

    def log_message(self, *a):
        pass


puerto = int(sys.argv[1]) if len(sys.argv) > 1 else 4173
ThreadingHTTPServer(('127.0.0.1', puerto), partial(SinCache, directory='.')).serve_forever()
