"""Ponte local pra impressao vinda do navegador (Alakazam).

O Chromium do Raspberry nao consegue imprimir na fila raw da Zebra (sem
driver, ver sdd/STATE.md do Alakazam "Etiquetas via Raspberry"): o dialogo
nativo de impressao exige driver, que nao existe pra ZPL no Raspbian.

Esse modulo sobe um servidor HTTP em localhost, chamado pela pagina do
Alakazam via fetch: ela pede o ZPL pronto ao backend (formato=zpl) e manda o
corpo cru pra ca, que executa `lp -o raw` — mesmo comando que send_to_printer()
ja usa em pedidos_ifood_gui.py pro fluxo EPL. Roda dentro do Rotom (nao como
processo avulso) pra ganhar o auto-update que o Rotom ja tem — ver D-187 no
repo do Alakazam.
"""
import subprocess
import threading
from http.server import BaseHTTPRequestHandler, HTTPServer

PORT = 9876
ALLOWED_ORIGINS = {
    "https://darkstores-264e6.web.app",
    "https://darkstores-264e6.firebaseapp.com",
}


class _Handler(BaseHTTPRequestHandler):
    def _cors(self):
        origin = self.headers.get("Origin", "")
        if origin in ALLOWED_ORIGINS:
            self.send_header("Access-Control-Allow-Origin", origin)
        self.send_header("Access-Control-Allow-Methods", "POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")

    def do_OPTIONS(self):
        self.send_response(204)
        self._cors()
        self.end_headers()

    def do_GET(self):
        if self.path == "/status":
            self.send_response(200)
            self._cors()
            self.send_header("Content-Type", "application/json")
            self.end_headers()
            self.wfile.write(b'{"ok":true}')
            return
        self.send_response(404)
        self._cors()
        self.end_headers()

    def do_POST(self):
        if self.path != "/print":
            self.send_response(404)
            self._cors()
            self.end_headers()
            return
        length = int(self.headers.get("Content-Length", 0))
        body = self.rfile.read(length)
        try:
            # Mesmo comando de send_to_printer() (sem -d: usa a impressora
            # padrao do CUPS) — nao duplica logica, so troca EPL por ZPL cru.
            proc = subprocess.run(
                ["lp", "-o", "raw"], input=body,
                stdout=subprocess.PIPE, stderr=subprocess.PIPE,
            )
            if proc.returncode != 0:
                raise RuntimeError(proc.stderr.decode("utf-8", "ignore"))
            self.send_response(200)
            self._cors()
            self.send_header("Content-Type", "application/json")
            self.end_headers()
            self.wfile.write(b'{"ok":true}')
        except Exception as exc:
            self.send_response(500)
            self._cors()
            self.send_header("Content-Type", "application/json")
            self.end_headers()
            self.wfile.write(('{"ok":false,"error":%r}' % str(exc)).encode())

    def log_message(self, fmt, *args):
        pass


def start():
    """Sobe o servidor numa thread daemon. Chamado uma vez, na subida do app."""
    server = HTTPServer(("0.0.0.0", PORT), _Handler)
    threading.Thread(target=server.serve_forever, daemon=True).start()
    return server
