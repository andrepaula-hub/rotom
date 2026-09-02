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
import re
import subprocess
import threading
from http.server import BaseHTTPRequestHandler, HTTPServer

PORT = 9876
ALLOWED_ORIGINS = {
    "https://darkstores-264e6.web.app",
    "https://darkstores-264e6.firebaseapp.com",
}
# Canais de preview do Hosting (`firebase hosting:channel:deploy`) ganham um
# subdominio proprio por canal, com um hash que muda se o canal expirar e for
# recriado (ex: darkstores-264e6--staging-hpjmn7nk.web.app). Prototipo do
# Separador so existe nesses canais (nunca no link oficial, ver
# docs/modulos/separador.md do Alakazam) — sem isso na allowlist, o
# navegador bloqueia a resposta daqui e a impressao nao sai, mesmo com
# /api/etiquetas/gerar respondendo 200 (achado real 2026-09-02, Raspberry de
# Setor Bueno: back confirmou 200, quem barrou foi o CORS aqui).
ORIGEM_PREVIEW_RE = re.compile(r"^https://darkstores-264e6--[a-z0-9-]+\.web\.app$")


def _origem_permitida(origin: str) -> bool:
    return origin in ALLOWED_ORIGINS or bool(ORIGEM_PREVIEW_RE.match(origin))


class _Handler(BaseHTTPRequestHandler):
    def _cors(self):
        origin = self.headers.get("Origin", "")
        if _origem_permitida(origin):
            self.send_header("Access-Control-Allow-Origin", origin)
            # Private Network Access (Chrome): pagina publica (https) pedindo
            # recurso de rede privada (localhost) exige esse header no
            # preflight, sem CORS normal nao bastar. Sem ele o Chrome recusa
            # a chamada pra localhost:9876 de qualquer origem, oficial ou
            # preview — nao so nao devolve resposta, nem chega a mandar o
            # POST de verdade.
            if self.headers.get("Access-Control-Request-Private-Network") == "true":
                self.send_header("Access-Control-Allow-Private-Network", "true")
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
