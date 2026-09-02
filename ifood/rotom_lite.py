#!/usr/bin/env python3
"""Rotom enxuto: so o necessario pro Raspberry imprimir via navegador.

Sem GUI, sem login de Google Sheets, sem iFood — isso tudo mora em
pedidos_ifood_gui.py, que este processo nao roda. Duas coisas so:

1. Sobe o print_bridge (servidor local que o Alakazam chama pra imprimir).
2. Se auto-atualiza, consultando o MESMO manifest que o Rotom completo usa
   (D-187, repo shopper-darkstores) — reaproveita o mecanismo de distribuicao
   em massa em vez de reinventar um pra este script sozinho.

Rodar direto (`python3 rotom_lite.py`) ou, no Raspberry, via systemd — ver
scripts/rotom-lite.service.
"""
import hashlib
import json
import os
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
import zipfile

import print_bridge

MANIFEST_URL = "https://alakazam-backend-723462849448.us-east4.run.app/internal/rotom-manifest"
CHECK_INTERVAL_SECONDS = 300
VERSION_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "rotom_version.json")
DEVICE_NAME = os.environ.get("ROTOM_DEVICE_NAME", "")

# Arquivos que este processo mantem atualizados — o ZIP publicado carrega
# tambem pedidos_ifood_gui.py (o app completo usa o mesmo ZIP), mas o lite
# so troca os proprios.
ARQUIVOS_GERENCIADOS = ["print_bridge.py", "rotom_lite.py"]


def _log(msg):
    print(f"[rotom_lite] {msg}", flush=True)


def _versao_atual():
    try:
        with open(VERSION_FILE, "r", encoding="utf-8") as fh:
            return str(json.load(fh).get("version") or "").strip()
    except Exception:
        return ""


def _consultar_manifest(versao_atual):
    query = {"deviceName": DEVICE_NAME, "currentVersion": versao_atual}
    url = MANIFEST_URL + "?" + urllib.parse.urlencode(query)
    with urllib.request.urlopen(url, timeout=20) as resp:
        return json.loads(resp.read().decode("utf-8"))


def _baixar(url, destino):
    with urllib.request.urlopen(url, timeout=60) as resp, open(destino, "wb") as fh:
        fh.write(resp.read())


def _sha256(caminho):
    digest = hashlib.sha256()
    with open(caminho, "rb") as fh:
        for chunk in iter(lambda: fh.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _aplicar_update(zip_url, versao_nova, sha256_esperado):
    aqui = os.path.dirname(os.path.abspath(__file__))
    zip_path = os.path.join(aqui, "_update_download.zip")
    _baixar(zip_url, zip_path)
    if sha256_esperado and _sha256(zip_path) != sha256_esperado:
        os.remove(zip_path)
        raise RuntimeError("sha256 do ZIP nao bate com o manifest")

    with zipfile.ZipFile(zip_path, "r") as zf:
        for nome in ARQUIVOS_GERENCIADOS:
            zf.extract(f"ifood/{nome}", "/tmp/rotom_update_stage")
            os.replace(f"/tmp/rotom_update_stage/ifood/{nome}", os.path.join(aqui, nome))
    os.remove(zip_path)

    with open(VERSION_FILE, "w", encoding="utf-8") as fh:
        json.dump({"version": versao_nova}, fh)

    _log(f"atualizado pra {versao_nova}, reiniciando")
    os.execv(sys.executable, [sys.executable, os.path.abspath(__file__)])


def _checar_update_uma_vez():
    versao_atual = _versao_atual()
    try:
        payload = _consultar_manifest(versao_atual)
    except (urllib.error.URLError, TimeoutError) as exc:
        _log(f"manifest indisponivel, ignora desta vez: {exc}")
        return

    versao_nova = str(payload.get("version") or "").strip()
    zip_url = str(payload.get("zipUrl") or payload.get("zip_url") or "").strip()
    if not versao_nova or not zip_url or versao_nova == versao_atual:
        return

    _log(f"versao nova disponivel: {versao_nova} (atual: {versao_atual or '?'})")
    _aplicar_update(zip_url, versao_nova, str(payload.get("sha256") or "").strip())


def main():
    print_bridge.start()
    _log(f"print_bridge no ar na porta {print_bridge.PORT}")
    while True:
        _checar_update_uma_vez()
        time.sleep(CHECK_INTERVAL_SECONDS)


if __name__ == "__main__":
    main()
