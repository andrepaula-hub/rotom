#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import hashlib
import json
import os, sys, socket, subprocess, unicodedata, re, random
import shutil
import tempfile
import threading
import queue
import time
import urllib.parse
import urllib.request
import zipfile
from datetime import datetime
import tkinter as tk
import tkinter.font as tkfont
from tkinter import Tk, BOTH, RIGHT, LEFT, Y, X, YES, messagebox, simpledialog, ttk, BooleanVar, StringVar, IntVar
import qrcode
from PIL import Image, ImageDraw, ImageFont, ImageTk

# ========= Google Sheets =========
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from google.auth.transport.requests import Request

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
APP_ROOT = os.path.dirname(BASE_DIR)
INSTALL_BASE_DIR = os.path.dirname(APP_ROOT)
TARGET_APP_DIRNAME = "reboot_QR"
TARGET_APP_ROOT = os.path.join(INSTALL_BASE_DIR, TARGET_APP_DIRNAME)
VERSION_INFO_PATH = os.path.join(BASE_DIR, "rotom_version.json")

def _detect_app_version():
    override = os.environ.get("ROTOM_VERSION", "").strip()
    if override:
        return override
    if os.path.exists(VERSION_INFO_PATH):
        try:
            with open(VERSION_INFO_PATH, "r", encoding="utf-8") as fh:
                info = json.load(fh)
            for key in ("commit", "commitSha", "version"):
                value = str(info.get(key) or "").strip()
                if value:
                    return value[:12]
        except Exception:
            pass
    try:
        proc = subprocess.run(
            ["git", "-C", APP_ROOT, "rev-parse", "--short=12", "HEAD"],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            check=False,
        )
        if proc.returncode == 0:
            value = proc.stdout.decode("utf-8", "ignore").strip()
            if value:
                return value
    except Exception:
        pass
    return "local"

APP_VERSION = _detect_app_version()
UPDATE_PRESERVE = {
    "ifood/token.json",
    "ifood/client_secret.json",
    "ifood/pedidos_ifood_gui_config.json",
}
SHEETS_SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
DRIVE_UPDATE_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.readonly",
]
DEFAULT_TOKEN_PATHS = [
    "/home/pi/reboot_QR/ifood/token.json",
    os.path.join(BASE_DIR, "token.json"),
]
DEFAULT_CLIENT_SECRET_PATHS = [
    "/home/pi/reboot_QR/ifood/client_secret.json",
    os.path.join(BASE_DIR, "client_secret.json"),
]

def _resolve_existing_path(candidates):
    for path in candidates:
        if os.path.exists(path):
            return path
    return candidates[-1]

TOKEN_PATH         = os.environ.get("IFOOD_TOKEN_PATH",         _resolve_existing_path(DEFAULT_TOKEN_PATHS))
CLIENT_SECRET_PATH = os.environ.get("IFOOD_CLIENT_SECRET_PATH", _resolve_existing_path(DEFAULT_CLIENT_SECRET_PATHS))

ERROR_LOG_PATH = os.path.join(BASE_DIR, "app_error.log")

def _log_error(message):
    """Registra erros em arquivo (visível no modo desenvolvedor)."""
    line = f"[{datetime.now().strftime('%d/%m/%Y %H:%M:%S')}] {message}"
    _safe_log(line)
    try:
        with open(ERROR_LOG_PATH, "a", encoding="utf-8") as fh:
            fh.write(line + "\n")
    except Exception:
        pass

def _safe_log(message):
    text = str(message)
    try:
        print(text)
    except UnicodeEncodeError:
        encoding = getattr(sys.stdout, "encoding", None) or "ascii"
        safe_text = text.encode(encoding, errors="replace").decode(encoding, errors="replace")
        print(safe_text)

def _get_google_credentials(scopes):
    creds = None
    if os.path.exists(TOKEN_PATH):
        try:
            creds = Credentials.from_authorized_user_file(TOKEN_PATH, scopes)
        except Exception:
            creds = None
    if not creds or not creds.valid or not creds.has_scopes(scopes):
        if creds and creds.expired and creds.refresh_token and creds.has_scopes(scopes):
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CLIENT_SECRET_PATH, scopes)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_PATH, "w") as token:
            token.write(creds.to_json())
    return creds

def obter_dados_google_sheets(spreadsheet_id, range_name):
    _safe_log(f"[SHEETS] Carregando: {spreadsheet_id} - {range_name}")
    creds = _get_google_credentials(SHEETS_SCOPES)
    service = build('sheets', 'v4', credentials=creds)
    result  = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id, range=range_name).execute()
    values = result.get('values', [])
    _safe_log(f"[SHEETS] {len(values)} linhas recebidas")
    return values


HISTORY_SHEET_NAME = "Histórico de Impressão"
HISTORY_HEADER = ["codigo_shopper", "data_hora", "quantidade_etiquetas", "volumes"]
APOIO_BUSCA_SHEET_NAME = "Apoio Busca"
APOIO_BUSCA_HEADER_CODIGO = "codigo_pedido_shopper"
APOIO_BUSCA_HEADER_OPERADOR = "OPERADOR"
APOIO_BUSCA_HEADER_CRACHA = "CRACHA"

def _get_sheets_service():
    creds = _get_google_credentials(SHEETS_SCOPES)
    return build('sheets', 'v4', credentials=creds)

def _friendly_sheets_error(exc):
    """Traduz erros comuns da API Sheets em orientação prática."""
    s = str(exc)
    if "403" in s or "PERMISSION_DENIED" in s:
        return ("Sem permissão de ESCRITA na planilha (erro 403).\n"
                "A conta Google usada no login precisa ser EDITORA nesta planilha.\n"
                "Peça o acesso de Editor ou use 'Reautenticar Google' nas Configurações "
                "para entrar com uma conta que já seja Editora.")
    if "404" in s or "NOT_FOUND" in s:
        return "Planilha não encontrada (erro 404). Confira o ID da loja selecionada."
    if "invalid_grant" in s or "invalid_scope" in s:
        return ("Sessão do Google expirada ou inválida.\n"
                "Use 'Reautenticar Google' nas Configurações para entrar novamente.")
    return s

def reautenticar_google():
    """Apaga o token salvo e refaz o login no navegador (permite trocar a conta).
    Retorna o objeto de credenciais novo (levanta exceção em caso de falha)."""
    try:
        if TOKEN_PATH and os.path.exists(TOKEN_PATH):
            os.remove(TOKEN_PATH)
    except Exception as e:
        _log_error(f"Falha ao remover token para reautenticação: {e}")
    return _get_google_credentials(SHEETS_SCOPES)

def testar_escrita_planilha(spreadsheet_id):
    """Teste não destrutivo de escrita: garante a aba do histórico e regrava
    o cabeçalho (idempotente). Levanta exceção se a conta não for Editora."""
    service = _get_sheets_service()
    _ensure_history_sheet(service, spreadsheet_id)
    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=f"'{HISTORY_SHEET_NAME}'!A1:D1",
        valueInputOption="RAW",
        body={"values": [HISTORY_HEADER]},
    ).execute()

def _ensure_history_sheet(service, spreadsheet_id):
    """Garante que a aba do histórico exista, com cabeçalho."""
    meta = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    titles = [s["properties"]["title"] for s in meta.get("sheets", [])]
    if HISTORY_SHEET_NAME in titles:
        return
    service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={"requests": [{"addSheet": {"properties": {"title": HISTORY_SHEET_NAME}}}]},
    ).execute()
    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=f"'{HISTORY_SHEET_NAME}'!A1:D1",
        valueInputOption="RAW",
        body={"values": [HISTORY_HEADER]},
    ).execute()

def registrar_historico_impressao(spreadsheet_id, codigo, etiquetas, volumes):
    """Upsert na aba 'Histórico de Impressão' (colunas: código shopper,
    data/hora, quantidade acumulada de etiquetas, volumes da última impressão).
    Concorrência tratada com ler-somar-gravar: relê a linha imediatamente
    antes de gravar, minimizando (sem eliminar) corridas entre duas máquinas.
    Requer que a conta Google logada tenha acesso de EDITOR na planilha."""
    service = _get_sheets_service()
    _ensure_history_sheet(service, spreadsheet_id)
    stamp = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    rng = f"'{HISTORY_SHEET_NAME}'!A:D"
    rows = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id, range=rng).execute().get("values", [])
    row_ix = None
    for i, row in enumerate(rows):
        if i == 0:
            continue  # cabeçalho
        if row and str(row[0]).strip() == str(codigo).strip():
            row_ix = i
            break
    if row_ix is not None:
        try:
            atual = int(str(rows[row_ix][2]).strip() or 0) if len(rows[row_ix]) > 2 else 0
        except Exception:
            atual = 0
        novo = [str(codigo), stamp, str(atual + int(etiquetas)), str(volumes)]
        service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=f"'{HISTORY_SHEET_NAME}'!A{row_ix + 1}:D{row_ix + 1}",
            valueInputOption="RAW",
            body={"values": [novo]},
        ).execute()
    else:
        service.spreadsheets().values().append(
            spreadsheetId=spreadsheet_id,
            range=rng,
            valueInputOption="RAW",
            insertDataOption="INSERT_ROWS",
            body={"values": [[str(codigo), stamp, str(int(etiquetas)), str(volumes)]]},
        ).execute()

def _norm_status(s):
    """Compara status tolerando caixa, espaços e underscores:
    'ARRIVED_AT_ORIGIN' == 'arrived at origin'."""
    s = (s or "").replace("_", " ").upper()
    return " ".join(s.split())

def _parse_status_list(csv_text):
    return {_norm_status(x) for x in (csv_text or "").split(",") if x.strip()}

def ler_historico_volumes(spreadsheet_id):
    """Mapa codigo_shopper -> volumes (última impressão), da aba Histórico."""
    try:
        service = _get_sheets_service()
        rows = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=f"'{HISTORY_SHEET_NAME}'!A:D").execute().get("values", [])
        mapa = {}
        for r in rows[1:]:
            if r and str(r[0]).strip():
                vol = str(r[3]).strip() if len(r) > 3 and r[3] is not None else ""
                mapa[str(r[0]).strip()] = vol
        return mapa
    except Exception:
        return {}

UPDATE_MARKER_PATH = os.path.join(BASE_DIR, "update_marker.json")

def registrar_atualizacao_na_link(spreadsheet_id):
    """Acrescenta na aba LINK uma linha de registro: quando esta máquina foi
    atualizada e para qual versão. Colunas D/E (updated_at, maquina) — a
    coluna zip_url fica vazia, então o leitor do update ignora essas linhas.
    Requer acesso de Editor (mesma exigência do Histórico)."""
    quando = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    try:
        with open(UPDATE_MARKER_PATH, encoding="utf-8") as fh:
            quando = (json.load(fh).get("quando") or quando)
    except Exception:
        pass
    maquina = socket.gethostname() or "?"
    service = _get_sheets_service()
    # garante o cabeçalho estendido (A1:E1) sem tocar no zip_url/version/sha256
    try:
        head = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id, range="'LINK'!A1:E1"
        ).execute().get("values", [[]])
        head = (head[0] + [""] * 5)[:5]
        if head[3] != "updated_at" or head[4] != "maquina":
            head[3], head[4] = "updated_at", "maquina"
            service.spreadsheets().values().update(
                spreadsheetId=spreadsheet_id, range="'LINK'!A1:E1",
                valueInputOption="RAW", body={"values": [head]}).execute()
    except Exception:
        pass
    service.spreadsheets().values().append(
        spreadsheetId=spreadsheet_id,
        range="'LINK'!A:E",
        valueInputOption="RAW",
        insertDataOption="INSERT_ROWS",
        body={"values": [["", APP_VERSION, "", quando, maquina]]},
    ).execute()

def ler_codigos_ja_impressos(spreadsheet_id):
    """Conjunto de códigos com registro no histórico (indicador 'já impresso')."""
    try:
        service = _get_sheets_service()
        rows = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=f"'{HISTORY_SHEET_NAME}'!A:A").execute().get("values", [])
        return {str(r[0]).strip() for r in rows[1:] if r and str(r[0]).strip()}
    except Exception:
        return set()   # aba ainda não existe ou sem acesso: sem indicador

def _norm_operador(nome):
    """Chave de comparação de operador: sem espaços duplicados, maiúsculas."""
    return sanitize_code_exact(nome).upper()

def obter_mapa_apoio_busca(spreadsheet_id):
    """Lê a aba oculta de apoio e devolve DOIS mapas:
    - 'por_codigo':   codigo do pedido -> {operador, cracha}
    - 'por_operador': operador (normalizado) -> cracha
    O mapa por operador é o que garante a busca por crachá mesmo quando a
    aba estiver defasada em relação à dash (códigos antigos)."""
    vazio = {"por_codigo": {}, "por_operador": {}}
    try:
        values = obter_dados_google_sheets(spreadsheet_id, f"'{APOIO_BUSCA_SHEET_NAME}'!A:C")
    except Exception as e:
        _safe_log(f"[APOIO BUSCA] Falha ao ler aba oculta: {e}")
        return vazio
    if not values:
        return vazio
    header = values[0]
    ix_codigo = _find_header_index(header, any_of=[APOIO_BUSCA_HEADER_CODIGO])
    ix_operador = _find_header_index(header, any_of=[APOIO_BUSCA_HEADER_OPERADOR])
    ix_cracha = _find_header_index(header, any_of=[APOIO_BUSCA_HEADER_CRACHA])
    if ix_codigo is None:
        _safe_log("[APOIO BUSCA] Cabeçalho de código não encontrado.")
        return vazio
    por_codigo, por_operador = {}, {}
    for row in values[1:]:
        def get(i):
            return (str(row[i]).strip() if i is not None and i < len(row) and row[i] is not None else "")
        codigo = sanitize_code_exact(get(ix_codigo))
        operador = get(ix_operador)
        cracha = get(ix_cracha)
        if codigo:
            por_codigo[codigo] = {"operador": operador, "cracha": cracha}
        if operador and cracha:
            por_operador[_norm_operador(operador)] = cracha
    _safe_log(f"[APOIO BUSCA] {len(por_codigo)} códigos e "
              f"{len(por_operador)} operadores/crachás carregados")
    return {"por_codigo": por_codigo, "por_operador": por_operador}

def aplicar_apoio_busca_aos_pedidos(pedidos, apoio_map):
    """Anexa o crachá (só em memória) e preenche operador quando vier vazio.
    Passo 1: vínculo exato por CÓDIGO do pedido (também completa o operador).
    Passo 2: para quem ficou sem crachá, vínculo por OPERADOR — assim buscar
    um crachá encontra TODOS os pedidos daquele operador, mesmo que a aba
    Apoio Busca não tenha uma linha para cada pedido atual da dash."""
    if not pedidos or not apoio_map:
        return {"matches": 0, "crachas": 0, "operadores_preenchidos": 0}
    por_codigo = apoio_map.get("por_codigo", {})
    por_operador = apoio_map.get("por_operador", {})
    matches = 0
    crachas = 0
    operadores_preenchidos = 0
    for pedido in pedidos:
        codigo = sanitize_code_exact(pedido.get("codigo"))
        suporte = por_codigo.get(codigo)
        if suporte:
            matches += 1
            cracha = str(suporte.get("cracha") or "").strip()
            operador = str(suporte.get("operador") or "").strip()
            if cracha:
                pedido["_cracha_busca"] = cracha
                crachas += 1
            if operador and not str(pedido.get("operador") or "").strip():
                pedido["operador"] = operador
                operadores_preenchidos += 1
        if not pedido.get("_cracha_busca"):
            cr = por_operador.get(_norm_operador(pedido.get("operador")))
            if cr:
                pedido["_cracha_busca"] = cr
                crachas += 1
    return {
        "matches": matches,
        "crachas": crachas,
        "operadores_preenchidos": operadores_preenchidos
    }
# ==================================


# ============ CONFIG ============
SPREADSHEET_ID = "1aEAqY6vqLGqHt93HsurI_RZPIP6-joi0grPWJsUXtyU"
RANGE_NAME     = "Dash Pedidos a fazer!A:Z"
UPDATE_LINK_RANGE = "LINK!A:Z"

HEADER_CODIGO_SHOPPER = "codigo_pedido_shopper"
HEADER_NUM_IFOOD      = "Nº Pedido IFOOD"
HEADER_STATUS         = "STATUS IFOOD"
HEADER_CLIENTE        = "NOME CLIENTE"
HEADER_COLETA         = "CÓDIGO COLETA"
HEADER_DEVOLUCAO      = "CODIGO DE DEVOLUCAO"
HEADER_TIPO           = "TIPO"
HEADER_OPERADOR       = "OPERADOR"
HEADER_INICIO_BANDA   = "INÍCIO BANDA"
HEADER_FIM_BANDA      = "FIM BANDA"
HEADER_UPDATE_ZIP_URL = "zip_url"
DEFAULT_ROTOM_MANIFEST_URL = "https://alakazam-backend-723462849448.us-east4.run.app/internal/rotom-manifest"

# Lojas (Darks) e suas planilhas — seletor na barra superior
DARK_STORES = {
    "Alto de Pinheiros":  "18V0elH6j9LzGVSYMa4BAOFhKLE5ZIzDXzRyv95A9PEA",
    "Barra Funda":        "1LmNzyzB4Bm2kNuxZh7aczPNmhb4aH9QETjnVAks4UFk",
    "Brooklin":           "1jmygbPURWEJ7m7AVfHVF3cnBMPsyhdNRVSoj2i6V-vI",
    "Campinas":           "1Kp14i6WKnt4aMPvOMHN1JHtqc7D-yJvHULCKYTGPHwE",
    "Consolação":         "1-OOv2BwpzbG1wHitx746U3MUF6JE09FlcD0tJNgzSMU",
    "Higienópolis":       "1d15E8HvDptDsDwZTfSqDZ3gCoEru-PDig_lPr74unuc",
    "Moema":              "1GNdCY-C6A-uw5N-zcFNMJXhTih1TN3Q8DTsiKdqLKfg",
    "Mooca":              "1QQIzFRS6HM-G7bYGV68bC-_tnzg9hHX0BV77iwjCr5M",
    "Morumbi":            "1aEAqY6vqLGqHt93HsurI_RZPIP6-joi0grPWJsUXtyU",
    "Pamplona / Jardins": "1MyU_e2yfhQYnHlyppVI86kB9GZGzJ3KAXYBvQsNIwCM",
    "Pinheiros":          "1ukjC5yfgKFfuDNTxZcxLDXzC-fGz0IPrtOA2hXN6YCI",
    "Riberão":            "1MGOgJsmT3_cvXwsITK-xXJgmW5nMdsOOCLq64AZivu8",
    "São Caetano":        "1YSKrPRpuWRdPXZoc84_98_NIlHNMp5dNfG8mUAZGEbM",
    "Tatuapé":            "1270aj1twvRipYcfaSz-qTr2H55Dupxtcq4Wc-31GUAo",
    "Vila Guilherme":     "155IMHu6NQGdh0mEcGXOYEeJqqGJIxZlz4VkHX4tLB_g",
    "Vila Mariana":       "1UTFa9JaKM3nJlmtp9eDJ8902xIEJ3xup01X6E3dl55U",
    "Vila Olímpia":       "1Joul-YWKrHUw_-O1xu80ntDTtX1xtzNpOBQKCL4AIFk",
}
DARK_STORE_CODES = {
    "Alto de Pinheiros":  "LJ120001",
    "Barra Funda":        "LJ130001",
    "Brooklin":           "LJ160001",
    "Campinas":           "LJ180001",
    "Consolação":         "LJ220001",
    "Higienópolis":       "LJ100001",
    "Moema":              "LJ080001",
    "Mooca":              "LJ240001",
    "Morumbi":            "LJ140001",
    "Pamplona / Jardins": "LJ060001",
    "Pinheiros":          "LJ090001",
    "Riberão":            "LJ230001",
    "São Caetano":        "LJ200001",
    "Tatuapé":            "LJ190001",
    "Vila Guilherme":     "LJ210001",
    "Vila Mariana":       "LJ150001",
    "Vila Olímpia":       "LJ110001",
}
DARK_OUTRA = "Outra…"

PRINTER_HOST    = ""
PRINTER_PORT    = 9100
FORCE_EPL_MODE  = False
AUTO_REFRESH_MS = 60_000
# ================================

CONFIG_PATH = os.path.join(BASE_DIR, "pedidos_ifood_gui_config.json")

# Defaults de layout de etiqueta EPL
# Fonte EPL: 1-5 (bitmap) ou CG Triumvirate (fontes vetoriais).
# Multiplicador: 1-8 (largura x altura).
LABEL_LAYOUT_DEFAULTS = {
    "font_pedido_num":   "4",   # número do pedido
    "font_labels":       "3",   # labels (PEDIDO:, CLIENTE:, etc.)
    "font_coleta_val":   "4",   # valor do código de coleta
    "font_volume":       "5",   # VOLUMES X/Y
    "mult_pedido_num":   "3",   # multiplicador número pedido (largura,altura)
    "mult_volume":       "2",   # multiplicador volume
    "bold_coleta":       "0",   # 1=negrito, 0=normal
    "bold_pedido_num":   "0",   # 1=negrito, 0=normal
    # Nome do cliente na ETIQUETA EXTRA (espaço liberado pela remoção do QR esquerdo)
    "font_cliente":      "4",   # fonte do nome do cliente (etiqueta extra)
    "mult_cliente":      "2",   # multiplicador do nome do cliente (etiqueta extra)
    "bold_cliente":      "0",   # 1=negrito, 0=normal
    "show_codigo_shopper": "0", # 1=mostra o código shopper (texto) abaixo do QR esquerdo
    "show_datetime":     "1",   # 1=carimbo DD/MM HH:MM na etiqueta Extra/Coleta
}

# Padrões de borda rotativos (visualmente distinguem lotes de impressão)
BORDER_PATTERNS = ["preta", "pb", "metade", "branca",
                   "cantos", "dupla", "metade_h", "pontilhada", "barra_topo", "invertida",
                   "meia_invertida", "pontos"]
BORDER_PATTERN_LABELS = {
    "preta":       "Preta (sólida)",
    "pb":          "Preto e branco (tracejada)",
    "metade":      "Metade preto / branco (vertical)",
    "branca":      "Branca (sem marca)",
    "cantos":      "Cantoneiras (cantos em L)",
    "dupla":       "Linha dupla",
    "metade_h":    "Metade preto / branco (horizontal)",
    "pontilhada":  "Pontilhada",
    "barra_topo":  "Faixas (topo e base)",
    "invertida":   "Invertida (fundo preto, texto branco)",
    "meia_invertida": "Meia-invertida (topo preto, texto branco)",
    "pontos": "Fundo pontilhado (trama na etiqueta toda)",
}

# Altura da banda preta do modelo meia-invertida (metade superior da etiqueta)
HALF_INVERT_H = 120

# Espessura da moldura (dots) — configurável
BORDER_THICKNESS_OPTIONS = {"fina": 4, "media": 8, "grossa": 12}
BORDER_THICKNESS_LABELS = {"fina": "Fina", "media": "Média", "grossa": "Grossa"}

DEFAULT_APP_CONFIG = {
    "raspberry_name": socket.gethostname() or "",
    "rotom_manifest_url": DEFAULT_ROTOM_MANIFEST_URL,
    "zebra_rows": "1",  # 1=linhas alternadas coloridas
    "status_colors": "1",  # 1=colorir linhas por status
    "color_separation": "#bfdbfe",   # azul claro — SEPARATION
    "color_cancelled":  "#fca5a5",   # vermelho claro — CANCELLED
    "print_extra_coleta": "1",
    "print_mode": "padrao",  # padrao (da impressora) | direta (OD) | transferencia (O)
    "border_enabled":  "1",        # 1=imprime bordas rotativas nas etiquetas
    "border_mode":     "rotate",   # rotate=cíclico (preta→pb→metade→branca) | random=aleatório
    "border_index":    "0",        # contador de rotação persistente (próximo padrão)
    "border_models":   ",".join(BORDER_PATTERNS),  # modelos habilitados na rotação (csv)
    "border_thickness": "media",   # espessura da moldura: fina | media | grossa
    "batch_ask_volumes":     "1",  # 1=pergunta os volumes de cada pedido na impressão em lote
    "batch_default_volumes": "1",  # volumes usados por pedido quando a pergunta está desativada
    "dark_store":      "Morumbi",  # loja selecionada no seletor de Dark
    "clear_search_after_print": "1",  # 1=limpa o campo de busca após imprimir
    "panel_refresh_seconds": "30",    # polling do Painel (segundos, mín. 10)
    "panel_preparing_statuses": "separation started",  # coluna PREPARANDO (csv)
    "panel_sep_done_statuses": "separation ended",     # marca a separação como concluída (csv)
    "panel_timer_warn_min": "3",   # cronômetro fica laranja a partir de (minutos)
    "panel_timer_alert_min": "5",  # cronômetro fica vermelho a partir de (minutos)
    "panel_going_statuses": "assign driver",       # coluna A CAMINHO: motoboy designado (csv)
    "panel_ready_statuses": "arrived at origin",   # coluna NA LOJA (csv)
    "panel_cancel_statuses": "cancelled, cancellation request",  # dispara alerta de devolução (csv)
    "history_enabled": "1",        # 1=registra impressões na aba Histórico de Impressão
    "dev_mode":        "0",        # 1=modo desenvolvedor (botão de log de erros)
    "spreadsheet_id":        SPREADSHEET_ID,
    "range_name":            RANGE_NAME,
    "header_codigo_shopper": HEADER_CODIGO_SHOPPER,
    "header_num_ifood":      HEADER_NUM_IFOOD,
    "header_status":         HEADER_STATUS,
    "header_cliente":        HEADER_CLIENTE,
    "header_coleta":         HEADER_COLETA,
    "header_devolucao":      HEADER_DEVOLUCAO,
    "header_tipo":           HEADER_TIPO,
    "header_operador":       HEADER_OPERADOR,
    "header_inicio_banda":   HEADER_INICIO_BANDA,
    "header_fim_banda":      HEADER_FIM_BANDA,
    "ui_scale": "115",
    **{f"lbl_{k}": v for k, v in LABEL_LAYOUT_DEFAULTS.items()},
}

UI_SCALE_OPTIONS = {
    "100": {"label": "Padrão",       "font": 10, "rowheight": 24, "heading_font": 10},
    "115": {"label": "Maior",        "font": 12, "rowheight": 30, "heading_font": 11},
    "130": {"label": "Muito maior",  "font": 14, "rowheight": 36, "heading_font": 12},
    "145": {"label": "Grande",       "font": 16, "rowheight": 40, "heading_font": 13},
    "160": {"label": "Muito grande", "font": 17, "rowheight": 44, "heading_font": 14},
    "175": {"label": "Enorme",       "font": 19, "rowheight": 48, "heading_font": 15},
    "200": {"label": "Gigante",      "font": 22, "rowheight": 56, "heading_font": 17},
}

TREE_COLUMN_WIDTHS = {
    "impresso": 46,
    "codigo": 240,
    "num_ifood": 130,
    "cliente": 240,
    "coleta": 90,
    "devolucao": 140,
    "tipo": 120,
    "operador": 140,
    "status": 180,
    "inicio_banda": 110,
    "fim_banda": 110,
}

def load_app_config():
    config = dict(DEFAULT_APP_CONFIG)
    should_persist = False
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                saved = json.load(f)
            if isinstance(saved, dict):
                for key in DEFAULT_APP_CONFIG:
                    value = saved.get(key)
                    if key == "ui_scale":
                        value = str(value).strip() if value is not None else ""
                        if value in UI_SCALE_OPTIONS:
                            config[key] = value
                    elif value is not None:
                        config[key] = str(value).strip() if isinstance(value, str) else str(value)

                # Migra configs antigas que limitavam o range e escondiam colunas novas.
                if config.get("range_name") == "Dash Pedidos a fazer!A:G":
                    config["range_name"] = RANGE_NAME
                    should_persist = True
        except Exception:
            pass
    if should_persist:
        save_app_config(config)
    return config

def save_app_config(config):
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=2)

APP_CONFIG = load_app_config()


def _lbl(key):
    """Retorna valor de layout da config, com fallback para default."""
    return APP_CONFIG.get(f"lbl_{key}", LABEL_LAYOUT_DEFAULTS[key])


def _cfg_value(config, key):
    return config.get(f"lbl_{key}", LABEL_LAYOUT_DEFAULTS[key])


def _cfg_bold(config, key):
    return "bold" if _cfg_value(config, key) == "1" else "normal"


def _cfg_font_size(config, key, base_scale=1.0):
    try:
        font_no = int(_cfg_value(config, key))
    except Exception:
        font_no = 3
    size_map = {1: 8, 2: 10, 3: 12, 4: 15, 5: 18, 6: 20, 7: 22, 8: 24}
    return int(size_map.get(font_no, 12) * base_scale)


def _cfg_mult_size(config, font_key, mult_key, base_scale=1.0):
    try:
        mult = int(_cfg_value(config, mult_key))
    except Exception:
        mult = 1
    mult = max(1, min(mult, 8))
    return _cfg_font_size(config, font_key, base_scale) + (mult - 1) * int(4 * base_scale)


def _make_qr_preview_image(data_text, target_size):
    qr = qrcode.QRCode(
        version=None,
        error_correction=qrcode.constants.ERROR_CORRECT_M,
        box_size=8,
        border=1,
    )
    qr.add_data((data_text or "").encode("utf-8"), optimize=0)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white").convert("L")
    img = img.resize((target_size, target_size), Image.NEAREST)
    return ImageTk.PhotoImage(img)


# Fonte para preview: tenta macOS, Linux, Windows
def _find_preview_font():
    candidates = [
        "/System/Library/Fonts/Helvetica.ttc",         # macOS
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",  # Linux
        "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
        "C:/Windows/Fonts/arial.ttf",                  # Windows
        "C:/Windows/Fonts/calibri.ttf",
        "C:/Windows/Fonts/segoeui.ttf",
    ]
    for p in candidates:
        if os.path.exists(p):
            return p
    return None
PREVIEW_FONT_PATH = _find_preview_font()
PREVIEW_EPL_FONT_SIZES = {
    "1": 14,
    "2": 18,
    "3": 22,
    "4": 30,
    "5": 38,
    "6": 46,
    "7": 54,
    "8": 62,
}


def _epl_bold_from_config(config, key):
    return "B" if _cfg_value(config, key) == "1" else "N"


def _epl_mult_from_config(config, key):
    value = _cfg_value(config, key)
    return f"{value},{value}"


# Largura aproximada (em dots) de cada fonte EPL2, usada para decidir se o
# nome do cliente cabe na etiqueta extra sem invadir o QR de coleta.
EPL_FONT_WIDTHS = {"1": 8, "2": 10, "3": 12, "4": 14, "5": 32, "6": 34, "7": 36, "8": 40}
EPL_FONT_HEIGHTS = {"1": 12, "2": 16, "3": 20, "4": 24, "5": 48, "6": 50, "7": 52, "8": 56}
# Espaço útil para o nome na etiqueta extra: de X=20 até ~X=600 (QR coleta começa em 620)
CLIENTE_AVAIL_DOTS = 584


def _first_last_name(name):
    parts = str(name).split()
    if len(parts) >= 2:
        return f"{parts[0]} {parts[-1]}"
    return str(name).strip()


def _fit_cliente_name(name, font_code, mult):
    """Devolve o nome que cabe na largura disponível da etiqueta extra.
    Se o nome completo não couber, usa apenas o primeiro + último nome.
    Como último recurso, trunca para não colidir com o QR de coleta."""
    name = str(name).strip()
    try:
        mult = max(1, int(mult))
    except Exception:
        mult = 1
    char_w = EPL_FONT_WIDTHS.get(str(font_code), 14) * mult
    max_chars = max(4, CLIENTE_AVAIL_DOTS // char_w)
    if len(name) <= max_chars:
        return name
    short = _first_last_name(name)
    if len(short) <= max_chars:
        return short
    return short[:max_chars].rstrip()


# Espaço útil para o contador de volumes (alinhado à direita, com folga da borda)
VOLUME_RIGHT_X = 762
VOLUME_MIN_X = 470
VOLUME_AVAIL_DOTS = VOLUME_RIGHT_X - VOLUME_MIN_X  # 292
_FONT_ORDER_DESC = ["8", "7", "6", "5", "4", "3", "2", "1"]


def _fit_volume(vol_str, font_code, mult):
    """Ajusta o contador de volumes para caber sem encostar na borda.
    Reduz primeiro a FONTE (mantendo o multiplicador, até no mínimo a fonte 4)
    e só então o multiplicador. Assim '12/15' fica um pouco menor que '1/3',
    mas sem cair para um tamanho minúsculo."""
    try:
        mult = max(1, int(mult))
    except Exception:
        mult = 1
    f = str(font_code)

    def width(fc, m):
        return len(vol_str) * EPL_FONT_WIDTHS.get(str(fc), 32) * m

    # 1) reduz a fonte (mantém o multiplicador) enquanto estiver acima da fonte 4
    while width(f, mult) > VOLUME_AVAIL_DOTS and f not in ("1", "2", "3", "4"):
        idx = _FONT_ORDER_DESC.index(f)
        f = _FONT_ORDER_DESC[idx + 1]
    # 2) se ainda não couber, reduz o multiplicador
    while width(f, mult) > VOLUME_AVAIL_DOTS and mult > 1:
        mult -= 1
    return f, mult


def _volume_y(font_code, mult, label_h=240):
    """Y do contador X/Y centralizado verticalmente na etiqueta (como o QR).
    O rótulo 'VOLUMES:' permanece fixo no topo do bloco."""
    h = EPL_FONT_HEIGHTS.get(str(font_code), 48) * max(1, int(mult))
    y = (label_h - h) // 2
    # nunca sobe a ponto de encostar no rótulo VOLUMES: (termina em ~y=44)
    return max(52, y)


def _build_epl_volume_for_config(config, codigo, num_ifood, cliente, coleta, vol_atual, vol_total):
    cod     = epl_escape(codigo)
    nif     = epl_escape(str(num_ifood))
    cli     = epl_escape(str(cliente)[:22])
    col     = epl_escape(str(coleta))
    vol_str = f"{vol_atual}/{vol_total}"

    f_nif  = _cfg_value(config, "font_pedido_num")
    _mn    = int(_cfg_value(config, "mult_pedido_num") or 1)
    if _cfg_value(config, "bold_pedido_num") == "1": _mn = min(_mn+1, 8)
    m_nif  = f"{_mn},{_mn}"
    f_lbl  = _cfg_value(config, "font_labels")
    f_col  = _cfg_value(config, "font_coleta_val")
    _mc    = 1
    if _cfg_value(config, "bold_coleta") == "1": _mc = 2
    m_col  = f"{_mc},{_mc}"
    f_vol, _mvol = _fit_volume(vol_str, _cfg_value(config, "font_volume"), _cfg_value(config, "mult_volume"))
    m_vol  = f"{_mvol},{_mvol}"
    vol_w  = len(vol_str) * EPL_FONT_WIDTHS.get(str(f_vol), 32) * _mvol
    vol_x  = max(VOLUME_MIN_X, VOLUME_RIGHT_X - vol_w)
    vol_y  = _volume_y(f_vol, _mvol)

    show_cod = _cfg_value(config, "show_codigo_shopper") == "1"
    lq_y, lq_s = (30, "5") if show_cod else (45, "6")

    pattern = config.get("_border_preview")
    inv, frame = _resolve_border(pattern)
    wm = _dots_background_epl_lines() if pattern == "pontos" else ""
    if inv == "full":
        fill = _invert_prefix()
    elif inv == "half":
        fill = _half_fill_prefix()
    else:
        fill = ""
    qr_patch = _qr_white_patch(20, lq_y, cod, lq_s) if (inv or wm) else ""

    texts = []
    def _t(x, y, f, mult, text):
        rev_flag, needs_patch = _text_mode(inv, y, _text_h(f, mult))
        if needs_patch:
            texts.append(_text_patch(x, y, _text_w(text, f, mult), _text_h(f, mult)))
        texts.append(f'A{x},{y},0,{f},{mult},{mult},{rev_flag},"{text}"\n')

    _t(190, 18,  f_lbl, 1,     "IFOOD")
    _t(190, 46,  f_lbl, 1,     "PEDIDO:")
    _t(310, 22,  f_nif, _mn,   nif)
    _t(190, 90,  f_lbl, 1,     "CLIENTE:")
    _t(190, 115, f_lbl, 1,     cli)
    _t(190, 155, f_lbl, 1,     "CODIGO:")
    _t(340, 150, f_col, _mc,   col)
    _t(560, 24,  f_lbl, 1,     "VOLUMES:")
    _t(vol_x, vol_y, f_vol, _mvol, vol_str)
    if show_cod:
        _t(20, 182, "1", 1, cod)

    border = _border_epl_lines(frame, invert=(inv == "full"))
    return (
        "I8,A,001\nN\nq800\nQ240,40\nD8\nS3\nZT\n" + _print_mode_line()
        + fill
        + wm
        + border
        + qr_patch
        + f'b20,{lq_y},Q,m2,s{lq_s},"{cod}"\n'
        + "".join(texts)
        + "P1\n"
    )


def _build_epl_extra_for_config(config, codigo, num_ifood, cliente, coleta):
    nif = epl_escape(str(num_ifood))
    col = epl_escape(str(coleta))

    f_nif = _cfg_value(config, "font_pedido_num")
    _mn   = int(_cfg_value(config, "mult_pedido_num") or 1)
    if _cfg_value(config, "bold_pedido_num") == "1": _mn = min(_mn+1, 8)
    m_nif = f"{_mn},{_mn}"
    f_lbl = _cfg_value(config, "font_labels")
    f_col = _cfg_value(config, "font_coleta_val")
    _mc   = 1
    if _cfg_value(config, "bold_coleta") == "1": _mc = 2
    m_col = f"{_mc},{_mc}"
    f_cli = _cfg_value(config, "font_cliente")
    _mcli = int(_cfg_value(config, "mult_cliente") or 1)
    if _cfg_value(config, "bold_cliente") == "1": _mcli = min(_mcli+1, 8)
    m_cli = f"{_mcli},{_mcli}"
    cli = epl_escape(_fit_cliente_name(cliente, f_cli, _mcli))

    pattern = config.get("_border_preview")
    inv, frame = _resolve_border(pattern)
    wm = _dots_background_epl_lines() if pattern == "pontos" else ""
    if inv == "full":
        fill = _invert_prefix()
    elif inv == "half":
        fill = _half_fill_prefix()
    else:
        fill = ""
    qr_patch = _qr_white_patch(630, 57, col, "6") if (inv or wm) else ""

    texts = []
    def _t(x, y, f, mult, text):
        rev_flag, needs_patch = _text_mode(inv, y, _text_h(f, mult))
        if needs_patch:
            texts.append(_text_patch(x, y, _text_w(text, f, mult), _text_h(f, mult)))
        texts.append(f'A{x},{y},0,{f},{mult},{mult},{rev_flag},"{text}"\n')

    _t(20, 18,  f_lbl, 1,   "IFOOD")
    _t(20, 44,  f_lbl, 1,   "PEDIDO:")
    _t(150, 22, f_nif, _mn, nif)
    _t(20, 82,  f_lbl, 1,   "CLIENTE:")
    _t(20, 106, f_cli, _mcli, cli)
    _t(20, 172, f_lbl, 1,   "CODIGO:")
    _t(150, 164, f_col, _mc, col)

    # Carimbo de impressão DD/MM HH:MM (desligável): canto inferior direito,
    # à esquerda do QR (x até 615; QR começa em 630) e acima da moldura
    # (y=200: termina em 216, folga mesmo com moldura grossa em 220).
    if _cfg_value(config, "show_datetime") == "1":
        stamp = _datetime_stamp()
        stamp_x = 615 - _text_w(stamp, "2", 1)
        col_end = 150 + _text_w(col, f_col, _mc)
        col_bottom = 164 + _text_h(f_col, _mc)
        if col_bottom < 196 or stamp_x >= col_end + 12:   # anti-colisão c/ coleta longa
            _t(stamp_x, 200, "2", 1, stamp)

    border = _border_epl_lines(frame, invert=(inv == "full"))
    return (
        "I8,A,001\nN\nq800\nQ240,40\nD8\nS3\nZT\n" + _print_mode_line()
        + fill
        + wm
        + border
        + qr_patch
        + "".join(texts)
        + f'b630,57,Q,m2,s6,"{col}"\n'
        "P1\n"
    )


def _preview_font(font_code):
    size = PREVIEW_EPL_FONT_SIZES.get(str(font_code), 22)
    if PREVIEW_FONT_PATH and os.path.exists(PREVIEW_FONT_PATH):
        try:
            return ImageFont.truetype(PREVIEW_FONT_PATH, size)
        except Exception:
            pass
    # Fallback: fonte padrão do Pillow
    try:
        return ImageFont.load_default(size=size)
    except Exception:
        return ImageFont.load_default()


def _render_text_bitmap(text, font_code, mult_x, mult_y, bold_flag):
    font = _preview_font(font_code)
    # Canvas generoso para não cortar texto longo ou bold
    temp = Image.new("L", (1400, 260), 0)
    draw = ImageDraw.Draw(temp)
    # Bold no preview: desenhar o texto 3x com pequenos offsets (0,0), (1,0), (0,1)
    # Isso simula negrito sem criar faixa preta
    if bold_flag == "B":
        for dx, dy in ((0, 0), (1, 0), (0, 1)):
            draw.text((4 + dx, 4 + dy), text, font=font, fill=255)
    else:
        draw.text((4, 4), text, font=font, fill=255)
    bbox = temp.getbbox()
    if not bbox:
        return None
    # Padding mínimo no crop para não cortar pixels de borda
    pad = 2
    bbox = (max(0, bbox[0]-pad), max(0, bbox[1]-pad),
            min(temp.width, bbox[2]+pad), min(temp.height, bbox[3]+pad))
    glyph = temp.crop(bbox)
    mult_x = max(1, int(mult_x))
    mult_y = max(1, int(mult_y))
    if mult_x > 1 or mult_y > 1:
        glyph = glyph.resize((glyph.width * mult_x, glyph.height * mult_y), Image.NEAREST)
    return glyph


def _render_qr_bitmap(payload, size_code):
    qr = qrcode.QRCode(
        version=None,
        error_correction=qrcode.constants.ERROR_CORRECT_M,
        box_size=8,
        border=1,
    )
    qr.add_data((payload or "").encode("utf-8"), optimize=0)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white").convert("L")
    target = max(60, 48 + int(size_code) * 12)
    return img.resize((target, target), Image.NEAREST)


def _render_epl_preview_image(epl_text, preview_width):
    label = Image.new("L", (800, 240), 255)
    draw = ImageDraw.Draw(label)
    for raw_line in epl_text.splitlines():
        line = raw_line.strip()
        if not line:
            continue
        box_match = re.match(r'^X(\d+),(\d+),(\d+),(\d+),(\d+)$', line)
        if box_match:
            x0, y0, th, x1, y1 = [int(v) for v in box_match.groups()]
            draw.rectangle([x0, y0, x1, y1], outline=0, width=max(1, th))
            continue
        lo_match = re.match(r'^LO(\d+),(\d+),(\d+),(\d+)$', line)
        if lo_match:
            x, y, ww, hh = [int(v) for v in lo_match.groups()]
            draw.rectangle([x, y, x + ww, y + hh], fill=0)   # preto
            continue
        lw_match = re.match(r'^LW(\d+),(\d+),(\d+),(\d+)$', line)
        if lw_match:
            x, y, ww, hh = [int(v) for v in lw_match.groups()]
            draw.rectangle([x, y, x + ww, y + hh], fill=255)  # branco
            continue
        qr_match = re.match(r'^b(\d+),(\d+),Q,m2,s(\d+),"(.+)"$', line)
        if qr_match:
            x, y, size_code, payload = qr_match.groups()
            qr_img = _render_qr_bitmap(payload, size_code)
            label.paste(qr_img, (int(x), int(y)))
            continue
        text_match = re.match(r'^A(\d+),(\d+),0,(\d+),(\d+),(\d+),([NBR]),"(.+)"$', line)
        if text_match:
            x, y, font_code, mult_x, mult_y, flag, text = text_match.groups()
            bold = "B" if flag == "B" else "N"
            glyph = _render_text_bitmap(text, font_code, mult_x, mult_y, bold)
            if glyph is not None:
                color = 255 if flag == "R" else 0   # R = reverso (texto branco)
                label.paste(color, (int(x), int(y)), glyph)
    scaled_height = int(round(preview_width * (240 / 800)))
    label = label.resize((preview_width, scaled_height), Image.NEAREST)
    return ImageTk.PhotoImage(label)


def _download_update_payload():
    manifest_url = str(APP_CONFIG.get("rotom_manifest_url") or "").strip()
    if manifest_url:
        query = {
            "storeCode": DARK_STORE_CODES.get(APP_CONFIG.get("dark_store", ""), ""),
            "storeName": APP_CONFIG.get("dark_store", ""),
            "deviceName": APP_CONFIG.get("raspberry_name", ""),
            "currentVersion": APP_VERSION,
        }
        sep = "&" if "?" in manifest_url else "?"
        url = manifest_url + sep + urllib.parse.urlencode(query)
        data = json.loads(_read_url_bytes(url, timeout=20).decode("utf-8"))
        zip_url = str(data.get("zipUrl") or data.get("zip_url") or "").strip()
        if zip_url:
            return {
                "zip_url": _normalize_download_url(zip_url),
                "version": str(data.get("version") or "").strip(),
                "sha256": str(data.get("sha256") or "").strip(),
            }

    values = obter_dados_google_sheets(APP_CONFIG["spreadsheet_id"], UPDATE_LINK_RANGE)
    if not values:
        raise RuntimeError("A aba LINK está vazia. Preencha ao menos a coluna zip_url.")

    header = values[0]
    ix_zip = _find_header_index(header, any_of=[HEADER_UPDATE_ZIP_URL])
    ix_ver = _find_header_index(header, any_of=["version"])
    ix_sha = _find_header_index(header, any_of=["sha256"])

    if ix_zip is None:
        raise RuntimeError("A aba LINK precisa ter uma coluna chamada zip_url.")

    for row in values[1:]:
        zip_url = row[ix_zip].strip() if ix_zip < len(row) and row[ix_zip] is not None else ""
        if not zip_url:
            continue
        version = row[ix_ver].strip() if ix_ver is not None and ix_ver < len(row) and row[ix_ver] is not None else ""
        sha256 = row[ix_sha].strip() if ix_sha is not None and ix_sha < len(row) and row[ix_sha] is not None else ""
        return {
            "zip_url": _normalize_download_url(zip_url),
            "version": version,
            "sha256": sha256,
        }

    raise RuntimeError("A aba LINK não possui nenhum valor preenchido na coluna zip_url.")


def _normalize_download_url(url):
    value = (url or "").strip()
    drive_file = re.search(r"/file/d/([a-zA-Z0-9_-]+)", value)
    if drive_file:
        return f"https://drive.google.com/uc?export=download&id={drive_file.group(1)}"
    drive_open = re.search(r"[?&]id=([a-zA-Z0-9_-]+)", value)
    if "drive.google.com" in value and drive_open:
        return f"https://drive.google.com/uc?export=download&id={drive_open.group(1)}"
    return value


def _extract_google_file_id(url):
    value = (url or "").strip()
    patterns = [
        r"/file/d/([a-zA-Z0-9_-]+)",
        r"[?&]id=([a-zA-Z0-9_-]+)",
    ]
    for pattern in patterns:
        match = re.search(pattern, value)
        if match:
            return match.group(1)
    return None


def _download_google_drive_file(file_id, output_path):
    creds = _get_google_credentials(DRIVE_UPDATE_SCOPES)
    drive = build("drive", "v3", credentials=creds)
    request = drive.files().get_media(fileId=file_id)
    with open(output_path, "wb") as fh:
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()


def _read_url_bytes(url, timeout=20):
    try:
        with urllib.request.urlopen(url, timeout=timeout) as response:
            return response.read()
    except Exception:
        proc = subprocess.run(
            ["curl", "-LfsS", url],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            check=False,
        )
        if proc.returncode != 0:
            raise RuntimeError(proc.stderr.decode("utf-8", "ignore") or f"Falha ao ler URL: {url}")
        return proc.stdout


def _download_file(url, output_path):
    drive_file_id = _extract_google_file_id(url) if "drive.google.com" in (url or "") else None
    if drive_file_id:
        _download_google_drive_file(drive_file_id, output_path)
        return
    try:
        with urllib.request.urlopen(url, timeout=60) as response, open(output_path, "wb") as fh:
            shutil.copyfileobj(response, fh)
        return
    except Exception:
        proc = subprocess.run(
            ["curl", "-LfsS", "-o", output_path, url],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            check=False,
        )
        if proc.returncode != 0:
            raise RuntimeError(proc.stderr.decode("utf-8", "ignore") or f"Falha ao baixar arquivo: {url}")


def _sha256_file(path):
    digest = hashlib.sha256()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _detect_extracted_root(extract_dir):
    entries = [os.path.join(extract_dir, name) for name in os.listdir(extract_dir)]
    dirs = [path for path in entries if os.path.isdir(path)]
    files = [path for path in entries if os.path.isfile(path)]
    if len(dirs) == 1 and not files:
        return dirs[0]
    return extract_dir


def _prepare_extracted_root(staging_dir, extract_dir):
    extracted_root = _detect_extracted_root(extract_dir)
    if extracted_root != extract_dir:
        return extracted_root
    normalized_root = os.path.join(staging_dir, "reboot_QR_NOVO")
    os.makedirs(normalized_root, exist_ok=True)
    for name in os.listdir(extract_dir):
        shutil.move(os.path.join(extract_dir, name), os.path.join(normalized_root, name))
    return normalized_root


def _build_updater_script(staging_dir, extracted_root):
    relaunch = os.path.join(TARGET_APP_ROOT, "ifood", "pedidos_ifood_gui.py")
    preserved = "\n".join(
        f'  if [ -f "$PRESERVE_SOURCE/{rel}" ]; then cp "$PRESERVE_SOURCE/{rel}" "{staging_dir}/preserve/{rel}"; fi'
        for rel in sorted(UPDATE_PRESERVE)
    )
    restore = "\n".join(
        f'  if [ -f "{staging_dir}/preserve/{rel}" ]; then mkdir -p "{TARGET_APP_ROOT}/{os.path.dirname(rel)}"; cp "{staging_dir}/preserve/{rel}" "{TARGET_APP_ROOT}/{rel}"; fi'
        for rel in sorted(UPDATE_PRESERVE)
    )
    return f"""#!/bin/zsh
set -e
sleep 2
TARGET_ROOT="{TARGET_APP_ROOT}"
PARENT_DIR="{INSTALL_BASE_DIR}"
PRESERVE_SOURCE="$TARGET_ROOT"
if [ ! -d "$PRESERVE_SOURCE" ]; then
  PRESERVE_SOURCE="{APP_ROOT}"
fi
mkdir -p "{staging_dir}/preserve"
{preserved}
rm -rf "$PARENT_DIR/{TARGET_APP_DIRNAME}_OLD3"
if [ -d "$PARENT_DIR/{TARGET_APP_DIRNAME}_OLD2" ]; then
  mv "$PARENT_DIR/{TARGET_APP_DIRNAME}_OLD2" "$PARENT_DIR/{TARGET_APP_DIRNAME}_OLD3"
fi
if [ -d "$PARENT_DIR/{TARGET_APP_DIRNAME}_OLD1" ]; then
  mv "$PARENT_DIR/{TARGET_APP_DIRNAME}_OLD1" "$PARENT_DIR/{TARGET_APP_DIRNAME}_OLD2"
fi
if [ -d "$TARGET_ROOT" ]; then
  mv "$TARGET_ROOT" "$PARENT_DIR/{TARGET_APP_DIRNAME}_OLD1"
fi
mv "{extracted_root}" "$TARGET_ROOT"
{restore}
echo "{{\"quando\": \"$(date '+%d/%m/%Y %H:%M:%S')\"}}" > "$TARGET_ROOT/ifood/update_marker.json"
cd "$TARGET_ROOT"
nohup python3 "{relaunch}" >/tmp/rebootqr_update.log 2>&1 &
"""


# ========= Parsing / Sort =========
def _normalize(s):
    s = unicodedata.normalize("NFKD", s or "")
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = "".join(ch if ch.isalnum() else " " for ch in s)
    return " ".join(s.upper().split())

def _find_header_index(header_row, must_have_tokens=None, any_of=None):
    norm_headers = [_normalize(h) for h in header_row]
    if any_of:
        any_norm = [_normalize(x) for x in any_of]
        for i, h in enumerate(norm_headers):
            if h in any_norm:
                return i
    if must_have_tokens:
        need = set(_normalize(" ".join(must_have_tokens)).split())
        for i, h in enumerate(norm_headers):
            if need.issubset(set(h.split())):
                return i
    return None

def _parse_horario(texto):
    texto = (texto or "").strip()
    if not texto: return None, ""
    for fmt in ("%H:%M:%S", "%H:%M"):
        try:
            t = datetime.strptime(texto, fmt).time()
            return t, t.strftime("%H:%M")
        except Exception:
            pass
    return None, texto

def sanitize_code_exact(s):
    if s is None: return ""
    s = str(s).strip()
    for ch in ["\u200b", "\u200c", "\u200d", "\ufeff", "\xa0"]:
        s = s.replace(ch, " ")
    s = re.sub(r"[\r\n\t]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def normalize_num_ifood(s):
    value = sanitize_code_exact(s)
    if value.isdigit() and len(value) < 4:
        return value.zfill(4)
    return value

def _status_key(status):
    s = (status or "").strip().upper()
    return (1 if s == "CONCLUDED" else 0, s)

def _inicio_key(pedido):
    dt = pedido.get("inicio_banda_dt")
    return (1, None) if dt is None else (0, dt)

def parse_pedidos(values):
    if not values: return []
    header = values[0]
    _safe_log(f"[PARSE] Header: {header}")

    ix_codigo = _find_header_index(header, any_of=[APP_CONFIG["header_codigo_shopper"]])
    ix_ifood  = _find_header_index(header, any_of=[APP_CONFIG["header_num_ifood"]])
    ix_status = _find_header_index(header, any_of=[APP_CONFIG["header_status"]])
    ix_cliente= _find_header_index(header, any_of=[APP_CONFIG["header_cliente"]])
    ix_coleta = _find_header_index(header, any_of=[APP_CONFIG["header_coleta"]])
    ix_devolucao = _find_header_index(header, any_of=[APP_CONFIG["header_devolucao"]])
    ix_tipo = _find_header_index(header, any_of=[APP_CONFIG["header_tipo"]])
    ix_operador = _find_header_index(header, any_of=[APP_CONFIG["header_operador"]])
    ix_ini    = (_find_header_index(header, must_have_tokens=["INICIO","BANDA"])
                 or _find_header_index(header, must_have_tokens=["INÍCIO","BANDA"])
                 or _find_header_index(header, any_of=[APP_CONFIG["header_inicio_banda"]]))
    ix_fim    = (_find_header_index(header, must_have_tokens=["FIM","BANDA"])
                 or _find_header_index(header, any_of=[APP_CONFIG["header_fim_banda"]]))

    if ix_ini is None and len(header) >= 7: ix_ini = 5
    if ix_fim is None and len(header) >= 8: ix_fim = 6

    _safe_log(f"[PARSE] Indices: codigo={ix_codigo} ifood={ix_ifood} status={ix_status} cliente={ix_cliente} coleta={ix_coleta} devolucao={ix_devolucao} tipo={ix_tipo} operador={ix_operador} ini={ix_ini} fim={ix_fim}")

    pedidos = []
    for _row_ix, row in enumerate(values[1:]):
        def get(i): return (row[i].strip() if i is not None and i < len(row) and row[i] is not None else "")
        codigo_raw = get(ix_codigo)
        if not codigo_raw: continue
        codigo = sanitize_code_exact(codigo_raw)
        ini_dt, ini_str = _parse_horario(get(ix_ini))
        _,      fim_str = _parse_horario(get(ix_fim))
        pedidos.append({
            "_row_ix":         _row_ix,   # posição na planilha = idade do pedido
            "codigo":          codigo,
            "num_ifood":       normalize_num_ifood(get(ix_ifood)),
            "status":          get(ix_status),
            "cliente":         get(ix_cliente),
            "coleta":          get(ix_coleta),
            "devolucao":       get(ix_devolucao),
            "tipo":            get(ix_tipo),
            "operador":        get(ix_operador),
            "inicio_banda":    ini_str,
            "fim_banda":       fim_str,
            "inicio_banda_dt": ini_dt,
        })

    pedidos.sort(key=lambda p: (_status_key(p["status"]), _inicio_key(p)))
    _safe_log(f"[PARSE] {len(pedidos)} pedidos parseados")
    return pedidos
# ==================================


# ========== Impressão EPL ==========
def epl_escape(s):
    return s.replace("\\", "\\\\").replace('"', '\\"')

def _epl_bold(key):
    # EPL2 não tem flag de negrito — sempre retorna N (normal)
    return "N"

def _epl_bold_mult(base_mult_str, bold_key):
    """Quando bold ativo, aumenta multiplicador X em 1 para simular negrito."""
    try:
        m = int(base_mult_str)
    except Exception:
        m = 1
    if _lbl(bold_key) == "1":
        m = min(m + 1, 8)
    return f"{m},{m}"

def _epl_mult(key):
    """Retorna 'n,n' para multiplicador simétrico."""
    m = _lbl(key)
    return f"{m},{m}"

def _border_thickness():
    key = APP_CONFIG.get("border_thickness", "media") if "APP_CONFIG" in globals() else "media"
    return BORDER_THICKNESS_OPTIONS.get(key, 8)

def _border_epl_lines(pattern, w=800, h=240, invert=False):
    """Gera comandos EPL2 de borda. Em modo invert=True as linhas saem brancas
    (LW) para aparecer sobre fundo preto; caso contrário pretas (LO)."""
    if not pattern:
        return ""
    m = 8           # margem da borda em relação à extremidade
    t = _border_thickness()  # espessura da linha (fina/média/grossa)
    x0, y0 = m, m
    x1, y1 = w - m, h - m
    D = "LW" if invert else "LO"     # cor da linha
    lines = []

    def box(xa, ya, xb, yb, th):     # moldura por 4 linhas (compatível c/ branca)
        lines.append(f"{D}{xa},{ya},{xb - xa},{th}")
        lines.append(f"{D}{xa},{yb - th},{xb - xa},{th}")
        lines.append(f"{D}{xa},{ya},{th},{yb - ya}")
        lines.append(f"{D}{xb - th},{ya},{th},{yb - ya}")

    if pattern == "preta":
        box(x0, y0, x1, y1, t)
    elif pattern == "pb":
        dash, gap = 36, 24
        step = dash + gap
        x = x0
        while x < x1:
            seg = min(dash, x1 - x)
            lines.append(f"{D}{x},{y0},{seg},{t}")
            lines.append(f"{D}{x},{y1 - t},{seg},{t}")
            x += step
        y = y0
        while y < y1:
            seg = min(dash, y1 - y)
            lines.append(f"{D}{x0},{y},{t},{seg}")
            lines.append(f"{D}{x1 - t},{y},{t},{seg}")
            y += step
    elif pattern == "metade":
        cx = w // 2
        lines.append(f"{D}{x0},{y0},{cx - x0},{t}")
        lines.append(f"{D}{x0},{y1 - t},{cx - x0},{t}")
        lines.append(f"{D}{x0},{y0},{t},{y1 - y0}")
    elif pattern == "branca":
        pass
    elif pattern == "cantos":
        L = 70
        lines.append(f"{D}{x0},{y0},{L},{t}");          lines.append(f"{D}{x0},{y0},{t},{L}")
        lines.append(f"{D}{x1 - L},{y0},{L},{t}");      lines.append(f"{D}{x1 - t},{y0},{t},{L}")
        lines.append(f"{D}{x0},{y1 - t},{L},{t}");      lines.append(f"{D}{x0},{y1 - L},{t},{L}")
        lines.append(f"{D}{x1 - L},{y1 - t},{L},{t}");  lines.append(f"{D}{x1 - t},{y1 - L},{t},{L}")
    elif pattern == "dupla":
        tt, gap = max(2, t // 2), 10
        box(x0, y0, x1, y1, tt)
        box(x0 + gap, y0 + gap, x1 - gap, y1 - gap, tt)
    elif pattern == "metade_h":
        cy = h // 2
        lines.append(f"{D}{x0},{y0},{x1 - x0},{t}")
        lines.append(f"{D}{x0},{y0},{t},{cy - y0}")
        lines.append(f"{D}{x1 - t},{y0},{t},{cy - y0}")
    elif pattern == "pontilhada":
        dot, step = 8, 34
        x = x0
        while x < x1:
            lines.append(f"{D}{x},{y0},{dot},{t}")
            lines.append(f"{D}{x},{y1 - t},{dot},{t}")
            x += step
        y = y0
        while y < y1:
            lines.append(f"{D}{x0},{y},{t},{dot}")
            lines.append(f"{D}{x1 - t},{y},{t},{dot}")
            y += step
    elif pattern == "barra_topo":
        lines.append(f"{D}{x0},{y0},{x1 - x0},{t}")
        lines.append(f"{D}{x0},{y1 - t},{x1 - x0},{t}")
    elif pattern == "meia_frame":
        # moldura apenas na metade inferior (a banda preta fecha o topo)
        lines.append(f"{D}{x0},{y1 - t},{x1 - x0},{t}")
        lines.append(f"{D}{x0},{HALF_INVERT_H},{t},{y1 - HALF_INVERT_H}")
        lines.append(f"{D}{x1 - t},{HALF_INVERT_H},{t},{y1 - HALF_INVERT_H}")
    return "".join(line + "\n" for line in lines)


def _invert_prefix():
    """Preenche a etiqueta inteira de preto (base para etiqueta invertida)."""
    return "LO0,0,800,240\n"


def _qr_white_patch(x, y, data, s, margin=10):
    """Retângulo branco (LW) atrás do QR, para ele ficar legível sobre fundo preto."""
    try:
        q = qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_M, border=0)
        q.add_data(str(data)); q.make(fit=True)
        size = q.modules_count * int(s)
    except Exception:
        size = 180
    px = max(0, x - margin); py = max(0, y - margin)
    return f"LW{px},{py},{size + 2 * margin},{size + 2 * margin}\n"


def _resolve_border(pattern):
    """Decide o modo de fundo e qual moldura desenhar.
    Retorna (inv, frame): inv ∈ {False, "full", "half"}.
    'invertida'      → fundo preto inteiro, texto branco, moldura branca.
    'meia_invertida' → banda preta só no topo (HALF_INVERT_H), texto do topo
                       em branco; elementos que cruzam a banda ganham janela branca.
    'pontos'         → fundo branco + trama de pontos na etiqueta toda + moldura preta."""
    if pattern == "invertida":
        return "full", "preta"   # invertido + moldura branca sólida
    if pattern == "meia_invertida":
        return "half", "meia_frame"
    if pattern == "pontos":
        return False, "preta"
    return False, pattern


def _half_fill_prefix():
    """Banda preta na metade superior (base do modelo meia-invertida)."""
    return f"LO0,0,800,{HALF_INVERT_H}\n"


def _text_mode(inv, y, h):
    """Para um elemento de texto em (y, altura h), retorna (flag_reverso, precisa_janela).
    full → tudo reverso; half → reverso se couber inteiro na banda preta,
    janela branca (LW) se cruzar a fronteira, normal se estiver abaixo."""
    if inv == "full" or inv is True:
        return "R", False
    if inv == "half":
        if y + h <= HALF_INVERT_H - 2:
            return "R", False
        if y >= HALF_INVERT_H:
            return "N", False
        return "N", True
    return "N", False


def _text_patch(x, y, w, h, margin=6):
    """Janela branca atrás de um texto que cruza a banda preta."""
    px, py = max(0, x - margin), max(0, y - margin)
    return f"LW{px},{py},{w + 2 * margin},{h + 2 * margin}\n"


def _text_w(text, font_code, mult):
    return len(str(text)) * EPL_FONT_WIDTHS.get(str(font_code), 14) * max(1, int(mult))


def _text_h(font_code, mult):
    return EPL_FONT_HEIGHTS.get(str(font_code), 24) * max(1, int(mult))


# ---------- Fundo pontilhado (trama de pontos, 1-bit) ----------
_DOTS_CACHE = {}

def _dots_background_epl_lines(invert=False, step=11, dot=2, w=800, h=240, margin=4):
    """Trama de pontos 2×2 em grade cobrindo a etiqueta inteira (modelo
    'pontos'). Impressão térmica é 1-bit: o efeito 'claro' vem da baixa
    cobertura (~3%%), que não atrapalha os textos por cima; o QR ganha
    janela branca. invert=True gera pontos brancos (LW) para fundo preto.
    Comandos 100%% ASCII — seguros no pipeline CP1252/CRLF (sem GW binário)."""
    key = (invert, step, dot, w, h, margin)
    if key in _DOTS_CACHE:
        return _DOTS_CACHE[key]
    D = "LW" if invert else "LO"
    lines = []
    for y in range(margin, h - margin - dot, step):
        for x in range(margin, w - margin - dot, step):
            lines.append(f"{D}{x},{y},{dot},{dot}")
    out = "".join(line + "\n" for line in lines)
    _DOTS_CACHE[key] = out
    return out


def _datetime_stamp():
    """Carimbo de impressão baseado no relógio local (Raspberry): DD/MM HH:MM."""
    now = datetime.now()
    return f"{now.day:02d}{chr(47)}{now.month:02d} {now.hour:02d}:{now.minute:02d}"

def _print_mode_line():
    """Comando EPL do modo de impressão, conforme a configuração:
    'direta' → OD (térmica direta, sem ribbon); 'transferencia' → O
    (reseta opções = transferência térmica nos modelos TT); 'padrao' →
    nada (mantém o modo configurado na impressora, comportamento antigo)."""
    mode = APP_CONFIG.get("print_mode", "padrao")
    if mode == "direta":
        return "OD\n"
    if mode == "transferencia":
        return "O\n"
    return ""

def build_epl_volume(codigo, num_ifood, cliente, coleta, vol_atual, vol_total, border_pattern=None):
    """Etiqueta normal: QR código + texto + VOLUMES X/Y"""
    cod     = epl_escape(codigo)
    nif     = epl_escape(str(num_ifood))
    cli     = epl_escape(str(cliente)[:22])
    col     = epl_escape(str(coleta))
    vol_str = f"{vol_atual}" + chr(47) + f"{vol_total}"

    f_nif  = _lbl("font_pedido_num");   m_nif  = _epl_bold_mult(_lbl("mult_pedido_num"), "bold_pedido_num")
    f_lbl  = _lbl("font_labels")
    f_col  = _lbl("font_coleta_val");   m_col  = _epl_bold_mult("1", "bold_coleta")
    # Contador de volumes: mantém tamanho grande (auto-ajuste só se realmente não couber)
    # e é alinhado à direita para acomodar números de 2 dígitos (ex.: '12/15').
    f_vol, _mvol = _fit_volume(vol_str, _lbl("font_volume"), _lbl("mult_volume"))
    m_vol  = f"{_mvol},{_mvol}"
    vol_w  = len(vol_str) * EPL_FONT_WIDTHS.get(str(f_vol), 32) * _mvol
    vol_x  = max(VOLUME_MIN_X, VOLUME_RIGHT_X - vol_w)
    vol_y  = _volume_y(f_vol, _mvol)

    # Código shopper opcional, em texto abaixo do QR esquerdo
    show_cod = _lbl("show_codigo_shopper") == "1"
    lq_y, lq_s = (30, "5") if show_cod else (45, "6")

    inv, frame = _resolve_border(border_pattern)
    wm = _dots_background_epl_lines() if border_pattern == "pontos" else ""
    if inv == "full":
        fill = _invert_prefix()
    elif inv == "half":
        fill = _half_fill_prefix()
    else:
        fill = ""
    qr_patch = _qr_white_patch(20, lq_y, cod, lq_s) if (inv or wm) else ""

    _mn_i  = int(str(m_nif).split(",")[0])
    _mc_i  = int(str(m_col).split(",")[0])
    texts = []
    def _t(x, y, f, mult, text):
        rev_flag, needs_patch = _text_mode(inv, y, _text_h(f, mult))
        if needs_patch:
            texts.append(_text_patch(x, y, _text_w(text, f, mult), _text_h(f, mult)))
        texts.append(f'A{x},{y},0,{f},{mult},{mult},{rev_flag},"{text}"\n')

    _t(190, 18,  f_lbl, 1,      "IFOOD")
    _t(190, 46,  f_lbl, 1,      "PEDIDO:")
    _t(310, 22,  f_nif, _mn_i,  nif)
    _t(190, 90,  f_lbl, 1,      "CLIENTE:")
    _t(190, 115, f_lbl, 1,      cli)
    _t(190, 155, f_lbl, 1,      "CODIGO:")
    _t(340, 150, f_col, _mc_i,  col)
    _t(560, 24,  f_lbl, 1,      "VOLUMES:")
    _t(vol_x, vol_y, f_vol, _mvol, vol_str)
    if show_cod:
        _t(20, 182, "1", 1, cod)

    border = _border_epl_lines(frame, invert=(inv == "full"))
    return (
        "I8,A,001\nN\nq800\nQ240,40\nD8\nS3\nZT\n" + _print_mode_line()
        + fill
        + wm
        + border
        + qr_patch
        + f'b20,{lq_y},Q,m2,s{lq_s},"{cod}"\n'
        + "".join(texts)
        + "P1\n"
    )

def build_epl_extra(codigo, num_ifood, cliente, coleta, border_pattern=None):
    """Etiqueta extra (coleta): SEM QR esquerdo, nome do cliente ampliado + QR coleta à direita"""
    nif = epl_escape(str(num_ifood))
    col = epl_escape(str(coleta))

    f_nif = _lbl("font_pedido_num");  m_nif = _epl_bold_mult(_lbl("mult_pedido_num"), "bold_pedido_num")
    f_lbl = _lbl("font_labels")
    f_col = _lbl("font_coleta_val");  m_col = _epl_bold_mult("1", "bold_coleta")
    f_cli = _lbl("font_cliente")
    _mcli = int(_lbl("mult_cliente") or 1)
    if _lbl("bold_cliente") == "1": _mcli = min(_mcli + 1, 8)
    m_cli = f"{_mcli},{_mcli}"
    cli = epl_escape(_fit_cliente_name(cliente, f_cli, _mcli))

    inv, frame = _resolve_border(border_pattern)
    wm = _dots_background_epl_lines() if border_pattern == "pontos" else ""
    if inv == "full":
        fill = _invert_prefix()
    elif inv == "half":
        fill = _half_fill_prefix()
    else:
        fill = ""
    qr_patch = _qr_white_patch(630, 57, col, "6") if (inv or wm) else ""

    _mn_i = int(str(m_nif).split(",")[0])
    _mc_i = int(str(m_col).split(",")[0])
    texts = []
    def _t(x, y, f, mult, text):
        rev_flag, needs_patch = _text_mode(inv, y, _text_h(f, mult))
        if needs_patch:
            texts.append(_text_patch(x, y, _text_w(text, f, mult), _text_h(f, mult)))
        texts.append(f'A{x},{y},0,{f},{mult},{mult},{rev_flag},"{text}"\n')

    _t(20, 18,  f_lbl, 1,     "IFOOD")
    _t(20, 44,  f_lbl, 1,     "PEDIDO:")
    _t(150, 22, f_nif, _mn_i, nif)
    _t(20, 82,  f_lbl, 1,     "CLIENTE:")
    _t(20, 106, f_cli, _mcli, cli)
    _t(20, 172, f_lbl, 1,     "CODIGO:")
    _t(150, 164, f_col, _mc_i, col)

    # Carimbo de impressão DD/MM HH:MM (desligável): canto inferior direito,
    # à esquerda do QR e acima da moldura (ver builder de config).
    if _lbl("show_datetime") == "1":
        stamp = _datetime_stamp()
        stamp_x = 615 - _text_w(stamp, "2", 1)
        col_end = 150 + _text_w(col, f_col, _mc_i)
        col_bottom = 164 + _text_h(f_col, _mc_i)
        if col_bottom < 196 or stamp_x >= col_end + 12:
            _t(stamp_x, 200, "2", 1, stamp)

    border = _border_epl_lines(frame, invert=(inv == "full"))
    return (
        "I8,A,001\nN\nq800\nQ240,40\nD8\nS3\nZT\n" + _print_mode_line()
        + fill
        + wm
        + border
        + qr_patch
        + "".join(texts)
        + f'b630,57,Q,m2,s6,"{col}"\n'
        "P1\n"
    )

def build_epl_border_test(index, total, pattern, label_name):
    """Etiqueta de teste: aplica a borda e identifica o modelo pelo nome.
    Não usa dados reais nem mexe na rotação — serve só para visualizar os modelos."""
    big  = epl_escape((pattern or "SEM BORDA").upper())
    name = epl_escape(str(label_name))
    inv, frame = _resolve_border(pattern)
    wm = _dots_background_epl_lines() if pattern == "pontos" else ""
    if inv == "full":
        fill = _invert_prefix()
    elif inv == "half":
        fill = _half_fill_prefix()
    else:
        fill = ""

    texts = []
    def _t(x, y, f, mult, text):
        rev_flag, needs_patch = _text_mode(inv, y, _text_h(f, mult))
        if needs_patch:
            texts.append(_text_patch(x, y, _text_w(text, f, mult), _text_h(f, mult)))
        texts.append(f'A{x},{y},0,{f},{mult},{mult},{rev_flag},"{text}"\n')

    _t(40, 28,  "3", 1, f"TESTE DE BORDA  {index}/{total}")
    _t(40, 78,  "4", 2, big)
    _t(40, 168, "3", 1, name)

    border = _border_epl_lines(frame, invert=(inv == "full"))
    return (
        "I8,A,001\nN\nq800\nQ240,40\nD8\nS3\nZT\n" + _print_mode_line()
        + fill
        + wm
        + border
        + "".join(texts)
        + "P1\n"
    )

def _cups_default_exists():
    try:
        out = subprocess.run(["lpstat", "-d"], stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=False, text=True)
        return "system default destination:" in out.stdout or "destino padrão" in out.stdout.lower()
    except Exception:
        return False

def _normalize_eol(payload, use_lf=False):
    if not use_lf:
        payload = payload.replace("\n", "\r\n")
    # CP1252 (Latin Ocidental) preserva acentos do português (ã, õ, ç, á, é...).
    # Deve casar com o codepage do comando EPL 'I8,A,001'.
    try:
        return payload.encode("cp1252", errors="replace")
    except Exception:
        return payload.encode("ascii", errors="ignore")

def _printer_available():
    """Checagem leve de disponibilidade da impressora (para o indicador da UI).
    Com PRINTER_HOST: testa conexão TCP real (porta 9100).
    Sem host: verifica se há impressora padrão no SO (CUPS/Windows) — indica
    'configurada', não garante que está ligada."""
    try:
        if PRINTER_HOST:
            with socket.create_connection((PRINTER_HOST, PRINTER_PORT), timeout=2):
                return True
        if sys.platform.startswith("win"):
            try:
                import win32print
                return bool(win32print.GetDefaultPrinter())
            except Exception:
                return False
        return _cups_default_exists()
    except Exception:
        return False

def send_to_printer(epl_text, use_lf=False):
    raw = _normalize_eol(epl_text, use_lf=use_lf)
    if PRINTER_HOST:
        with socket.create_connection((PRINTER_HOST, PRINTER_PORT), timeout=5) as s:
            s.sendall(raw)
        return
    if sys.platform.startswith("win"):
        try:
            import win32print
        except Exception as e:
            raise RuntimeError("win32print não disponível") from e
        printer_name = win32print.GetDefaultPrinter()
        hPrinter = win32print.OpenPrinter(printer_name)
        try:
            win32print.StartDocPrinter(hPrinter, 1, ("EPL job", None, "RAW"))
            win32print.StartPagePrinter(hPrinter)
            win32print.WritePrinter(hPrinter, raw)
            win32print.EndPagePrinter(hPrinter)
            win32print.EndDocPrinter(hPrinter)
        finally:
            win32print.ClosePrinter(hPrinter)
        return
    if not _cups_default_exists():
        raise RuntimeError("Nenhuma impressora padrão no CUPS.")
    proc = subprocess.run(["lp", "-o", "raw"], input=raw, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    if proc.returncode != 0:
        raise RuntimeError(f"Erro impressão: {proc.stderr.decode('utf-8', 'ignore')}")
# ===================================


# ============= UI Tkinter ===========
def m_cfg(key, default):
    """Leitura de config para componentes do painel."""
    try:
        return APP_CONFIG.get(key, default)
    except Exception:
        return default

SHOPPER_GREEN = "#00A651"
SHOPPER_GREEN_DARK = "#00753a"
PANEL_BG = "#f4f6f5"

class PainelFrame(tk.Frame):
    """Telão de coleta estilo Kanban: 'A CAMINHO' (motoboy indo à loja) e
    'NA LOJA' (motoboy no local). Somente leitura; sem impressão/escrita."""

    def __init__(self, master, app, fullscreen=False):
        super().__init__(master, bg=PANEL_BG)
        self.app = app
        self.fullscreen = fullscreen
        big = 1.35 if fullscreen else 1.0
        self._f_num  = tkfont.Font(family="Helvetica", size=int(44 * big), weight="bold")
        self._f_vol  = tkfont.Font(family="Helvetica", size=int(17 * big), weight="bold")
        self._f_head = tkfont.Font(family="Helvetica", size=int(22 * big), weight="bold")
        self._f_col  = tkfont.Font(family="Helvetica", size=int(18 * big), weight="bold")
        self._f_info = tkfont.Font(family="Helvetica", size=int(12 * big))
        self._f_timer = tkfont.Font(family="Helvetica", size=int(15 * big), weight="bold")
        self._last_signature = None
        self._timer_labels = []   # [(Label, timestamp_de_entrada)] — cards da PREPARANDO

        # ── Cabeçalho ──
        head = tk.Frame(self, bg=SHOPPER_GREEN)
        head.pack(fill=X)
        self.title_lbl = tk.Label(head, text="PAINEL DE COLETA", bg=SHOPPER_GREEN,
                                  fg="white", font=self._f_head, padx=16, pady=10)
        self.title_lbl.pack(side=LEFT)
        tk.Label(head, text=APP_VERSION, bg=SHOPPER_GREEN, fg="#bfe9d2",
                 font=self._f_info).pack(side=LEFT)
        self.clock_lbl = tk.Label(head, text="--:--:--", bg=SHOPPER_GREEN,
                                  fg="white", font=self._f_head, padx=16)
        self.clock_lbl.pack(side=RIGHT)
        if not fullscreen:
            tk.Button(head, text="Modo painel (F11)", command=app.toggle_panel_fullscreen,
                      bg=SHOPPER_GREEN_DARK, fg="white", activebackground=SHOPPER_GREEN_DARK,
                      activeforeground="white", relief="flat", padx=10,
                      font=self._f_info).pack(side=RIGHT, padx=8, pady=8)
        self.status_lbl = tk.Label(head, text="", bg=SHOPPER_GREEN, fg="#d7ffe9",
                                   font=self._f_info, padx=10)
        self.status_lbl.pack(side=RIGHT)

        # ── Colunas Kanban (PREPARANDO → A CAMINHO → NA LOJA) ──
        self.COLUMNS = [
            # (chave, título, cor cabeçalho, fg cabeçalho, estilo do card, mostra volumes)
            ("preparing", "PREPARANDO", "#dfe7e2", "#4b5f54", "outline_gray", False),
            ("going",     "A CAMINHO",  "#e7efe9", SHOPPER_GREEN_DARK, "outline_green", True),
            ("ready",     "NA LOJA",    SHOPPER_GREEN, "white", "solid_green", True),
        ]
        body = tk.Frame(self, bg=PANEL_BG)
        body.pack(fill=BOTH, expand=YES, padx=12, pady=12)
        body.grid_rowconfigure(1, weight=1)
        self.count_lbls = {}
        self.areas = {}
        self.card_styles = {}
        n = len(self.COLUMNS)
        self.card_show_vol = {}
        for i, (key, title, hbg, hfg, style, show_vol) in enumerate(self.COLUMNS):
            body.grid_columnconfigure(i, weight=1, uniform="kan")
            padx = (0, 6) if i == 0 else ((6, 0) if i == n - 1 else (6, 6))
            lbl = tk.Label(body, text=f"{title} • 0", bg=hbg, fg=hfg,
                           font=self._f_col, pady=8)
            lbl.grid(row=0, column=i, sticky="we", padx=padx)
            self.count_lbls[key] = (lbl, title)
            self.areas[key] = self._make_scroll_column(body, i, padx)
            self.card_styles[key] = style
            self.card_show_vol[key] = show_vol

        # overlay de alerta de cancelamento (fica por cima do Kanban)
        self._alert_overlay = None

        self._tick_clock()

    def _make_scroll_column(self, body, col, padx):
        wrap = tk.Frame(body, bg=PANEL_BG)
        wrap.grid(row=1, column=col, sticky="nsew", padx=padx)
        canvas = tk.Canvas(wrap, bg=PANEL_BG, highlightthickness=0)
        vsb = ttk.Scrollbar(wrap, orient="vertical", command=canvas.yview)
        inner = tk.Frame(canvas, bg=PANEL_BG)
        inner_id = canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.configure(yscrollcommand=vsb.set)
        inner.bind("<Configure>", lambda _e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda e: canvas.itemconfigure(inner_id, width=e.width))
        canvas.pack(side=LEFT, fill=BOTH, expand=YES)
        vsb.pack(side=RIGHT, fill=Y)
        return inner

    @staticmethod
    def _fmt_elapsed(ts):
        s = max(0, int(time.time() - ts))
        return f"{s // 60:02d}:{s % 60:02d}"

    def _tick_clock(self):
        try:
            self.clock_lbl.config(text=datetime.now().strftime("%H:%M:%S"))
            for lbl, ts in self._timer_labels:
                try:
                    lbl.config(text=f"⏱ {self._fmt_elapsed(ts)}",
                               fg=self._timer_color(ts))
                except Exception:
                    pass  # card já reconstruído
            self.after(1000, self._tick_clock)
        except Exception:
            pass  # frame destruído

    _CARD_LOOKS = {
        "solid_green":   {"bg": SHOPPER_GREEN, "fg": "white", "border": 0},
        "outline_green": {"bg": "white", "fg": SHOPPER_GREEN_DARK, "border": 3},
        "outline_gray":  {"bg": "white", "fg": "#4b5f54", "border": 3,
                          "border_color": "#9fb3a8"},
    }

    TIMER_COLORS = {"ok": "#7a8c81", "warn": "#e67e22", "alert": "#c62828"}

    def _timer_color(self, ts):
        """Cinza até o 1º limite, laranja até o 2º, vermelho depois (config)."""
        try:
            warn = max(1, int(m_cfg("panel_timer_warn_min", "3"))) * 60
        except Exception:
            warn = 180
        try:
            alert = max(1, int(m_cfg("panel_timer_alert_min", "5"))) * 60
        except Exception:
            alert = 300
        s = time.time() - ts
        if s >= alert:
            return self.TIMER_COLORS["alert"]
        if s >= warn:
            return self.TIMER_COLORS["warn"]
        return self.TIMER_COLORS["ok"]

    def _card(self, parent, num_ifood, volumes, style, show_vol=True,
              since_ts=None, operador=""):
        look = self._CARD_LOOKS[style]
        card = tk.Frame(parent, bg=look["bg"], bd=0,
                        highlightbackground=look.get("border_color", SHOPPER_GREEN),
                        highlightthickness=look["border"])
        card.pack(fill=X, pady=4 if not show_vol else 6, ipady=2 if not show_vol else 4)
        tk.Label(card, text=str(num_ifood) or "—", bg=look["bg"], fg=look["fg"],
                 font=self._f_num).pack(pady=(6, 0) if not show_vol else (10, 0))
        if show_vol:
            if str(volumes).strip():
                tk.Label(card, text=f"{volumes} volume(s)", bg=look["bg"], fg=look["fg"],
                         font=self._f_vol).pack(pady=(0, 10))
            else:
                # sem volumes no histórico = etiquetas não impressas =
                # pedido ainda em separação, apesar do motoboy a caminho/na loja
                tk.Label(card, text="SEPARANDO", bg="#e67e22", fg="white",
                         font=self._f_vol, padx=14, pady=2).pack(pady=(2, 10))
        else:
            if str(operador).strip():
                # quem está preparando o pedido (primeiro nome do operador)
                tk.Label(card, text=str(operador), bg=look["bg"], fg="#4b5f54",
                         font=self._f_vol).pack(pady=(0, 0))
            if since_ts is not None:
                # cronômetro do tempo na coluna PREPARANDO (mm:ss, vivo,
                # com cor por limite: cinza → laranja → vermelho)
                lbl = tk.Label(card, text=f"⏱ {self._fmt_elapsed(since_ts)}",
                               bg=look["bg"], fg=self._timer_color(since_ts),
                               font=self._f_timer)
                lbl.pack(pady=(0, 6))
                self._timer_labels.append((lbl, since_ts))

    def update_data(self, columns_data, info_text=""):
        """columns_data: dict chave->lista de (num_ifood, volumes, ts_entrada).
        Chaves: preparing, going, ready. Reconstrói só se mudou."""
        signature = tuple(tuple(columns_data.get(k, ())) for k, *_ in self.COLUMNS)
        self.status_lbl.config(text=info_text)
        if signature == self._last_signature:
            return
        self._last_signature = signature
        self._timer_labels = []
        for key, *_ in self.COLUMNS:
            items = columns_data.get(key, [])
            area = self.areas[key]
            for w in area.winfo_children():
                w.destroy()
            if not items:
                tk.Label(area, text="— vazio —", bg=PANEL_BG, fg="#8aa596",
                         font=self._f_vol).pack(pady=30)
            else:
                for num, vol, ts, oper in items:
                    self._card(area, num, vol, self.card_styles[key],
                               show_vol=self.card_show_vol[key], since_ts=ts,
                               operador=oper)
            lbl, title = self.count_lbls[key]
            lbl.config(text=f"{title} • {len(items)}")

    def set_alerts(self, alerts):
        """Mostra/oculta o aviso de pedidos cancelados (overlay vermelho).
        alerts: lista de dicts {num, volumes, hora}. Some só com 'CIENTE'."""
        if self._alert_overlay is not None:
            try:
                self._alert_overlay.destroy()
            except Exception:
                pass
            self._alert_overlay = None
        if not alerts:
            return
        ov = tk.Frame(self, bg="#c62828", bd=0,
                      highlightbackground="white", highlightthickness=4)
        # sem altura no place: o overlay assume a altura natural do conteúdo
        ov.place(relx=0.5, rely=0.5, anchor="center", relwidth=0.72)
        tk.Label(ov, text="⚠  PEDIDO CANCELADO — DEVOLVER!", bg="#c62828",
                 fg="white", font=self._f_head).pack(pady=(18, 6), padx=24)
        for a in alerts[-6:]:   # mostra até 6, mais recentes
            vol = f" • {a['volumes']} volume(s)" if str(a.get("volumes") or "").strip() else ""
            tk.Label(ov, text=f"Nº {a['num']}{vol}  ({a['hora']})", bg="#c62828",
                     fg="white", font=self._f_vol).pack(pady=2)
        if len(alerts) > 6:
            tk.Label(ov, text=f"… e mais {len(alerts) - 6} pedido(s)", bg="#c62828",
                     fg="#ffd7d7", font=self._f_info).pack(pady=2)
        tk.Button(ov, text="CIENTE — fechar aviso", command=self.app.ack_cancel_alerts,
                  bg="white", fg="#c62828", activebackground="#ffecec",
                  activeforeground="#c62828", relief="flat", padx=18, pady=8,
                  font=self._f_col).pack(pady=(12, 18))
        ov.lift()
        self._alert_overlay = ov


class App(Tk):
    def __init__(self):
        super().__init__()
        self.title(f"Rotom • Pedidos iFood — {APP_VERSION}")
        self.geometry("1340x740")
        self.style = ttk.Style(self)
        self.tree_font = None
        self.tree_heading_font = None
        self._apply_ui_scale()
        self._is_reloading    = False
        self._all_pedidos     = []
        self._filtered_pedidos = []
        self.search_var       = StringVar()
        self.status_filter_var = StringVar(value="(todos)")
        self._sort_state      = {"col": None, "reverse": False}
        self._printing        = False

        # ── Notebook principal: abas Pedidos e Painel ──
        self.main_nb = ttk.Notebook(self)
        self.tab_pedidos = ttk.Frame(self.main_nb)
        self.main_nb.add(self.tab_pedidos, text="  Pedidos  ")

        # ── Top bar (dentro da aba Pedidos) ──
        top = ttk.Frame(self.tab_pedidos, padding=8)
        top.pack(fill=X)
        ttk.Button(top, text="Atualizar",     command=self.reload).pack(side=LEFT, padx=(0, 6))
        ttk.Button(top, text="Configurações", command=self.open_settings).pack(side=LEFT, padx=(0, 6))
        self.auto_var = BooleanVar(value=True)
        ttk.Checkbutton(top, text="Auto", variable=self.auto_var,
                        command=self._on_toggle_auto).pack(side=LEFT)
        ttk.Label(top, text="Loja:").pack(side=LEFT, padx=(10, 4))
        self.dark_var = StringVar(value=self._initial_dark_name())
        self.dark_cb = ttk.Combobox(top, textvariable=self.dark_var,
                                    values=list(DARK_STORES.keys()) + [DARK_OUTRA],
                                    state="readonly", width=14)
        self.dark_cb.pack(side=LEFT)
        self.dark_cb.bind("<<ComboboxSelected>>", self._on_dark_selected)
        ttk.Label(top, text="Buscar:").pack(side=LEFT, padx=(14, 4))
        self.search_entry = ttk.Entry(top, textvariable=self.search_var, width=13)
        self.search_entry.pack(side=LEFT)
        ttk.Button(top, text="Limpar",       command=self.clear_search).pack(side=LEFT, padx=(4, 0))
        ttk.Label(top, text="Status:").pack(side=LEFT, padx=(10, 4))
        self.status_filter_cb = ttk.Combobox(top, textvariable=self.status_filter_var,
                                             values=["(todos)"], state="readonly", width=12)
        self.status_filter_cb.pack(side=LEFT)
        self.status_filter_var.trace_add("write", self._on_search_change)
        ttk.Button(top, text="Imprimir selecionado", command=self.print_selected).pack(side=LEFT, padx=(12, 0))
        self.search_var.trace_add("write", self._on_search_change)

        # ── Barra de status (inferior) ──
        statusbar = ttk.Frame(self, padding=(8, 4))
        statusbar.pack(side="bottom", fill=X)
        self.status_lbl = ttk.Label(statusbar, text="Carregando…")
        self.status_lbl.pack(side=LEFT)
        # indicador da impressora substituído pela versão do programa
        # (o label continua existindo, sem pack, para compatibilidade)
        self.printer_lbl = ttk.Label(statusbar, text="")
        self.version_lbl = ttk.Label(statusbar, text=APP_VERSION,
                                     foreground="#00A651")
        self.version_lbl.pack(side=RIGHT)
        self.devlog_btn = ttk.Button(statusbar, text="Log de erros", command=self._open_error_log)
        self._refresh_devlog_button()
        self.progress = ttk.Progressbar(statusbar, orient="horizontal", mode="determinate", length=180)
        # (a barra de progresso só é exibida durante jobs grandes)

        # ── Atalhos de teclado ──
        self.bind("<F5>",       lambda _e: self.reload())
        self.bind("<Return>",   lambda _e: self.print_selected())
        self.bind("<KP_Enter>", lambda _e: self.print_selected())
        self.bind("<Control-f>", lambda _e: self._focus_search())
        self.bind("<Control-F>", lambda _e: self._focus_search())

        # ── Tabela ──
        table_wrap = ttk.Frame(self.tab_pedidos, padding=(8, 0, 8, 8))
        table_wrap.pack(fill=BOTH, expand=YES)
        columns = ("impresso", "codigo", "num_ifood", "cliente", "coleta", "devolucao", "tipo", "operador", "status", "inicio_banda", "fim_banda")
        self._tree_columns = columns
        self.tree = ttk.Treeview(table_wrap, columns=columns, show="headings", selectmode="extended")
        self._tree_headers = {"impresso": "IMP", "codigo": "CÓDIGO", "num_ifood": "PEDIDO IFOOD", "cliente": "CLIENTE",
                   "coleta": "COLETA", "devolucao": "DEVOLUÇÃO", "tipo": "TIPO", "operador": "OPERADOR", "status": "STATUS", "inicio_banda": "INÍCIO BANDA", "fim_banda": "FIM BANDA"}
        headers = self._tree_headers
        for col in columns:
            self.tree.heading(col, text=headers[col],
                              command=lambda c=col: self._on_heading_click(c))
            self.tree.column(col, width=TREE_COLUMN_WIDTHS[col],
                             anchor="center" if col in ("impresso", "inicio_banda", "fim_banda") else "w")
        scroll = ttk.Scrollbar(table_wrap, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.pack(side=LEFT, fill=BOTH, expand=YES)
        scroll.pack(side=RIGHT, fill=Y)
        self.tree.tag_configure("concluded",         foreground="#777777")
        self._apply_tree_colors()
        self.tree.bind("<Double-1>", lambda _: self.print_selected())

        # ── Aba Painel (telão de coleta) ──
        self.tab_painel = ttk.Frame(self.main_nb)
        self.main_nb.add(self.tab_painel, text="  Painel  ")
        self.panel_main = PainelFrame(self.tab_painel, self)
        self.panel_main.pack(fill=BOTH, expand=YES)
        self.main_nb.pack(fill=BOTH, expand=YES)
        self.main_nb.bind("<<NotebookTabChanged>>", self._on_main_tab_changed)

        # estado do painel
        self._panel_frames = [self.panel_main]
        self._panel_fullscreen_win = None
        self._panel_last_fetch = 0.0
        self._panel_loading = False
        self._history_volumes = {}
        # migração: o padrão antigo de "a caminho" era 'going to origin';
        # o mapeamento correto validado em loja é 'assign driver'
        if _norm_status(APP_CONFIG.get("panel_going_statuses", "")) == "GOING TO ORIGIN":
            APP_CONFIG["panel_going_statuses"] = "assign driver"
            try:
                save_app_config(dict(APP_CONFIG))
            except Exception:
                pass
        self._panel_prev_status = {}   # codigo -> status normalizado (p/ detectar transições)
        # memória por pedido: o status na planilha oscila entre a trilha da
        # separação e a do motoboy; estas flags são monotônicas (só ligam)
        self._panel_flags = {}         # codigo -> {"sep": bool, "assigned": bool, "arrived": bool}
        self._panel_flags_seeded = False
        self._panel_col_since = {}     # codigo -> (coluna, timestamp de entrada) p/ cronômetro
        self._cancel_alerts = []       # avisos de devolução pendentes de "CIENTE"
        self.bind("<F11>", lambda _e: self.toggle_panel_fullscreen())

        self._pending_update_report = os.path.exists(UPDATE_MARKER_PATH)
        self.reload()
        self.after(50, self.tree.focus_set)
        self.after(1_000, self._auto_tick)
        self.after(1_000, self._panel_heartbeat)
        self._ui_queue = queue.Queue()
        self.after(100, self._process_ui_queue)
        # monitor da impressora desativado (indicador saiu da barra de status)
        # self.after(500, self._start_printer_monitor)

    # ── auto-refresh ──
    def _auto_tick(self):
        try:
            # com o Painel ativo, quem atualiza é o polling dele (sem duplicar API)
            if self.auto_var.get() and not self._is_reloading and not self._panel_active():
                self.reload()
        finally:
            self.after(AUTO_REFRESH_MS, self._auto_tick)

    def _on_toggle_auto(self):
        self.status_lbl.config(text="Auto ligado (1 min)" if self.auto_var.get() else "Auto desligado")

    def _apply_tree_colors(self):
        """Aplica tags de cor na tree a partir do APP_CONFIG."""
        c_sep = APP_CONFIG.get("color_separation", "#bfdbfe")
        c_can = APP_CONFIG.get("color_cancelled",  "#fca5a5")
        # Derivar versão escura (odd) misturando levemente com cinza)
        def darken(hex_color):
            hex_color = hex_color.lstrip("#")
            r,g,b = int(hex_color[0:2],16), int(hex_color[2:4],16), int(hex_color[4:6],16)
            r,g,b = max(0,r-30), max(0,g-30), max(0,b-30)
            return f"#{r:02x}{g:02x}{b:02x}"
        self.tree.tag_configure("odd",            background="#a5d6a7")  # verde médio
        self.tree.tag_configure("even",           background="#ffffff")  # branco
        self.tree.tag_configure("st_sep_even",    background=c_sep,           foreground="#1e3a8a")
        self.tree.tag_configure("st_sep_odd",     background=darken(c_sep),   foreground="#1e3a8a")
        self.tree.tag_configure("st_cancel_even", background=c_can,           foreground="#7f1d1d")
        self.tree.tag_configure("st_cancel_odd",  background=darken(c_can),   foreground="#7f1d1d")

    # ── escala UI ──
    def _apply_ui_scale(self):
        scale_key = APP_CONFIG.get("ui_scale", "100")
        if scale_key not in UI_SCALE_OPTIONS:
            scale_key = "100"
        spec = UI_SCALE_OPTIONS[scale_key]
        df = tkfont.nametofont("TkDefaultFont")
        tf = df.copy(); tf.configure(size=spec["font"])
        hf = df.copy(); hf.configure(size=spec["heading_font"], weight="bold")
        self.tree_font = tkfont.Font(family="Helvetica", size=spec["font"])
        self.tree_heading_font = tkfont.Font(family="Helvetica", size=spec["heading_font"], weight="bold")
        self.option_add("*Font", tf)
        self.style.configure("Treeview", rowheight=spec["rowheight"])
        self.style.configure("Treeview.Heading")
        for s in ("TLabel", "TButton", "TCheckbutton", "TEntry"):
            self.style.configure(s, font=tf)
        if getattr(self, "tree", None) is not None:
            self.tree.configure(style="Treeview")
            for col in self.tree["columns"]:
                self.tree.heading(col, text=self.tree.heading(col, "text"))
                scale_ratio = spec["font"] / 10.0
                width = int(TREE_COLUMN_WIDTHS.get(col, 100) * min(scale_ratio, 1.8))
                self.tree.column(col, width=width)
            self.tk.call("ttk::style", "configure", "Treeview", "-rowheight", spec["rowheight"])
            self.tk.call("ttk::style", "configure", "Treeview", "-font", str(self.tree_font))
            self.tk.call("ttk::style", "configure", "Treeview.Heading", "-font", str(self.tree_heading_font))

    # ── busca ──
    def clear_search(self): self.search_var.set("")
    def _on_search_change(self, *_): self._apply_filter()

    def _focus_search(self):
        self.search_entry.focus_set()
        self.search_entry.select_range(0, "end")
        return "break"

    def _apply_filter(self):
        f = self.search_var.get().strip().lower()
        status_sel = (self.status_filter_var.get() or "(todos)").strip()
        pedidos = list(self._all_pedidos)
        if status_sel and status_sel != "(todos)":
            pedidos = [p for p in pedidos
                       if (p.get("status") or "").strip().upper() == status_sel.upper()]
        self._search_matched_cracha = False
        if f:
            pedidos = [
                p for p in pedidos
                if any(
                    f in (p.get(field) or "").lower()
                    for field in ("operador", "codigo", "num_ifood", "cliente", "_cracha_busca")
                )
            ]
            # busca por crachá: se algum resultado casou pelo crachá,
            # os NÃO impressos (IMP sem check) aparecem primeiro
            self._search_matched_cracha = any(
                f in (p.get("_cracha_busca") or "").lower() for p in pedidos
            )
        self._filtered_pedidos = pedidos
        self._render_rows()

    def _update_status_filter_options(self):
        """Atualiza as opções do filtro de status com os valores presentes na planilha."""
        seen = []
        for p in self._all_pedidos:
            s = (p.get("status") or "").strip().upper()
            if s and s not in seen:
                seen.append(s)
        values = ["(todos)"] + sorted(seen)
        self.status_filter_cb.configure(values=values)
        if self.status_filter_var.get() not in values:
            self.status_filter_var.set("(todos)")

    # ── Ordenação por coluna ──
    def _on_heading_click(self, col):
        """Cicla: ascendente → descendente → ordem padrão (STATUS + INÍCIO BANDA)."""
        st = self._sort_state
        if st["col"] != col:
            st["col"], st["reverse"] = col, False
        elif not st["reverse"]:
            st["reverse"] = True
        else:
            st["col"], st["reverse"] = None, False
        self._update_heading_arrows()
        self._render_rows()

    def _update_heading_arrows(self):
        st = self._sort_state
        for c, base in self._tree_headers.items():
            arrow = ""
            if st["col"] == c:
                arrow = "  ▼" if st["reverse"] else "  ▲"
            self.tree.heading(c, text=base + arrow)

    def _sorted_for_display(self):
        st = self._sort_state
        pedidos = list(self._filtered_pedidos)
        if not st["col"]:
            if getattr(self, "_search_matched_cracha", False):
                printed = getattr(self, "_printed_codes", set())
                # ordenação estável: não impressos primeiro, mantendo a ordem
                # padrão (STATUS + INÍCIO BANDA) dentro de cada grupo
                pedidos.sort(key=lambda p: (p.get("codigo") or "").strip() in printed)
            return pedidos
        col = st["col"]
        if col in ("inicio_banda", "fim_banda"):
            def keyf(p):
                t, _txt = _parse_horario(p.get(col) or "")
                return (1, "00:00", (p.get(col) or "").lower()) if t is None \
                    else (0, t.strftime("%H:%M:%S"), "")
        else:
            def keyf(p):
                v = (p.get(col) or "").strip()
                try:
                    return (0, float(v.replace(",", ".")), v.lower())
                except ValueError:
                    return (1, 0.0, v.lower())
        return sorted(pedidos, key=keyf, reverse=st["reverse"])

    def _render_rows(self):
        sel_codigos = {
            self._row_codigo(iid)
            for iid in self.tree.selection()
        }
        sel_codigos.discard(None)
        self.tree.delete(*self.tree.get_children())
        zebra = APP_CONFIG.get("zebra_rows", "1") == "1"
        use_status_colors = APP_CONFIG.get("status_colors", "1") == "1"
        printed = getattr(self, "_printed_codes", set())
        for p in self._filtered_pedidos:
            p["impresso"] = "✔" if (p.get("codigo") or "").strip() in printed else ""
        display = self._sorted_for_display()
        self._display_pedidos = display
        for idx, p in enumerate(display):
            s = (p["status"] or "").upper()
            is_odd = idx % 2 == 0
            is_cancelled = "CANCEL" in s
            is_separation = "SEPARATION" in s
            if use_status_colors and is_cancelled:
                # Vermelho com zebra
                tags = ("st_cancel_odd",) if is_odd else ("st_cancel_even",)
            elif use_status_colors and is_separation:
                # Azul com zebra
                tags = ("st_sep_odd",) if is_odd else ("st_sep_even",)
            elif zebra:
                # Todos os outros: só zebra verde/branco
                tags = ("odd",) if is_odd else ("even",)
            else:
                tags = ()
            self.tree.insert("", "end", iid=str(idx),
                values=(p.get("impresso", ""), p["codigo"], p["num_ifood"], p["cliente"], p["coleta"], p["devolucao"], p["tipo"], p["operador"],
                        p["status"], p["inicio_banda"], p["fim_banda"]), tags=tags)
        if sel_codigos:
            restore = [iid for iid in self.tree.get_children()
                       if self._row_codigo(iid) in sel_codigos]
            if restore:
                self.tree.selection_set(restore)
                self.tree.focus(restore[0])
        total = len(self._all_pedidos)
        exib  = len(self._filtered_pedidos)
        has_filter = bool(self.search_var.get().strip()) or \
            (self.status_filter_var.get() or "(todos)") != "(todos)"
        self.status_lbl.config(text=(
            f"{exib} de {total} pedidos" if has_filter
            else f"{total} pedidos (ordenados por STATUS e INÍCIO BANDA)"
        ) if total else "0 pedidos")

    def reload(self):
        if self._is_reloading: return
        self._is_reloading = True
        self.status_lbl.config(text="Atualizando…")
        try:
            values = obter_dados_google_sheets(APP_CONFIG["spreadsheet_id"], APP_CONFIG["range_name"])
            self._all_pedidos = parse_pedidos(values)
            apoio_map = obter_mapa_apoio_busca(APP_CONFIG["spreadsheet_id"])
            apoio_diag = aplicar_apoio_busca_aos_pedidos(self._all_pedidos, apoio_map)
            _safe_log(
                "[APOIO BUSCA] matches=%s crachas=%s operadores_preenchidos=%s" % (
                    apoio_diag["matches"],
                    apoio_diag["crachas"],
                    apoio_diag["operadores_preenchidos"],
                )
            )
            self._update_status_filter_options()
            hist = (ler_codigos_ja_impressos(APP_CONFIG["spreadsheet_id"])
                    if APP_CONFIG.get("history_enabled", "1") == "1" else set())
            self._printed_codes = hist | getattr(self, "_session_printed", set())
            if self._panel_active():
                self._history_volumes = ler_historico_volumes(APP_CONFIG["spreadsheet_id"])
        except Exception as e:
            _log_error(f"Erro ao ler planilha: {e}")
            messagebox.showerror("Erro ao ler planilha", str(e))
            self.status_lbl.config(text="Falha ao atualizar")
            self._is_reloading = False
            return
        self._apply_filter()
        self._refresh_panels()
        self._report_update_if_pending()
        self._is_reloading = False

    def _sample_preview_data(self):
        pedido = self._get_selected_pedido()
        if pedido:
            return {
                "codigo": pedido["codigo"],
                "num_ifood": pedido["num_ifood"] or "3375",
                "cliente": pedido["cliente"] or "Geisiane Souza",
                "coleta": pedido["coleta"] or "0321",
                "volume": "1/3",
            }
        return {
            "codigo": "1780432803_4777726_N",
            "num_ifood": "3375",
            "cliente": "Geisiane Souza",
            "coleta": "0321",
            "volume": "1/3",
        }

    def _print_extra_enabled(self, config=None):
        source = APP_CONFIG if config is None else config
        return source.get("print_extra_coleta", "1") == "1"

    def _render_label_preview(self, canvas, config, highlight_key=None):
        data = self._sample_preview_data()
        canvas.delete("all")
        canvas.create_rectangle(0, 0, 760, 520, fill="#ececec", outline="")
        canvas.preview_images = []
        label_w = 580
        label_h = int(label_w * (240 / 800))
        left = 20
        scale = label_w / 800.0

        def px(value):
            return int(round(value * scale))

        def draw_highlight(top_y, x1, y1, x2, y2):
            canvas.create_rectangle(
                left + px(x1), top_y + px(y1), left + px(x2), top_y + px(y2),
                outline="#f59e0b", width=3, dash=(6, 4), fill="#f59e0b", stipple="gray25"
            )

        def draw_preview_block(top_y, title, extra_qr=False):
            canvas.create_text(left, top_y - 18, text=title, anchor="w", fill="#4b5563", font=("Helvetica", 11, "bold"))
            canvas.create_rectangle(left, top_y, left + label_w, top_y + label_h, fill="#ffffff", outline="#cbd5e1", width=2)
            if extra_qr:
                epl = _build_epl_extra_for_config(
                    config,
                    codigo=data["codigo"],
                    num_ifood=data["num_ifood"],
                    cliente=data["cliente"],
                    coleta=data["coleta"],
                )
            else:
                vol_atual, vol_total = [int(x) for x in data["volume"].split("/", 1)]
                epl = _build_epl_volume_for_config(
                    config,
                    codigo=data["codigo"],
                    num_ifood=data["num_ifood"],
                    cliente=data["cliente"],
                    coleta=data["coleta"],
                    vol_atual=vol_atual,
                    vol_total=vol_total,
                )
            preview_img = _render_epl_preview_image(epl, label_w)
            canvas.preview_images.append(preview_img)
            canvas.create_image(left, top_y, image=preview_img, anchor="nw")

            if highlight_key in {"font_labels"}:
                draw_highlight(top_y, 190, 8, 625, 175)
            elif highlight_key in {"font_cliente", "mult_cliente", "bold_cliente"} and extra_qr:
                draw_highlight(top_y, 15, 100, 600, 170)
            elif highlight_key in {"font_pedido_num", "mult_pedido_num", "bold_pedido_num"}:
                draw_highlight(top_y, 300, 8, 470, 58)
            elif highlight_key in {"font_coleta_val", "bold_coleta"}:
                draw_highlight(top_y, 330, 140, 470, 182)
                if extra_qr:
                    draw_highlight(top_y, 615, 15, 760, 140)
            elif highlight_key in {"font_volume", "mult_volume"} and not extra_qr:
                draw_highlight(top_y, 520, 15, 760, 95)

        draw_preview_block(28, "Etiqueta volume")
        if self._print_extra_enabled(config):
            draw_preview_block(28 + label_h + 58, "Etiqueta extra coleta", extra_qr=True)
        else:
            top_y = 28 + label_h + 58
            canvas.create_text(left, top_y - 18, text="Etiqueta extra coleta", anchor="w", fill="#4b5563", font=("Helvetica", 11, "bold"))
            canvas.create_rectangle(left, top_y, left + label_w, top_y + label_h, fill="#f8fafc", outline="#cbd5e1", width=2)
            canvas.create_text(left + label_w / 2, top_y + label_h / 2,
                               text="Impressão da etiqueta extra desativada",
                               anchor="center", fill="#64748b", font=("Helvetica", 13, "bold"))

    def _confirm_large_print(self, total_labels, detail_text):
        result = {"confirmed": False}
        win = tk.Toplevel(self)
        win.title("Confirmar impressão")
        win.transient(self)
        win.grab_set()
        win.resizable(False, False)

        body = ttk.Frame(win, padding=18)
        body.pack(fill=BOTH, expand=YES)
        ttk.Label(body, text="Tem certeza que deseja imprimir", anchor="center").pack()
        ttk.Label(body, text=f"{total_labels} etiquetas", anchor="center",
                  font=("Helvetica", 16, "bold")).pack(pady=(6, 4))
        ttk.Label(body, text=detail_text, anchor="center", foreground="#6b7280").pack()

        buttons = ttk.Frame(body, padding=(0, 12, 0, 0))
        buttons.pack()

        def accept():
            result["confirmed"] = True
            win.destroy()

        def reject():
            win.destroy()

        ttk.Button(buttons, text="Sim", command=accept).pack(side=LEFT, padx=(0, 8))
        ttk.Button(buttons, text="Não", command=reject).pack(side=LEFT)
        win.protocol("WM_DELETE_WINDOW", reject)
        self.wait_window(win)
        return result["confirmed"]

    def run_self_update(self):
        if not messagebox.askyesno(
            "Atualizar sistema",
            "O sistema vai consultar o manifest Rotom, baixar o ZIP novo, renomear reboot_QR para reboot_QR_OLDN e instalar a nova versão como reboot_QR.\n\nDeseja continuar?",
            parent=self,
        ):
            return

        def worker():
            try:
                self.after(0, lambda: self.status_lbl.config(text="Consultando manifest da atualização…"))
                payload = _download_update_payload()
                zip_url = str(payload["zip_url"]).strip()
                expected_sha = str(payload.get("sha256", "")).strip().lower()
                version = str(payload.get("version", "")).strip()

                if version and version == APP_VERSION:
                    self.after(0, lambda: messagebox.showinfo("Atualização", f"Você já está na versão {APP_VERSION}.", parent=self))
                    self.after(0, lambda: self.status_lbl.config(text=f"Versão {APP_VERSION} já instalada"))
                    return

                staging_dir = tempfile.mkdtemp(prefix="rebootqr_update_")
                zip_path = os.path.join(staging_dir, "update.zip")
                extract_dir = os.path.join(staging_dir, "unzipped")
                os.makedirs(extract_dir, exist_ok=True)
                self.after(0, lambda: self.status_lbl.config(text="Baixando ZIP da atualização…"))
                _download_file(zip_url, zip_path)

                if expected_sha:
                    actual_sha = _sha256_file(zip_path).lower()
                    if actual_sha != expected_sha:
                        raise RuntimeError("Hash SHA256 do ZIP não confere com o manifesto.")

                with zipfile.ZipFile(zip_path, "r") as zf:
                    zf.extractall(extract_dir)

                extracted_root = _prepare_extracted_root(staging_dir, extract_dir)
                script_path = os.path.join(staging_dir, "apply_update.zsh")
                with open(script_path, "w", encoding="utf-8") as fh:
                    fh.write(_build_updater_script(staging_dir, extracted_root))
                os.chmod(script_path, 0o755)

                subprocess.Popen(["/bin/zsh", script_path], cwd=INSTALL_BASE_DIR)
                self.after(0, lambda: self.status_lbl.config(text="Atualização aplicada. Reiniciando em reboot_QR…"))
                self.after(0, self.destroy)
            except Exception as exc:
                self.after(0, lambda: self.status_lbl.config(text="Falha na atualização"))
                self.after(0, lambda: messagebox.showerror("Erro na atualização", str(exc), parent=self))

        threading.Thread(target=worker, daemon=True).start()

    # ── Configurações ──
    def open_settings(self):
        win = tk.Toplevel(self)
        win.title("Configurações")
        win.transient(self)
        win.grab_set()
        win.resizable(True, True)
        _orig_border_th = APP_CONFIG.get("border_thickness", "media")
        def _cancel_settings():
            # desfaz a espessura aplicada temporariamente para o preview
            APP_CONFIG["border_thickness"] = _orig_border_th
            win.destroy()
        win.protocol("WM_DELETE_WINDOW", _cancel_settings)
        screen_w = win.winfo_screenwidth()
        screen_h = win.winfo_screenheight()
        win_w = min(1260, max(900, screen_w - 80))
        win_h = min(860, max(620, screen_h - 120))
        win.geometry(f"{win_w}x{win_h}")
        win.minsize(760, 600)  # garante que botões sempre aparecem
        win.update_idletasks()
        # Botões ANTES do notebook no pack order (side=BOTTOM garante visibilidade)
        actions = ttk.Frame(win, padding=(10, 6, 10, 10))
        actions.pack(side="bottom", fill=X)
        ttk.Separator(win, orient="horizontal").pack(side="bottom", fill=X)

        notebook = ttk.Notebook(win)
        notebook.pack(fill=BOTH, expand=YES, padx=10, pady=(10, 0))

        # --- Suporte a rolagem nas abas (conteúdo mais alto que a janela) ---
        _scroll_targets = []

        def make_scrollable(container):
            canvas = tk.Canvas(container, highlightthickness=0)
            vsb = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
            canvas.configure(yscrollcommand=vsb.set)
            vsb.pack(side="right", fill="y")
            canvas.pack(side="left", fill="both", expand=True)
            inner = ttk.Frame(canvas, padding=12)
            win_id = canvas.create_window((0, 0), window=inner, anchor="nw")
            inner.bind("<Configure>", lambda _e: canvas.configure(scrollregion=canvas.bbox("all")))
            canvas.bind("<Configure>", lambda e: canvas.itemconfigure(win_id, width=e.width))
            _scroll_targets.append((canvas, inner))
            return inner

        def _bind_wheel_recursive(widget, canvas):
            def _on(e):
                d = 0
                if getattr(e, "delta", 0):
                    d = -1 if e.delta > 0 else 1
                elif getattr(e, "num", 0) == 4:
                    d = -1
                elif getattr(e, "num", 0) == 5:
                    d = 1
                canvas.yview_scroll(d, "units")
                return "break"
            for seq in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
                widget.bind(seq, _on)
            for ch in widget.winfo_children():
                _bind_wheel_recursive(ch, canvas)

        # ════════════════════════════════════════
        # ABA 1 — Planilha
        # ════════════════════════════════════════
        tab_sheet_outer = ttk.Frame(notebook)
        notebook.add(tab_sheet_outer, text="  Planilha  ")
        tab_sheet = make_scrollable(tab_sheet_outer)

        sheet_fields = [
            ("Nome do Raspberry",     "raspberry_name"),
            ("Manifest Rotom",        "rotom_manifest_url"),
            ("Loja / Planilha",       "spreadsheet_id"),
            ("Range",                 "range_name"),
            ("Tamanho da interface",  "ui_scale"),
            ("Etiqueta extra coleta", "print_extra_coleta"),
        ("Modo de impressão", "print_mode"),
            ("Coluna código shopper", "header_codigo_shopper"),
            ("Coluna pedido iFood",   "header_num_ifood"),
            ("Coluna status",         "header_status"),
            ("Coluna cliente",        "header_cliente"),
            ("Coluna coleta",         "header_coleta"),
            ("Coluna devolução",      "header_devolucao"),
            ("Coluna tipo",           "header_tipo"),
            ("Coluna operador",       "header_operador"),
            ("Coluna início banda",   "header_inicio_banda"),
            ("Coluna fim banda",      "header_fim_banda"),
        ("Linhas alternadas (zebra)", "zebra_rows"),
        ("Cores por status", "status_colors"),
        ("Impressão em lote", "batch_ask_volumes"),
        ("Lote: volumes padrão", "batch_default_volumes"),
        ("Limpar busca ao imprimir", "clear_search_after_print"),
        ("Painel: atualizar a cada", "panel_refresh_seconds"),
        ("Painel: cronômetro laranja", "panel_timer_warn_min"),
        ("Painel: cronômetro vermelho", "panel_timer_alert_min"),
        ("Painel: 'preparando'", "panel_preparing_statuses"),
        ("Painel: 'separação ok'", "panel_sep_done_statuses"),
        ("Painel: 'a caminho'", "panel_going_statuses"),
        ("Painel: 'na loja'", "panel_ready_statuses"),
        ("Histórico de impressão", "history_enabled"),
        ("Modo desenvolvedor", "dev_mode"),
        ]
        sheet_vars = {}
        use_two_columns = win_h < 760 or win_w < 1100
        for row, (label, key) in enumerate(sheet_fields):
            if use_two_columns:
                col_group = row % 2
                row_pos = row // 2
                label_col = col_group * 2
                value_col = label_col + 1
                label_pad = (0, 8) if col_group == 0 else (24, 8)
            else:
                row_pos = row
                label_col = 0
                value_col = 1
                label_pad = (0, 8)

            ttk.Label(tab_sheet, text=label + ":").grid(row=row_pos, column=label_col, sticky="w", padx=label_pad, pady=4)
            if key == "spreadsheet_id":
                # nome exclusivo: o loop reusa 'var' nos campos seguintes e os
                # closures abaixo precisam apontar SEMPRE para este StringVar
                sid_var = StringVar(value=APP_CONFIG.get("spreadsheet_id", ""))
                sheet_vars[key] = sid_var
                sfrm = ttk.Frame(tab_sheet)
                sfrm.grid(row=row_pos, column=value_col, sticky="we", pady=4)

                def _loja_do_id():
                    sid = sid_var.get().strip()
                    for nm, s in DARK_STORES.items():
                        if s == sid:
                            return nm
                    return DARK_OUTRA

                loja_var = StringVar(value=_loja_do_id())
                loja_cb = ttk.Combobox(sfrm, textvariable=loja_var,
                                       values=list(DARK_STORES.keys()) + [DARK_OUTRA],
                                       state="readonly", width=17)
                loja_cb.pack(side=LEFT)
                id_entry = ttk.Entry(sfrm, textvariable=sid_var, width=42)
                id_entry.pack(side=LEFT, padx=(8, 0), fill=X, expand=True)

                def _on_loja_settings(_e=None):
                    nm = loja_var.get()
                    if nm == DARK_OUTRA:
                        id_entry.configure(state="normal")
                        id_entry.focus_set()
                        id_entry.select_range(0, "end")
                    else:
                        sid_var.set(DARK_STORES[nm])
                        id_entry.configure(state="readonly")

                loja_cb.bind("<<ComboboxSelected>>", _on_loja_settings)
                if loja_var.get() != DARK_OUTRA:
                    id_entry.configure(state="readonly")
                # usado pelo "Restaurar padrões" para ressincronizar o combobox
                win._loja_resync = lambda: (loja_var.set(_loja_do_id()), _on_loja_settings())
            elif key == "ui_scale":
                labels = [f'{c["label"]} ({s}%)' for s, c in UI_SCALE_OPTIONS.items()]
                rmap   = {f'{c["label"]} ({s}%)': s for s, c in UI_SCALE_OPTIONS.items()}
                cur    = APP_CONFIG.get("ui_scale", "100")
                var    = StringVar(value=f'{UI_SCALE_OPTIONS[cur]["label"]} ({cur}%)')
                sheet_vars[key] = (var, rmap)
                ttk.Combobox(tab_sheet, textvariable=var, values=labels, state="readonly",
                             width=26).grid(row=row_pos, column=value_col, sticky="we", pady=4)
            elif key == "print_extra_coleta":
                var = BooleanVar(value=self._print_extra_enabled())
                sheet_vars[key] = var
                ttk.Checkbutton(tab_sheet, text="Imprimir etiqueta extra de coleta", variable=var).grid(row=row_pos, column=value_col, sticky="w", pady=4)
            elif key == "print_mode":
                PRINT_MODES = {
                    "padrao": "Padrão da impressora",
                    "direta": "Térmica direta (sem ribbon)",
                    "transferencia": "Transferência térmica (com ribbon)",
                }
                atual = APP_CONFIG.get("print_mode", "padrao")
                var = StringVar(value=PRINT_MODES.get(atual, PRINT_MODES["padrao"]))
                sheet_vars[key] = (var, PRINT_MODES)
                ttk.Combobox(tab_sheet, textvariable=var,
                             values=list(PRINT_MODES.values()),
                             state="readonly", width=32
                             ).grid(row=row_pos, column=value_col, sticky="w", pady=4)
            elif key == "panel_refresh_seconds":
                var = StringVar(value=APP_CONFIG.get("panel_refresh_seconds", "30"))
                sheet_vars[key] = var
                pfrm = ttk.Frame(tab_sheet)
                pfrm.grid(row=row_pos, column=value_col, sticky="w", pady=4)
                ttk.Spinbox(pfrm, textvariable=var, from_=10, to=600,
                            increment=5, width=6).pack(side=LEFT)
                ttk.Label(pfrm, text=" segundos (só enquanto a aba Painel/telão está ativa)",
                          foreground="#888").pack(side=LEFT)
            elif key in ("panel_timer_warn_min", "panel_timer_alert_min"):
                var = StringVar(value=APP_CONFIG.get(key, DEFAULT_APP_CONFIG[key]))
                sheet_vars[key] = var
                pfrm = ttk.Frame(tab_sheet)
                pfrm.grid(row=row_pos, column=value_col, sticky="w", pady=4)
                ttk.Spinbox(pfrm, textvariable=var, from_=1, to=120,
                            increment=1, width=6).pack(side=LEFT)
                ttk.Label(pfrm, text=" minutos na coluna PREPARANDO",
                          foreground="#888").pack(side=LEFT)
            elif key in ("panel_preparing_statuses", "panel_sep_done_statuses",
                         "panel_going_statuses", "panel_ready_statuses"):
                var = StringVar(value=APP_CONFIG.get(key, DEFAULT_APP_CONFIG[key]))
                sheet_vars[key] = var
                pfrm = ttk.Frame(tab_sheet)
                pfrm.grid(row=row_pos, column=value_col, sticky="we", pady=4)
                ttk.Entry(pfrm, textvariable=var, width=34).pack(side=LEFT)
                ttk.Label(pfrm, text=" status separados por vírgula (ignora _/maiúsculas)",
                          foreground="#888").pack(side=LEFT)
            elif key == "clear_search_after_print":
                var = BooleanVar(value=APP_CONFIG.get("clear_search_after_print", "1") == "1")
                sheet_vars[key] = var
                ttk.Checkbutton(tab_sheet,
                                text="Limpar o campo de busca automaticamente após cada impressão",
                                variable=var).grid(row=row_pos, column=value_col, sticky="w", pady=4)
            elif key == "history_enabled":
                var = BooleanVar(value=APP_CONFIG.get("history_enabled", "1") == "1")
                sheet_vars[key] = var
                ttk.Checkbutton(tab_sheet,
                                text="Registrar impressões na aba 'Histórico de Impressão' (requer Editor na planilha)",
                                variable=var).grid(row=row_pos, column=value_col, sticky="w", pady=4)
            elif key == "dev_mode":
                var = BooleanVar(value=APP_CONFIG.get("dev_mode", "0") == "1")
                sheet_vars[key] = var
                ttk.Checkbutton(tab_sheet,
                                text="Exibir botão 'Log de erros' na barra de status",
                                variable=var).grid(row=row_pos, column=value_col, sticky="w", pady=4)
            elif key == "batch_ask_volumes":
                var = BooleanVar(value=APP_CONFIG.get("batch_ask_volumes", "1") == "1")
                sheet_vars[key] = var
                ttk.Checkbutton(tab_sheet, text="Perguntar volumes de cada pedido no lote",
                                variable=var).grid(row=row_pos, column=value_col, sticky="w", pady=4)
            elif key == "batch_default_volumes":
                var = StringVar(value=APP_CONFIG.get("batch_default_volumes", "1"))
                sheet_vars[key] = var
                batch_frame = ttk.Frame(tab_sheet)
                batch_frame.grid(row=row_pos, column=value_col, sticky="w", pady=4)
                ttk.Spinbox(batch_frame, textvariable=var, from_=1, to=99,
                            width=5).pack(side=LEFT)
                ttk.Label(batch_frame, text=" volumes/pedido quando a pergunta está desativada",
                          foreground="#888").pack(side=LEFT)
            elif key == "zebra_rows":
                var = BooleanVar(value=APP_CONFIG.get("zebra_rows","1") == "1")
                sheet_vars[key] = var
                ttk.Checkbutton(tab_sheet, text="Ativar cores alternadas nas linhas", variable=var).grid(row=row_pos, column=value_col, sticky="w", pady=4)
            elif key == "status_colors":
                var = BooleanVar(value=APP_CONFIG.get("status_colors","1") == "1")
                sheet_vars[key] = var
                color_frame = ttk.Frame(tab_sheet)
                color_frame.grid(row=row_pos, column=value_col, sticky="w", pady=4)
                ttk.Checkbutton(color_frame, text="Ativar cores por status", variable=var).pack(side=LEFT)
                # Cor SEPARATION
                sep_color_var = StringVar(value=APP_CONFIG.get("color_separation","#bfdbfe"))
                sheet_vars["color_separation"] = sep_color_var
                ttk.Label(color_frame, text="  SEPARATION:").pack(side=LEFT)
                sep_swatch = tk.Label(color_frame, bg=sep_color_var.get(), width=3, relief="solid")
                sep_swatch.pack(side=LEFT, padx=(2,0))
                def pick_sep_color(sv=sep_color_var, sw=sep_swatch):
                    from tkinter import colorchooser
                    c = colorchooser.askcolor(color=sv.get(), title="Cor SEPARATION")[1]
                    if c: sv.set(c); sw.config(bg=c)
                ttk.Button(color_frame, text="...", width=3, command=pick_sep_color).pack(side=LEFT, padx=(2,4))
                # Cor CANCELLED
                can_color_var = StringVar(value=APP_CONFIG.get("color_cancelled","#fca5a5"))
                sheet_vars["color_cancelled"] = can_color_var
                ttk.Label(color_frame, text="CANCELLED:").pack(side=LEFT)
                can_swatch = tk.Label(color_frame, bg=can_color_var.get(), width=3, relief="solid")
                can_swatch.pack(side=LEFT, padx=(2,0))
                def pick_can_color(cv=can_color_var, cw=can_swatch):
                    from tkinter import colorchooser
                    c = colorchooser.askcolor(color=cv.get(), title="Cor CANCELLED")[1]
                    if c: cv.set(c); cw.config(bg=c)
                ttk.Button(color_frame, text="...", width=3, command=pick_can_color).pack(side=LEFT, padx=(2,0))
            else:
                var = StringVar(value=APP_CONFIG.get(key, ""))
                sheet_vars[key] = var
                ttk.Entry(tab_sheet, textvariable=var,
                          width=44 if key in ("spreadsheet_id", "range_name") else 30
                          ).grid(row=row_pos, column=value_col, sticky="we", pady=4)
        tab_sheet.grid_columnconfigure(1, weight=1)
        if use_two_columns:
            tab_sheet.grid_columnconfigure(3, weight=1)

        # --- Conta Google (login, permissões e diagnóstico de escrita) ---
        gfrm = ttk.LabelFrame(tab_sheet, text=" Conta Google ", padding=8)
        gfrm.grid(row=99, column=0, columnspan=4, sticky="we", pady=(14, 4))
        ttk.Label(gfrm, foreground="#666", justify="left", text=(
            "Para o Histórico de Impressão funcionar, a conta usada no login precisa ser "
            "EDITORA na planilha da loja.\n"
            "Reautenticar apaga o login salvo e abre o navegador para entrar novamente "
            "(permite trocar de conta).")).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 6))

        def _do_reauth():
            if not messagebox.askyesno(
                    "Reautenticar Google",
                    "O login salvo será apagado e o navegador abrirá para um novo login.\n"
                    "Use uma conta que seja EDITORA nas planilhas das lojas.\n\nContinuar?",
                    parent=win):
                return
            self.status_lbl.config(text="Aguardando login no navegador…")
            def run():
                try:
                    reautenticar_google()
                    self._ui_call(lambda: (
                        self.status_lbl.config(text="Reautenticado com sucesso."),
                        self._show_toast("Login Google refeito com sucesso ✓", ok=True)))
                except Exception as e:
                    _log_error(f"Reautenticação falhou: {e}")
                    self._ui_call(lambda err=_friendly_sheets_error(e): messagebox.showerror(
                        "Reautenticação falhou", err, parent=win))
            threading.Thread(target=run, daemon=True).start()

        def _do_write_test():
            sid = sheet_vars["spreadsheet_id"].get().strip() or APP_CONFIG.get("spreadsheet_id", "")
            self.status_lbl.config(text="Testando escrita na planilha…")
            def run():
                try:
                    testar_escrita_planilha(sid)
                    self._ui_call(lambda: (
                        self.status_lbl.config(text="Escrita OK."),
                        messagebox.showinfo(
                            "Teste de escrita",
                            "Escrita na planilha funcionando ✓\n"
                            f"A aba '{HISTORY_SHEET_NAME}' está pronta.", parent=win)))
                except Exception as e:
                    _log_error(f"Teste de escrita falhou: {e}")
                    self._ui_call(lambda err=_friendly_sheets_error(e): messagebox.showerror(
                        "Teste de escrita falhou", err, parent=win))
            threading.Thread(target=run, daemon=True).start()

        ttk.Button(gfrm, text="Reautenticar Google…", command=_do_reauth).grid(
            row=1, column=0, sticky="w")
        ttk.Button(gfrm, text="Testar escrita na planilha", command=_do_write_test).grid(
            row=1, column=1, sticky="w", padx=(10, 0))

        # ════════════════════════════════════════
        # ABA 2 — Layout da Etiqueta
        # ════════════════════════════════════════
        tab_lbl_outer = ttk.Frame(notebook)
        notebook.add(tab_lbl_outer, text="  Layout da Etiqueta  ")
        tab_lbl = make_scrollable(tab_lbl_outer)

        lbl_vars = {}
        preview_config = {f"lbl_{k}": _lbl(k) for k in LABEL_LAYOUT_DEFAULTS}
        hover_state = {"key": None}

        # Fonte EPL 1-5 (ou 6/7/8 para algumas Zebra)
        FONT_OPTIONS  = ["1", "2", "3", "4", "5", "6", "7", "8"]
        MULT_OPTIONS  = ["1", "2", "3", "4", "5", "6", "7", "8"]

        def add_row(parent, row, label, key, options, tooltip=""):
            lbl = ttk.Label(parent, text=label + ":")
            lbl.grid(row=row, column=0, sticky="w", padx=(0, 8), pady=5)
            var = StringVar(value=_lbl(key))
            lbl_vars[key] = var
            cb = ttk.Combobox(parent, textvariable=var, values=options, state="readonly", width=8)
            cb.grid(row=row, column=1, sticky="w", pady=5)
            var.trace_add("write", lambda *_args, name=key, value=var: update_preview_value(name, value.get()))
            bind_preview_hover(lbl, key)
            bind_preview_hover(cb, key)
            if tooltip:
                tip = ttk.Label(parent, text=tooltip, foreground="#888")
                tip.grid(row=row, column=2, sticky="w", padx=(8, 0))
                bind_preview_hover(tip, key)

        def add_check(parent, row, label, key):
            var = BooleanVar(value=_lbl(key) == "1")
            lbl_vars[key] = var
            chk = ttk.Checkbutton(parent, text=label, variable=var)
            chk.grid(row=row, column=0, columnspan=3, sticky="w", pady=5)
            var.trace_add("write", lambda *_args, name=key, value=var: update_preview_value(name, "1" if value.get() else "0"))
            bind_preview_hover(chk, key)

        # Separador visual
        def sep(parent, row, text):
            ttk.Separator(parent, orient="horizontal").grid(
                row=row, column=0, columnspan=3, sticky="ew", pady=(10, 2))
            ttk.Label(parent, text=text, font=("", 9, "bold")).grid(
                row=row+1, column=0, columnspan=3, sticky="w", pady=(0, 4))

        def update_preview_value(key, raw_value):
            preview_config[f"lbl_{key}"] = str(raw_value)
            self._render_label_preview(preview_canvas, preview_config, hover_state["key"])

        def update_sheet_preview(*_):
            if "print_extra_coleta" in sheet_vars and isinstance(sheet_vars["print_extra_coleta"], BooleanVar):
                preview_config["print_extra_coleta"] = "1" if sheet_vars["print_extra_coleta"].get() else "0"
                self._render_label_preview(preview_canvas, preview_config, hover_state["key"])

        def bind_preview_hover(widget, key):
            widget.bind("<Enter>", lambda _e, k=key: set_preview_hover(k))
            widget.bind("<Leave>", lambda _e: set_preview_hover(None))

        def set_preview_hover(key):
            hover_state["key"] = key
            self._render_label_preview(preview_canvas, preview_config, hover_state["key"])

        layout_controls = ttk.Frame(tab_lbl)
        preview_frame = ttk.LabelFrame(tab_lbl, text="Preview ao vivo", padding=10)
        if win_w < 1180:
            layout_controls.grid(row=0, column=0, sticky="nw")
            preview_frame.grid(row=1, column=0, sticky="w", pady=(18, 0))
        else:
            layout_controls.grid(row=0, column=0, sticky="nw")
            preview_frame.grid(row=0, column=1, sticky="n", padx=(20, 0))
        preview_canvas = tk.Canvas(preview_frame, width=700, height=440, highlightthickness=0, bg="#ececec")
        preview_canvas.pack()
        preview_config["print_extra_coleta"] = APP_CONFIG.get("print_extra_coleta", "1")
        if "print_extra_coleta" in sheet_vars and isinstance(sheet_vars["print_extra_coleta"], BooleanVar):
            sheet_vars["print_extra_coleta"].trace_add("write", update_sheet_preview)
        # Forçar render após o canvas estar visível
        win.after(100, lambda: self._render_label_preview(preview_canvas, preview_config, None))

        r = 0
        sep(layout_controls, r, "Número do Pedido"); r += 2
        add_row(layout_controls, r, "Tamanho da fonte",  "font_pedido_num",  FONT_OPTIONS,
                "EPL: 1=menor … 5=maior"); r += 1
        add_row(layout_controls, r, "Multiplicador",     "mult_pedido_num",  MULT_OPTIONS,
                "Escala do caractere (1x … 4x)"); r += 1
        add_check(layout_controls, r, "Negrito",         "bold_pedido_num"); r += 1

        sep(layout_controls, r, "Labels (PEDIDO:, CLIENTE:, VOLUMES:)"); r += 2
        add_row(layout_controls, r, "Tamanho da fonte",  "font_labels",      FONT_OPTIONS,
                "Rótulos fixos da etiqueta"); r += 1

        sep(layout_controls, r, "Código de Coleta"); r += 2
        add_row(layout_controls, r, "Tamanho da fonte",  "font_coleta_val",  FONT_OPTIONS,
                "Valor do código de coleta"); r += 1
        add_check(layout_controls, r, "Negrito",         "bold_coleta"); r += 1

        sep(layout_controls, r, "Volumes (X/Y)"); r += 2
        add_row(layout_controls, r, "Tamanho da fonte",  "font_volume",      FONT_OPTIONS,
                "Número de volumes na etiqueta"); r += 1
        add_row(layout_controls, r, "Multiplicador",     "mult_volume",      MULT_OPTIONS,
                "Reduz sozinho se não couber"); r += 1

        sep(layout_controls, r, "Código Shopper (etiqueta volume)"); r += 2
        add_check(layout_controls, r, "Mostrar código shopper abaixo do QR", "show_codigo_shopper"); r += 1

        sep(layout_controls, r, "Nome do Cliente (etiqueta extra)"); r += 2
        add_row(layout_controls, r, "Tamanho da fonte",  "font_cliente",     FONT_OPTIONS,
                "Nome ampliado na etiqueta extra"); r += 1
        add_row(layout_controls, r, "Multiplicador",     "mult_cliente",     MULT_OPTIONS,
                "Escala do caractere (1x … 8x)"); r += 1
        add_check(layout_controls, r, "Negrito",         "bold_cliente"); r += 1
        add_check(layout_controls, r, "Carimbo de data/hora da impressão (DD/MM HH:MM)",
                  "show_datetime"); r += 1

        sep(layout_controls, r, "Bordas rotativas (por lote de impressão)"); r += 2
        border_enabled_var = BooleanVar(value=APP_CONFIG.get("border_enabled", "0") == "1")
        ttk.Checkbutton(layout_controls, text="Ativar bordas nas etiquetas",
                        variable=border_enabled_var).grid(row=r, column=0, columnspan=3, sticky="w", pady=5); r += 1

        ttk.Label(layout_controls, text="Modo:").grid(row=r, column=0, sticky="w", padx=(0, 8), pady=5)
        _mode_disp = {"rotate": "Rotação (cíclica)", "random": "Aleatório"}
        _mode_rev  = {v: k for k, v in _mode_disp.items()}
        border_mode_var = StringVar(value=_mode_disp.get(APP_CONFIG.get("border_mode", "rotate"), "Rotação (cíclica)"))
        ttk.Combobox(layout_controls, textvariable=border_mode_var,
                     values=list(_mode_disp.values()), state="readonly", width=18
                     ).grid(row=r, column=1, columnspan=2, sticky="w", pady=5); r += 1

        ttk.Label(layout_controls, text="Espessura da moldura:").grid(row=r, column=0, sticky="w", padx=(0, 8), pady=5)
        _th_rev = {v: k for k, v in BORDER_THICKNESS_LABELS.items()}
        _th_cur = APP_CONFIG.get("border_thickness", "media")
        border_th_var = StringVar(value=BORDER_THICKNESS_LABELS.get(_th_cur, "Média"))
        _th_cb = ttk.Combobox(layout_controls, textvariable=border_th_var,
                              values=list(BORDER_THICKNESS_LABELS.values()), state="readonly", width=10)
        _th_cb.grid(row=r, column=1, columnspan=2, sticky="w", pady=5); r += 1
        def _on_th_change(*_):
            # aplica temporariamente para o preview refletir a espessura escolhida
            APP_CONFIG["border_thickness"] = _th_rev.get(border_th_var.get(), "media")
            self._render_label_preview(preview_canvas, preview_config, hover_state["key"])
        border_th_var.trace_add("write", _on_th_change)

        ttk.Label(layout_controls, text="Modelos habilitados na rotação:").grid(
            row=r, column=0, columnspan=3, sticky="w", pady=(8, 2)); r += 1
        _enabled_now = set((APP_CONFIG.get("border_models", "") or ",".join(BORDER_PATTERNS)).split(","))
        border_model_vars = {}
        models_frame = ttk.Frame(layout_controls)
        models_frame.grid(row=r, column=0, columnspan=3, sticky="w"); r += 1
        for i, bp in enumerate(BORDER_PATTERNS):
            v = BooleanVar(value=bp in _enabled_now)
            border_model_vars[bp] = v
            ttk.Checkbutton(models_frame, text=BORDER_PATTERN_LABELS[bp], variable=v).grid(
                row=i // 2, column=i % 2, sticky="w", padx=(0, 18), pady=2)

        ttk.Label(layout_controls, text="Visualizar padrão:").grid(row=r, column=0, sticky="w", padx=(0, 8), pady=5)
        _bprev_opts = {"(sem borda)": None}
        for _bp in BORDER_PATTERNS:
            _bprev_opts[BORDER_PATTERN_LABELS[_bp]] = _bp
        _bprev_default = BORDER_PATTERN_LABELS["preta"] if border_enabled_var.get() else "(sem borda)"
        preview_border_var = StringVar(value=_bprev_default)
        preview_config["_border_preview"] = _bprev_opts[_bprev_default]
        def _on_border_preview(*_):
            preview_config["_border_preview"] = _bprev_opts.get(preview_border_var.get())
            self._render_label_preview(preview_canvas, preview_config, hover_state["key"])
        preview_border_var.trace_add("write", _on_border_preview)
        ttk.Combobox(layout_controls, textvariable=preview_border_var,
                     values=list(_bprev_opts.keys()), state="readonly", width=28
                     ).grid(row=r, column=1, columnspan=2, sticky="w", pady=5); r += 1
        ttk.Label(layout_controls,
                  text="A cada impressão o lote inteiro sai com 1 padrão; o próximo lote avança na sequência.",
                  foreground="#888", wraplength=380, justify="left"
                  ).grid(row=r, column=0, columnspan=3, sticky="w", pady=(0, 4)); r += 1
        ttk.Button(layout_controls, text="Imprimir teste de todos os modelos",
                   command=self.print_border_test
                   ).grid(row=r, column=0, columnspan=3, sticky="w", pady=(2, 6)); r += 1

        tab_lbl.grid_columnconfigure(1, weight=1)
        self._render_label_preview(preview_canvas, preview_config, hover_state["key"])

        # ════════════════════════════════════════
        # Botões de ação — já criados no topo (side=bottom)
        # ════════════════════════════════════════
        ttk.Button(actions, text="Atualizar sistema", command=self.run_self_update).pack(side=LEFT, padx=(0, 8))

        def save_and_reload():
            updated = dict(APP_CONFIG)

            # --- Planilha ---
            for key, wd in sheet_vars.items():
                if key == "print_mode":
                    var, modes = wd
                    inverso = {v: k for k, v in modes.items()}
                    updated[key] = inverso.get(var.get(), "padrao")
                    continue
                if key == "ui_scale":
                    var, rmap = wd
                    v = rmap.get(var.get().strip())
                    if not v:
                        messagebox.showwarning("Campo obrigatório", "Preencha: ui_scale", parent=win)
                        return
                    updated[key] = v
                else:
                    if isinstance(wd, BooleanVar):
                        updated[key] = "1" if wd.get() else "0"
                    else:
                        v = wd.get().strip()
                        if key in ("color_separation", "color_cancelled"):
                            updated[key] = v if v else DEFAULT_APP_CONFIG.get(key, "")
                        elif key in ("panel_timer_warn_min", "panel_timer_alert_min"):
                            if not v.isdigit() or not (1 <= int(v) <= 120):
                                messagebox.showwarning("Valor inválido",
                                    "Painel: os limites do cronômetro devem ser números entre 1 e 120 minutos.",
                                    parent=win)
                                return
                            updated[key] = v
                        elif key == "panel_refresh_seconds":
                            if not v.isdigit() or not (10 <= int(v) <= 600):
                                messagebox.showwarning("Valor inválido",
                                    "Painel: o intervalo deve ser um número entre 10 e 600 segundos.",
                                    parent=win)
                                return
                            updated[key] = v
                        elif key == "batch_default_volumes":
                            if not v.isdigit() or not (1 <= int(v) <= 99):
                                messagebox.showwarning("Valor inválido",
                                    "Lote: volumes padrão deve ser um número entre 1 e 99.", parent=win)
                                return
                            updated[key] = v
                        elif not v:
                            messagebox.showwarning("Campo obrigatório", f"Preencha: {key}", parent=win)
                            return
                        else:
                            updated[key] = v

            # --- Layout ---
            for key, var in lbl_vars.items():
                if isinstance(var, BooleanVar):
                    updated[f"lbl_{key}"] = "1" if var.get() else "0"
                else:
                    updated[f"lbl_{key}"] = var.get()

            # --- Bordas rotativas ---
            updated["border_enabled"] = "1" if border_enabled_var.get() else "0"
            updated["border_mode"] = _mode_rev.get(border_mode_var.get(), "rotate")
            updated["border_thickness"] = _th_rev.get(border_th_var.get(), "media")
            enabled_models = [bp for bp in BORDER_PATTERNS if border_model_vars[bp].get()]
            if not enabled_models:
                messagebox.showwarning("Modelos de etiqueta",
                    "Habilite pelo menos um modelo de borda para a rotação.", parent=win)
                return
            updated["border_models"] = ",".join(enabled_models)
            # se a lista mudou, reinicia o contador para não pular modelos
            if updated["border_models"] != APP_CONFIG.get("border_models", ""):
                updated["border_index"] = "0"

            # coerência dos limites do cronômetro
            try:
                if int(updated.get("panel_timer_warn_min", "3")) >= int(updated.get("panel_timer_alert_min", "5")):
                    messagebox.showwarning("Valor inválido",
                        "Painel: o limite laranja deve ser MENOR que o vermelho.", parent=win)
                    return
            except Exception:
                pass

            try:
                save_app_config(updated)
            except Exception as e:
                messagebox.showerror("Erro ao salvar", str(e), parent=win)
                return

            APP_CONFIG.clear()
            APP_CONFIG.update(updated)
            self._apply_ui_scale()
            self._apply_tree_colors()
            self._refresh_devlog_button()
            self.dark_var.set(self._initial_dark_name())
            win.destroy()
            self.reload()
            self.status_lbl.config(text="Configuração salva.")

        def reset_defaults():
            # Planilha
            for key, wd in sheet_vars.items():
                if key == "print_mode":
                    var, modes = wd
                    var.set(modes.get(DEFAULT_APP_CONFIG["print_mode"], "Padrão da impressora"))
                    continue
                if key == "ui_scale":
                    var, _ = wd
                    s = DEFAULT_APP_CONFIG[key]
                    var.set(f'{UI_SCALE_OPTIONS[s]["label"]} ({s}%)')
                else:
                    wd.set(DEFAULT_APP_CONFIG[key])
            if hasattr(win, "_loja_resync"):
                win._loja_resync()
            # Layout
            for key, var in lbl_vars.items():
                default = LABEL_LAYOUT_DEFAULTS[key]
                if isinstance(var, BooleanVar):
                    var.set(default == "1")
                else:
                    var.set(default)
                preview_config[f"lbl_{key}"] = default
            # Bordas
            border_enabled_var.set(DEFAULT_APP_CONFIG["border_enabled"] == "1")
            border_mode_var.set(_mode_disp.get(DEFAULT_APP_CONFIG["border_mode"], "Rotação (cíclica)"))
            border_th_var.set(BORDER_THICKNESS_LABELS.get(DEFAULT_APP_CONFIG["border_thickness"], "Média"))
            for bp, v in border_model_vars.items():
                v.set(True)
            preview_border_var.set("(sem borda)")
            preview_config["_border_preview"] = None
            self._render_label_preview(preview_canvas, preview_config, hover_state["key"])

        ttk.Button(actions, text="Restaurar padrões", command=reset_defaults).pack(side=LEFT)
        ttk.Button(actions, text="Cancelar",          command=_cancel_settings).pack(side=RIGHT)
        ttk.Button(actions, text="Salvar e recarregar", command=save_and_reload).pack(side=RIGHT, padx=(0, 8))

        # Habilita rolagem por roda do mouse em todo o conteúdo das abas roláveis
        for _canvas, _inner in _scroll_targets:
            _bind_wheel_recursive(_inner, _canvas)
            _canvas.update_idletasks()
            _canvas.configure(scrollregion=_canvas.bbox("all"))

    # ── Impressão ──
    def _get_selected_pedido(self):
        pedidos = self._get_selected_pedidos()
        return pedidos[0] if pedidos else None

    def _ui_call(self, fn):
        """Encaminha fn para a thread da UI por fila (thread-safe em qualquer
        build do Tcl, ao contrário de chamar after() de outra thread)."""
        try:
            self._ui_queue.put(fn)
        except Exception:
            pass

    def _process_ui_queue(self):
        """Drena a fila de chamadas vindas de threads (impressão, monitor…)."""
        try:
            while True:
                fn = self._ui_queue.get_nowait()
                try:
                    fn()
                except Exception as e:
                    _log_error(f"Falha em callback de UI: {e}")
        except queue.Empty:
            pass
        except Exception:
            pass
        try:
            self.after(100, self._process_ui_queue)
        except Exception:
            pass

    def _row_codigo(self, iid):
        """Código shopper de uma linha da tabela (coluna 'codigo')."""
        try:
            return self.tree.set(iid, "codigo") or None
        except Exception:
            return None

    def _get_selected_pedidos(self):
        """Retorna os pedidos selecionados na tabela, na ordem exibida."""
        pedidos = []
        for iid in self.tree.selection():
            codigo = self._row_codigo(iid)
            if not codigo:
                continue
            for p in self._filtered_pedidos:
                if p["codigo"] == codigo:
                    pedidos.append(p)
                    break
        return pedidos

    def print_selected(self):
        if self._printing:
            self._show_toast("Aguarde: impressão em andamento…", ok=False)
            return
        pedidos = self._get_selected_pedidos()
        if not pedidos:
            messagebox.showwarning("Seleção obrigatória", "Selecione um pedido para imprimir.")
            return
        if len(pedidos) == 1:
            self.print_for(pedidos[0])
        else:
            self.print_batch(pedidos)

    def _ask_volumes(self, pedido, pos=None, total=None):
        title = "Imprimir etiquetas"
        if pos and total:
            title = f"Volumes — pedido {pos}/{total}"
        prompt = f"Quantos volumes?\nPedido: {pedido['codigo']}"
        if pedido.get("cliente"):
            prompt += f"\nCliente: {pedido['cliente']}"
        return simpledialog.askinteger(title, prompt, minvalue=1, maxvalue=99, parent=self)

    def _batch_default_volumes(self):
        try:
            return max(1, min(99, int(APP_CONFIG.get("batch_default_volumes", "1"))))
        except Exception:
            return 1

    def _build_pedido_payload(self, pedido, n, include_extra, border_pattern):
        chunks = []
        for vol in range(1, n + 1):
            chunks.append(build_epl_volume(
                codigo=pedido["codigo"], num_ifood=pedido["num_ifood"],
                cliente=pedido["cliente"], coleta=pedido["coleta"],
                vol_atual=vol, vol_total=n, border_pattern=border_pattern,
            ))
        if include_extra:
            chunks.append(build_epl_extra(
                codigo=pedido["codigo"], num_ifood=pedido["num_ifood"],
                cliente=pedido["cliente"], coleta=pedido["coleta"],
                border_pattern=border_pattern,
            ))
        return "".join(chunks)

    def _confirm_reprint(self, pedido, pos=None, total=None):
        """Se o pedido já tem check IMP, confirma antes de imprimir de novo."""
        if (pedido.get("codigo") or "").strip() not in getattr(self, "_printed_codes", set()):
            return True
        title = "Pedido já impresso"
        if pos and total:
            title += f" — {pos}/{total}"
        return messagebox.askyesno(
            title,
            f"O pedido {pedido['codigo']}\njá foi impresso anteriormente.\n\nDeseja imprimir novamente?",
            parent=self)

    def print_for(self, pedido):
        if not self._confirm_reprint(pedido):
            return
        n = self._ask_volumes(pedido)
        if not n: return
        include_extra = self._print_extra_enabled()
        total_labels = n + (1 if include_extra else 0)
        if total_labels > 5:
            detail = f"{n} volume(s)" + (" + 1 extra de coleta" if include_extra else "")
            if not self._confirm_large_print(total_labels, detail):
                return
        border_pattern = self._next_border_pattern()
        payload = self._build_pedido_payload(pedido, n, include_extra, border_pattern)
        if include_extra:
            summary = f"{total_labels} etiqueta(s) impressas ({n} vol. + 1 extra) — {pedido['codigo']}"
        else:
            summary = f"{total_labels} etiqueta(s) impressas ({n} vol., sem extra) — {pedido['codigo']}"
        meta = {"codigo": pedido["codigo"], "volumes": n}
        self._execute_print_jobs([(total_labels, payload, meta)], summary)

    def print_batch(self, pedidos):
        """Impressão em lote: a borda avança a cada pedido (modelos diferentes na
        pilha); todas as etiquetas do mesmo pedido saem com a mesma borda."""
        ask = APP_CONFIG.get("batch_ask_volumes", "1") == "1"
        default_vol = self._batch_default_volumes()
        include_extra = self._print_extra_enabled()

        specs, skipped = [], 0
        for i, pedido in enumerate(pedidos, 1):
            if not self._confirm_reprint(pedido, i, len(pedidos)):
                skipped += 1
                continue
            if ask:
                n = self._ask_volumes(pedido, i, len(pedidos))
                if not n:
                    skipped += 1
                    continue
            else:
                n = default_vol
            specs.append((pedido, n))
        if not specs:
            self._show_toast("Lote cancelado: nenhum pedido para imprimir.", ok=False)
            return

        total_labels = sum(n + (1 if include_extra else 0) for _, n in specs)
        detail = f"{len(specs)} pedido(s)" + (f", {skipped} pulado(s)" if skipped else "")
        detail += " — borda avança a cada pedido"
        if total_labels > 5 and not self._confirm_large_print(total_labels, detail):
            return

        jobs = []
        for pedido, n in specs:
            border_pattern = self._next_border_pattern()
            labels = n + (1 if include_extra else 0)
            meta = {"codigo": pedido["codigo"], "volumes": n}
            jobs.append((labels, self._build_pedido_payload(pedido, n, include_extra, border_pattern), meta))
        summary = f"Lote concluído: {total_labels} etiqueta(s) de {len(specs)} pedido(s)"
        if skipped:
            summary += f" ({skipped} pulado(s))"
        self._execute_print_jobs(jobs, summary)

    def _execute_print_jobs(self, jobs, done_message):
        """Envia os jobs em uma thread para não travar a UI; mostra progresso
        em jobs grandes (20+ etiquetas) e toast/barra de status ao concluir.
        Cada job é (n_etiquetas, payload, meta); com meta e histórico ativado,
        registra a impressão na aba 'Histórico de Impressão' após o envio —
        falha no histórico NUNCA bloqueia a impressão (vira aviso + log)."""
        total_labels = sum(c for c, _p, _m in jobs)
        show_progress = total_labels >= 20
        if show_progress:
            self.progress.configure(maximum=total_labels, value=0)
            self.progress.pack(side=RIGHT, padx=(0, 12))
        self._printing = True
        self.status_lbl.config(text=f"Imprimindo {total_labels} etiqueta(s)…")
        history_on = APP_CONFIG.get("history_enabled", "1") == "1"
        spreadsheet_id = APP_CONFIG.get("spreadsheet_id", "")

        def worker():
            printed = 0
            hist_fail = 0
            new_codes = []
            try:
                for count, payload, meta in jobs:
                    if FORCE_EPL_MODE:
                        payload = '! U1 setvar "device.languages" "epl"\r\n' + payload
                    send_to_printer(payload)
                    printed += count
                    if show_progress:
                        self._ui_call(lambda v=printed: self.progress.configure(value=v))
                    if meta:
                        new_codes.append(meta["codigo"])   # check IMP imediato, mesmo sem histórico
                    if history_on and meta:
                        try:
                            registrar_historico_impressao(
                                spreadsheet_id, meta["codigo"], count, meta["volumes"])
                        except Exception as he:
                            hist_fail += 1
                            _log_error(f"Histórico de Impressão falhou ({meta['codigo']}): "
                                       f"{_friendly_sheets_error(he)}")
                msg = done_message
                if hist_fail:
                    msg += f"  ⚠ histórico falhou em {hist_fail} pedido(s) — veja o log"
                self._ui_call(lambda m=msg, c=list(new_codes): self._finish_print(True, m, show_progress, c))
            except Exception as e:
                _log_error(f"Erro na impressão: {e}")
                self._ui_call(lambda err=str(e): self._finish_print(False, err, show_progress))

        threading.Thread(target=worker, daemon=True).start()

    def _finish_print(self, ok, message, show_progress, new_codes=None):
        self._printing = False
        if show_progress:
            self.progress.pack_forget()
        if ok:
            if new_codes:
                codes = {str(c).strip() for c in new_codes}
                self._session_printed = getattr(self, "_session_printed", set()) | codes
                self._printed_codes = getattr(self, "_printed_codes", set()) | codes
                if (APP_CONFIG.get("clear_search_after_print", "1") == "1"
                        and self.search_var.get().strip()):
                    self.search_var.set("")   # o trace refaz o filtro e re-renderiza
                else:
                    self._render_rows()
            self.status_lbl.config(text=message)
            self._show_toast(message, ok=("⚠" not in message))
        else:
            self.status_lbl.config(text="Erro na impressão")
            messagebox.showerror("Erro na impressão", message)

    # ── Toast (feedback não modal) ──
    def _show_toast(self, message, ok=True, ms=2600):
        try:
            toast = tk.Toplevel(self)
            toast.overrideredirect(True)
            toast.attributes("-topmost", True)
            bg = "#00A651" if ok else "#b45309"
            frame = tk.Frame(toast, bg=bg, padx=14, pady=8)
            frame.pack()
            tk.Label(frame, text=message, bg=bg, fg="white",
                     font=("Helvetica", 11, "bold"), justify="left").pack()
            toast.update_idletasks()
            x = self.winfo_rootx() + self.winfo_width() - toast.winfo_reqwidth() - 24
            y = self.winfo_rooty() + self.winfo_height() - toast.winfo_reqheight() - 24
            toast.geometry(f"+{max(x, 0)}+{max(y, 0)}")
            toast.after(ms, toast.destroy)
        except Exception:
            # Toast é cosmético: nunca pode derrubar o app.
            pass

    def _report_update_if_pending(self):
        """Após uma atualização, registra na aba LINK quando ela foi aplicada.
        Roda em thread; se a escrita falhar (ex.: 403), o marcador permanece
        e o registro é re-tentado no próximo refresh — bom para Pi sempre ligado."""
        if not getattr(self, "_pending_update_report", False):
            return
        self._pending_update_report = False   # evita corrida entre reloads
        sid = APP_CONFIG.get("spreadsheet_id", "")

        def worker():
            try:
                registrar_atualizacao_na_link(sid)
                try:
                    os.remove(UPDATE_MARKER_PATH)
                except Exception:
                    pass
                self._ui_call(lambda: self._show_toast(
                    f"Atualização para {APP_VERSION} registrada na aba LINK ✓", ok=True))
            except Exception as e:
                _log_error(f"Registro da atualização na LINK falhou: {_friendly_sheets_error(e)}")
                # marcador permanece: re-tenta no próximo reload
                self._pending_update_report = True

        threading.Thread(target=worker, daemon=True).start()

    # ── Painel de coleta (aba + telão) ──
    def _panel_active(self):
        """Polling só roda com a aba Painel visível ou o telão aberto."""
        try:
            if self._panel_fullscreen_win is not None and self._panel_fullscreen_win.winfo_exists():
                return True
            return self.main_nb.select() == str(self.tab_painel)
        except Exception:
            return False

    def _panel_interval(self):
        try:
            return max(10, int(APP_CONFIG.get("panel_refresh_seconds", "30")))
        except Exception:
            return 30

    def _on_main_tab_changed(self, _e=None):
        if self._panel_active():
            self._panel_last_fetch = 0.0   # força refresh imediato ao entrar
            self._refresh_panels()

    def _panel_heartbeat(self):
        """Batimento de 1s: dispara a leitura quando o intervalo vencer,
        somente com o painel ativo (economia de cota nas máquinas de impressão)."""
        try:
            if (self._panel_active() and not self._panel_loading and not self._is_reloading
                    and time.time() - self._panel_last_fetch >= self._panel_interval()):
                self._panel_last_fetch = time.time()
                self._panel_fetch_async()
        except Exception as e:
            _log_error(f"Painel heartbeat: {e}")
        self.after(1_000, self._panel_heartbeat)

    def _panel_fetch_async(self):
        """Lê dash + volumes do histórico em thread; sem popups (telão)."""
        self._panel_loading = True
        sid = APP_CONFIG.get("spreadsheet_id", "")
        rng = APP_CONFIG.get("range_name", "")

        def worker():
            try:
                values = obter_dados_google_sheets(sid, rng)
                pedidos = parse_pedidos(values)
                volumes = ler_historico_volumes(sid)
                def apply():
                    self._all_pedidos = pedidos
                    self._history_volumes = volumes
                    self._update_status_filter_options()
                    self._apply_filter()          # tabela ganha o refresh de graça
                    self._refresh_panels()
                self._ui_call(apply)
            except Exception as e:
                _log_error(f"Painel: falha ao atualizar: {e}")
                self._ui_call(lambda: self._refresh_panels(offline=True))
            finally:
                self._panel_loading = False

        threading.Thread(target=worker, daemon=True).start()

    def _detect_cancellations(self):
        """Alerta de devolução: dispara apenas na TRANSIÇÃO de um status
        acompanhado para cancelado (nunca para cancelados pré-existentes
        na primeira leitura, nem repetido para o mesmo pedido)."""
        cancel_set = _parse_status_list(
            APP_CONFIG.get("panel_cancel_statuses", "cancelled, cancellation request"))
        prev = self._panel_prev_status
        first_run = not prev
        novos = {}
        for p in self._all_pedidos:
            codigo = (p.get("codigo") or "").strip()
            if not codigo:
                continue
            st = _norm_status(p.get("status"))
            novos[codigo] = st
            if first_run or st not in cancel_set:
                continue
            anterior = prev.get(codigo)
            if anterior is not None and anterior not in cancel_set:
                vol = self._history_volumes.get(codigo, "")
                self._cancel_alerts.append({
                    "codigo": codigo,
                    "num": str(p.get("num_ifood") or "—"),
                    "volumes": vol,
                    "hora": datetime.now().strftime("%H:%M"),
                })
                _log_error(f"Painel: pedido cancelado — devolver "
                           f"(iFood {p.get('num_ifood')}, código {codigo})")
        self._panel_prev_status = novos

    def ack_cancel_alerts(self):
        """Botão CIENTE: limpa os avisos em todos os painéis (aba e telão)."""
        self._cancel_alerts = []
        for frame in list(self._panel_frames):
            try:
                if frame.winfo_exists():
                    frame.set_alerts([])
            except Exception:
                pass
        self._update_panel_tab_badge()

    def _update_panel_tab_badge(self):
        try:
            text = "  Painel ⚠  " if self._cancel_alerts else "  Painel  "
            self.main_nb.tab(self.tab_painel, text=text)
        except Exception:
            pass

    def _refresh_panels(self, offline=False):
        self._detect_cancellations()
        prep_set  = _parse_status_list(APP_CONFIG.get("panel_preparing_statuses", "separation started"))
        going_set = _parse_status_list(APP_CONFIG.get("panel_going_statuses", "assign driver"))
        ready_set = _parse_status_list(APP_CONFIG.get("panel_ready_statuses", "arrived at origin"))
        sep_done_set = _parse_status_list(APP_CONFIG.get("panel_sep_done_statuses", "separation ended"))
        tracked = prep_set | going_set | ready_set | sep_done_set

        self._panel_flags_seeded = True
        agora = time.time()
        vivos = set()
        cols = {"preparing": [], "going": [], "ready": []}
        for p in self._all_pedidos:
            st = _norm_status(p.get("status"))
            codigo = (p.get("codigo") or "").strip()
            if not st or not codigo:
                continue
            vivos.add(codigo)
            flags = self._panel_flags.setdefault(
                codigo, {"sep": False, "assigned": False, "arrived": False})

            # memória: flags só LIGAM conforme os status são observados
            # (segura o card na coluna quando o status oscila entre a trilha
            # da separação e a do motoboy)
            if st in sep_done_set:
                flags["sep"] = True
            if st in going_set:
                flags["assigned"] = True
            if st in ready_set:
                flags["arrived"] = True

            if st not in tracked:
                continue   # concluído/despachado/etc: fora do painel

            # coluna segue o MOTOBOY; "não preparado" é sinalizado no card
            # (sem volumes no histórico = etiquetas não impressas = SEPARANDO)
            if flags["arrived"]:
                col = "ready"
            elif flags["assigned"]:
                col = "going"
            else:
                col = "preparing"

            # cronômetro: marca quando o pedido ENTROU na coluna atual
            since = self._panel_col_since.get(codigo)
            if since is None or since[0] != col:
                since = (col, agora)
                self._panel_col_since[codigo] = since

            vol = self._history_volumes.get(codigo, "")
            row_ix = p.get("_row_ix", 10**9)
            oper = (str(p.get("operador") or "").strip().split() or [""])[0].capitalize()
            cols[col].append((row_ix, str(p.get("num_ifood") or "—"), vol,
                              since[1] if col == "preparing" else None,
                              oper if col == "preparing" else ""))

        # memória não cresce para sempre: esquece códigos que saíram da planilha
        self._panel_flags = {c: f for c, f in self._panel_flags.items() if c in vivos}
        self._panel_col_since = {c: s for c, s in self._panel_col_since.items() if c in vivos}

        # urgência: os pedidos mais antigos (linhas mais acima na planilha) primeiro
        for k in cols:
            cols[k] = [(num, vol, ts, oper)
                       for _ix, num, vol, ts, oper in sorted(cols[k], key=lambda t: t[0])]
        stamp = datetime.now().strftime("%H:%M")
        info = (f"sem conexão — última atualização {stamp}" if offline
                else f"atualizado {stamp} • a cada {self._panel_interval()}s")
        for frame in list(self._panel_frames):
            try:
                if frame.winfo_exists():
                    frame.update_data(cols, info)
                    frame.set_alerts(self._cancel_alerts)
            except Exception:
                pass
        self._update_panel_tab_badge()

    def toggle_panel_fullscreen(self):
        """Modo telão: fullscreen sem bordas com o painel; Esc/F11 sai."""
        win = self._panel_fullscreen_win
        if win is not None and win.winfo_exists():
            self._close_panel_fullscreen()
            return
        win = tk.Toplevel(self)
        # geometria explícita: garante tela cheia mesmo sem window manager
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        win.geometry(f"{sw}x{sh}+0+0")
        try:
            win.attributes("-fullscreen", True)
        except Exception:
            win.overrideredirect(True)   # fallback: sem bordas
        win.configure(bg=PANEL_BG)
        frame = PainelFrame(win, self, fullscreen=True)
        frame.pack(fill=BOTH, expand=YES)
        self._panel_frames.append(frame)
        self._panel_fullscreen_win = win
        win.bind("<Escape>", lambda _e: self._close_panel_fullscreen())
        win.bind("<F11>",    lambda _e: self._close_panel_fullscreen())
        win.protocol("WM_DELETE_WINDOW", self._close_panel_fullscreen)
        self._panel_last_fetch = 0.0   # refresh imediato ao abrir o telão
        self._refresh_panels()

    def _close_panel_fullscreen(self):
        win = self._panel_fullscreen_win
        self._panel_fullscreen_win = None
        if win is not None:
            self._panel_frames = [f for f in self._panel_frames
                                  if str(f).find(str(win)) != 0]
            try:
                win.destroy()
            except Exception:
                pass

    # ── Indicador de conexão da impressora ──
    def _start_printer_monitor(self):
        self._printer_online = None
        self._printer_tick()

    def _printer_tick(self):
        def check():
            online = _printer_available()
            self._ui_call(lambda: self._set_printer_status(online))
        threading.Thread(target=check, daemon=True).start()
        self.after(15_000, self._printer_tick)

    # ── Seletor de Loja (Dark) ──
    def _initial_dark_name(self):
        sid = APP_CONFIG.get("spreadsheet_id", "")
        for name, s in DARK_STORES.items():
            if s == sid:
                return name
        return DARK_OUTRA

    def _on_dark_selected(self, _event=None):
        name = self.dark_var.get()
        prev_name = self._initial_dark_name()
        if name == DARK_OUTRA:
            atual = APP_CONFIG.get("spreadsheet_id", "")
            sid = simpledialog.askstring(
                "Outra loja",
                "Cole o ID da planilha (spreadsheet_id):",
                initialvalue=atual, parent=self)
            if not sid or not sid.strip():
                self.dark_var.set(prev_name)   # cancelou: volta à anterior
                return
            sid = sid.strip()
        else:
            sid = DARK_STORES[name]
        if sid == APP_CONFIG.get("spreadsheet_id", ""):
            return
        APP_CONFIG["spreadsheet_id"] = sid
        APP_CONFIG["dark_store"] = name
        try:
            save_app_config(dict(APP_CONFIG))
        except Exception as e:
            _log_error(f"Falha ao salvar seleção de loja: {e}")
        self._printed_codes = set()
        self._session_printed = set()
        self._show_toast(f"Loja: {name}", ok=True)
        self.reload()

    # ── Log de erros (modo desenvolvedor) ──
    def _refresh_devlog_button(self):
        if APP_CONFIG.get("dev_mode", "0") == "1":
            self.devlog_btn.pack(side=RIGHT, padx=(0, 12))
        else:
            self.devlog_btn.pack_forget()

    def _open_error_log(self):
        win = tk.Toplevel(self)
        win.title("Log de erros (modo desenvolvedor)")
        win.geometry("820x420")
        txt = tk.Text(win, wrap="none", font=("Courier", 9))
        ysb = ttk.Scrollbar(win, orient="vertical", command=txt.yview)
        txt.configure(yscrollcommand=ysb.set)
        bar = ttk.Frame(win, padding=6)
        bar.pack(side="bottom", fill=X)
        ysb.pack(side=RIGHT, fill=Y)
        txt.pack(fill=BOTH, expand=True)

        def load():
            txt.configure(state="normal")
            txt.delete("1.0", "end")
            try:
                with open(ERROR_LOG_PATH, encoding="utf-8") as fh:
                    lines = fh.readlines()[-500:]
                txt.insert("1.0", "".join(lines) or "(log vazio)")
            except FileNotFoundError:
                txt.insert("1.0", "(nenhum erro registrado)")
            except Exception as e:
                txt.insert("1.0", f"(falha ao ler o log: {e})")
            txt.configure(state="disabled")
            txt.see("end")

        def clear():
            try:
                open(ERROR_LOG_PATH, "w").close()
            except Exception:
                pass
            load()

        ttk.Button(bar, text="Atualizar", command=load).pack(side=LEFT)
        ttk.Button(bar, text="Limpar log", command=clear).pack(side=LEFT, padx=(8, 0))
        ttk.Button(bar, text="Fechar", command=win.destroy).pack(side=RIGHT)
        load()

    def _set_printer_status(self, online):
        if online == getattr(self, "_printer_online", None):
            return
        self._printer_online = online
        if online:
            self.printer_lbl.config(text="Impressora: ● online", foreground="#00A651")
        else:
            self.printer_lbl.config(text="Impressora: ● offline", foreground="#dc2626")

    def _enabled_border_patterns(self):
        """Modelos habilitados para a rotação (checkboxes nas configurações)."""
        raw = APP_CONFIG.get("border_models", "") or ""
        enabled = [p.strip() for p in raw.split(",") if p.strip() in BORDER_PATTERNS]
        return enabled or list(BORDER_PATTERNS)

    def _next_border_pattern(self):
        """Define o padrão de borda do próximo job de impressão.
        Mesmo padrão para todas as etiquetas do lote (volumes + extra).
        Em modo 'rotate', avança o contador e persiste para continuar de onde parou.
        A rotação/sorteio usa apenas os modelos habilitados nas configurações."""
        if APP_CONFIG.get("border_enabled", "0") != "1":
            return None
        patterns = self._enabled_border_patterns()
        if APP_CONFIG.get("border_mode", "rotate") == "random":
            return random.choice(patterns)
        try:
            idx = int(APP_CONFIG.get("border_index", "0"))
        except Exception:
            idx = 0
        pattern = patterns[idx % len(patterns)]
        APP_CONFIG["border_index"] = str((idx + 1) % len(patterns))
        try:
            save_app_config(dict(APP_CONFIG))
        except Exception:
            pass
        return pattern

    def print_border_test(self):
        """Imprime uma etiqueta de cada modelo de borda (incl. 'sem borda'),
        com o nome do modelo em cada uma. Não altera o contador de rotação."""
        patterns = [None] + list(BORDER_PATTERNS)
        total = len(patterns)
        if not self._confirm_large_print(
                total, f"{total} etiquetas de teste (1 de cada modelo de borda)"):
            return
        try:
            chunks = []
            for i, pat in enumerate(patterns, 1):
                name = "(sem borda)" if pat is None else BORDER_PATTERN_LABELS.get(pat, pat)
                chunks.append(build_epl_border_test(i, total, pat, name))
            self._execute_print_jobs(
                [(total, "".join(chunks), None)],
                f"{total} etiquetas de teste enviadas (1 de cada modelo de borda).")
        except Exception as e:
            messagebox.showerror("Erro na impressão", str(e))


def main():
    App().mainloop()

if __name__ == "__main__":
    main()
