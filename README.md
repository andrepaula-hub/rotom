# Rotom

Software de impressão de etiquetas iFood usado nos Raspberrys das dark stores.

## MVP de release

O pacote publicado preserva credenciais e config locais no Raspberry:

- `ifood/token.json`
- `ifood/client_secret.json`
- `ifood/pedidos_ifood_gui_config.json`

Gerar ZIP:

```bash
bash scripts/build_zip.sh
```

O ZIP sai em `dist/rotom-<hash>.zip` e inclui:

- `ifood/pedidos_ifood_gui.py`
- `ifood/print_bridge.py`
- `ifood/rotom_lite.py`
- `ifood/rotom_version.json`

O hash do commit publicado é a versão exibida no app.

## Ponte de impressão pro Alakazam (`ifood/print_bridge.py`)

Servidor HTTP local (porta 9876), sobe numa thread junto com o app. Existe
pra o navegador (tela de etiquetas do Alakazam) conseguir imprimir na fila
raw da Zebra — o Chromium sozinho não consegue, não tem driver ZPL
instalado. A página pede o ZPL pronto ao backend do Alakazam e manda pra cá
via `fetch`; aqui só roda `lp -o raw`, mesmo comando que `send_to_printer()`
já usa pro EPL.

## Rotom lite (`ifood/rotom_lite.py`)

Variante sem GUI, sem login de Google Sheets, sem iFood — só o
`print_bridge` (acima) + auto-update, consultando o mesmo manifest que o
Rotom completo usa. Pra loja que só precisa imprimir pelo Alakazam, sem a
tela de pedidos iFood do Rotom.

Instalar num Raspberry:

```bash
git clone https://github.com/andrepaula-hub/rotom.git /home/pi/rotom
sudo cp /home/pi/rotom/scripts/rotom-lite.service /etc/systemd/system/
sudo systemctl enable --now rotom-lite
```

Sem `venv`, sem dependência nenhuma (só biblioteca padrão do Python) — não
precisa de `requirements.txt` pra essa variante. `systemctl status
rotom-lite` mostra se subiu; `journalctl -u rotom-lite -f` acompanha o log.
