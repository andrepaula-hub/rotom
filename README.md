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
- `ifood/rotom_version.json`

O hash do commit publicado é a versão exibida no app.

## Ponte de impressão pro Alakazam (`ifood/print_bridge.py`)

Servidor HTTP local (porta 9876), sobe numa thread junto com o app. Existe
pra o navegador (tela de etiquetas do Alakazam) conseguir imprimir na fila
raw da Zebra — o Chromium sozinho não consegue, não tem driver ZPL
instalado. A página pede o ZPL pronto ao backend do Alakazam e manda pra cá
via `fetch`; aqui só roda `lp -o raw`, mesmo comando que `send_to_printer()`
já usa pro EPL.
