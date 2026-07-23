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
- `ifood/rotom_version.json`

