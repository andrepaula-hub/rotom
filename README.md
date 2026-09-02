# Rotom

Ponte de impressão local pros Raspberrys das dark stores: o Alakazam gera a
etiqueta (pedido, produto, endereço) e imprime direto do navegador. Esse
repo é só o "tradutor" que roda em cada Raspberry — sem GUI, sem login de
Google Sheets, sem iFood. Essas features viviam aqui antes
(`ifood/pedidos_ifood_gui.py`) e saíram: o Alakazam assumiu a impressão de
pedido, então carregar esse peso todo em cada loja parou de fazer sentido.

## Como funciona

- `ifood/print_bridge.py` — servidor HTTP local (porta 9876). A tela de
  etiquetas do Alakazam pede o ZPL pronto ao backend e manda pra cá via
  `fetch`; aqui só roda `lp -o raw` (a fila da Zebra é raw, sem driver — o
  Chromium sozinho não consegue imprimir nela).
- `ifood/rotom_lite.py` — sobe o `print_bridge` numa thread e se
  auto-atualiza, consultando o manifest do Alakazam
  (`/internal/rotom-manifest`). Só biblioteca padrão do Python, sem `venv`
  nem dependência nenhuma.

## Instalar num Raspberry

```bash
git clone https://github.com/andrepaula-hub/rotom.git /home/pi/rotom
sudo cp /home/pi/rotom/scripts/rotom-lite.service /etc/systemd/system/
sudo systemctl enable --now rotom-lite
```

`systemctl status rotom-lite` mostra se subiu; `journalctl -u rotom-lite -f`
acompanha o log. O `systemd` garante que volta sozinho se a energia cair ou
o Raspberry reiniciar.

## Publicar uma versão nova

```bash
bash scripts/build_zip.sh
```

Gera `dist/rotom-<hash>.zip` com `ifood/print_bridge.py`,
`ifood/rotom_lite.py` e `ifood/rotom_version.json`. O hash do commit é a
versão que aparece no manifest — suba esse ZIP pro canal de distribuição do
Alakazam (`gs://darkstores-264e6.firebasestorage.app/rotom/`) pra todo
Raspberry rodando `rotom-lite` puxar sozinho.
