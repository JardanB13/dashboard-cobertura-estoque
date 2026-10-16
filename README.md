# Cobertura Estoque · Almoxarifados + Equipamentos

Dashboard (mapa) com a cobertura dos almoxarifados e, por cidade, os equipamentos **novos**, **reaproveitados** e **retirados**, com a quebra por **Natureza** (fonte: SINAPSE).

## Estrutura

```
index.html                      # o dashboard (Vue 3 + MapLibre, via CDN)
data/cobertura.js               # almoxarifados, cidades e distâncias
data/equipamentos.js            # novos x reaproveitados x retirados por cidade  ← atualizado pelo script
data/ufs.js                     # limites estaduais (IBGE), usado quando o mapa de ruas não carrega
scripts/atualizar_equipamentos.py
```

## Publicar no GitHub Pages (uma vez)

1. Crie um repositório no GitHub (ex.: `dashboard-cobertura-estoque`) e envie estes arquivos.
2. No repositório: **Settings → Pages → Build and deployment → Source: Deploy from a branch**, branch `main`, pasta `/ (root)`. Salve.
3. Em 1–2 minutos o site fica em `https://<seu-usuario-ou-org>.github.io/dashboard-cobertura-estoque/`.

> Repositório **público** = site público. Para deixar restrito à empresa, use um repositório privado em uma organização com GitHub Enterprise (Pages com controle de acesso).

## Atualizar os dados de equipamentos

```bash
pip install -r requirements.txt
python scripts/atualizar_equipamentos.py "C:\caminho\Relatorio_SINAPSE.xlsx"
git add data/equipamentos.js
git commit -m "Atualiza equipamentos"
git push
```

O site atualiza sozinho após o `push`.

**Nunca suba a planilha do SINAPSE**: ela tem nome e CPF/CNPJ de clientes. O `.gitignore` já bloqueia `.xlsx`/`.csv`, e o script grava no site apenas totais por cidade.

## Regras de cálculo

- Só itens com N° de Série (equipamentos); materiais de consumo ficam fora.
- `REAP-` = reaproveitado · `RET-` = retirado do cliente · sem prefixo = novo.
- Região pelo contrato: OP/RD final 1 = PB, 2 = RN, 3 = PE, 4 = BA.
- Cidades da planilha que não têm almoxarifado vinculado no mapa não aparecem (o script informa quantas).
