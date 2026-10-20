# Cobertura Estoque · Almoxarifados + Equipamentos

Dashboard (mapa) com a cobertura dos almoxarifados e, por cidade, os equipamentos **novos**, **reaproveitados** e **retirados** (fonte: SINAPSE), mais o **destino dos retirados** na triagem e o **status dos REAP convertidos**.

Site: https://jardanb13.github.io/dashboard-cobertura-estoque/

## Arquivos (todos na raiz)

```
index.html                  # o dashboard (Vue 3 + MapLibre, via CDN)
cobertura.js                # almoxarifados, cidades e distâncias
equipamentos.js             # novos x reaproveitados x retirados por cidade   ← atualizar_equipamentos.py
conversao.js                # destino dos retirados + status dos REAP         ← atualizar_conversao.py
ufs.js                      # limites estaduais (IBGE)
atualizar_equipamentos.py
atualizar_conversao.py
```

## Atualizar os dados do mês

```bash
pip install -r requirements.txt
python atualizar_equipamentos.py "C:\caminho\Aplicados_Nordeste.xlsx" "C:\caminho\Aplicados_Sul_Sudeste.xlsx" --categorias "C:\caminho\CATEGORIA_TECNOLOGICA_DOS_EQUIPAMENTOS.xlsx"
python atualizar_conversao.py "C:\caminho\Movimentacao_Nordeste.xlsx" "C:\caminho\Aplicados_Nordeste.xlsx" "C:\caminho\Movimentacao_Sul.xlsx" "C:\caminho\Aplicados_Sul_Sudeste.xlsx"
```

Depois suba `equipamentos.js` e `conversao.js` (Add file → Upload files → Commit changes). O site atualiza sozinho em ~1 minuto.

**Nunca suba as planilhas do SINAPSE**: elas têm nome e CPF/CNPJ de clientes. O `.gitignore` bloqueia `.xlsx`/`.csv`, e os scripts gravam no site apenas totais por cidade/almoxarifado.

## Regras de cálculo

- Só projetos de operação (Projeto começando com `OP-`) e só itens com N° de Série (equipamentos); materiais de consumo ficam fora.
- Categoria tecnológica (card "Equipamentos por categoria"): pelo nome do aparelho na planilha de categorias; modelos que não estão nela são classificados pelo nome (Wi-Fi 6/AX, AC/dual band, bridge, rádio). O script lista no final o que ficou em "Outros".
- Região: almoxarifados de SP, PR e MG = Sul/Sudeste; os demais = Norte/Nordeste.
- `REAP-` = reaproveitado · `RET-` = retirado do cliente · sem prefixo = novo.
- Estado da cidade pelo almoxarifado da coluna Projeto (ex.: OP-OSASCO → Osasco/SP); sem almoxarifado, pelo contrato (final 1 = PB, 2 = RN, 3 = PE, 4 = BA).
- `atualizar_equipamentos.py` aceita várias planilhas do mesmo mês (ex.: Nordeste + Sul/Sudeste) e soma tudo.
- Visão geral ("Todos") soma também as cidades sem almoxarifado vinculado; ao filtrar um almoxarifado, só as cidades dele.
- Conversão (movimentação CM): saída `RET-` e entrada do destino (`REAP-`, `SUC-`, `DET-`, `DEF-`) para a mesma série.
- Status de cada RET convertido em REAP no fim do período: **instalado** (aplicado no cliente depois da conversão),
  **com o técnico** (entregue por RM e não instalado nem devolvido) ou **no almoxarifado**.
