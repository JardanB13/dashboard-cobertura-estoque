import json
import pandas as pd

# Carregando a base de dados original (substitua pelo nome do seu arquivo de origem)
# df = pd.read_excel('sua_base_de_dados.xlsx')

# Exemplo de tratamento e filtro para excluir obsoletos da contagem de aplicados em cliente
# Certifique-se de que a coluna de categoria e status estejam mapeadas corretamente no seu script:
df_filtrado = df[~df['Categoria'].str.upper().isin(['OBSOLETO', 'OBSOLETOS'])]

# Agrupando os dados para atualizar o JSON (mantendo a estrutura que alimenta o equipamentos.js)
tot_novos = df_filtrado[df_filtrado['Status'] == 'Novo']['Quantidade'].sum()
tot_reaproveitados = df_filtrado[df_filtrado['Status'] == 'Reaproveitado']['Quantidade'].sum()

# Atualize o dicionário que gera o JSON para o frontend:
dados_saida = {
    "meta": {
        "periodo": "set/2026",
        "tot": {
            "n": int(tot_novos),
            "r": int(tot_reaproveitados),
            # Mantenha as outras chaves conforme o seu script atual...
        }
    }
}

# Salvando o arquivo equipamentos.js atualizado com a variável global correta
with open('equipamentos.js', 'w', encoding='utf-8') as f:
    f.write(f"window.DADOS_EQUIPAMENTOS = {json.dumps(dados_saida, ensure_ascii=False, indent=2)};")

print("Arquivo equipamentos.js atualizado com sucesso!")
git add equipamentos.js atualizar_equipamentos.py
git commit -m "fix: alinha o total de material aplicado de setembro excluindo itens obsoletos"
git push origin main

