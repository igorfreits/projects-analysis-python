import pandas as pd
import openpyxl

data_path = 'data-analysis-python/Controle - Data Quality - FCM/'
file_path = data_path + 'Controle - FCM.xlsx'

controle_xls = pd.ExcelFile(file_path)
before = controle_xls.parse(sheet_name='Before')
after = controle_xls.parse(sheet_name='After')

# Garante que as chaves estejam como texto
for col_key in ['Handle ACC', 'Grupo Empresarial', 'Cliente']:
    before[col_key] = before[col_key].astype(str)
    after[col_key] = after[col_key].astype(str)

# Definir chaves fixas para merge
chaves = ['Handle ACC', 'Grupo Empresarial', 'Cliente', 'Data de Emissão']

# Faz merge usando as chaves
merged = pd.merge(before, after, on=chaves, how='inner', suffixes=('_before', '_after'))

# Lista de colunas a comparar (excluindo as chaves)
colunas_comuns = [col for col in before.columns if col not in chaves]

diferencas = []

for col in colunas_comuns:
    col_b = f"{col}_before"
    col_a = f"{col}_after"
    
    diferentes = merged[
        (merged[col_b] != merged[col_a]) |
        (merged[col_b].isnull() != merged[col_a].isnull())
    ][chaves + [col_b, col_a]]
    
    diferentes['Campo'] = col
    diferentes = diferentes.rename(columns={col_b: 'Valor Antes', col_a: 'Valor Depois'})
    
    colunas_ordenadas = chaves + ['Campo', 'Valor Antes', 'Valor Depois']
    diferentes = diferentes[colunas_ordenadas]
    
    diferencas.append(diferentes)

resultado_final = pd.concat(diferencas, ignore_index=True)

with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    resultado_final.to_excel(writer, sheet_name='Comparativo Benner', index=False)

print("Comparação concluída com sucesso.")
