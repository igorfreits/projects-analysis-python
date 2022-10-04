import pandas as pd

table = pd.read_excel('data/vendas.xlsx')
# colunas = ['ID Loja', 'ID Produto', 'Quantidade', 'Valor Final']

# print(table)

# Valor total de vendas
faturamento_toral = table['Valor Final'].sum()
print(f'Faturamento Total: \033[35mR${faturamento_toral:.2f}\033[m')

# faturamento por loja
# [[]] - Usar dois colchetes para selecionar mais de uma coluna
faturamento_loja = table[['ID Loja', 'Valor Final']]
# print(faturamento_loja)

# Agrupar por loja
# groupby - Agrupar por uma coluna
faturamento_loja = faturamento_loja.groupby('ID Loja').sum()
print(faturamento_loja)

# Quantidade de produtos vendidos
faturamento_produto = table[['ID Loja',
                             'Produto', 'Valor Final']].groupby(['ID Loja', 'Produto']).sum()
print(faturamento_produto)
