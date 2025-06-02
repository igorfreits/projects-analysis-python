import pandas as pd
from datetime import datetime

# Caminho dos arquivos
data_path = 'data-analysis-python/PROCESSADO ERRO/'

# Leitura dos arquivos
relatorio_dash = pd.read_excel(data_path + 'Relatorio - Dash.xlsx', sheet_name='Processado Erro - BASE')
base_xls = pd.ExcelFile(data_path + 'Base.xlsx')
base_novo = base_xls.parse(sheet_name='Novo Arquivo')
base_em_andamento = base_xls.parse(sheet_name='Benner - Processado Erro 0')
base_resolvidos = base_xls.parse(sheet_name='Resolvidos')

# Data atual
data_atual = pd.to_datetime('today').date()

# Definir as colunas desejadas
novas_colunas = [
    'Status', 'Handle PNR', 'Handle ACC', 'Localizadora', 'Status Requisicao',
    'OBTS', 'Grupo Empresarial', 'Serviço', 'Mensagem Erro', 'TIPO DE ERRO','EMPRESA', 'CATEGORIA DE ERRO',
    'RESPONSÁVEL', 'Data Inclusão'
]

# Criar base_novo diretamente do relatório, garantindo as colunas corretas
base_novo = relatorio_dash.reindex(columns=novas_colunas, fill_value=pd.NA)

# Atualizar status usando operações vetorizadas
base_novo['Status'] = 'Novo'
base_novo.loc[base_novo['Handle ACC'].isin(base_em_andamento['Handle ACC']), 'Status'] = 'Em Andamento'
base_novo.loc[base_novo['Handle ACC'].isin(base_resolvidos['Handle ACC']), 'Status'] = 'Resolvido'

# Identificar novos registros
novos_registros = base_novo[base_novo['Status'] == 'Novo'].copy()
print(f'\033[94m- Casos novos:\033[0m {len(novos_registros)}')

# Atualiza status para "Em Andamento" e adiciona na base_em_andamento
novos_registros['Status'] = 'Em Andamento'
base_em_andamento = pd.concat([base_em_andamento, novos_registros], ignore_index=True)

# Atualizar status de base_em_andamento com merge
temp_status = base_novo[['Handle ACC', 'Status']]
base_em_andamento = base_em_andamento.merge(temp_status, on='Handle ACC', how='left', suffixes=('', '_novo'))
base_em_andamento['Status'] = base_em_andamento['Status_novo'].fillna('Resolvido')
base_em_andamento.drop(columns=['Status_novo'], inplace=True)

# Identificar registros resolvidos
registros_resolvidos = base_em_andamento[base_em_andamento['Status'] == 'Resolvido'].copy()
registros_resolvidos['Data de Conclusão'] = data_atual

print(f'\033[94m- Total Processado Erro Hoje:\033[0m {len(relatorio_dash)}')
print(f'\033[94m- Casos resolvidos hoje:\033[0m {len(registros_resolvidos)}')

# Mover registros resolvidos para base_resolvidos
base_resolvidos = pd.concat([base_resolvidos, registros_resolvidos], ignore_index=True)
base_resolvidos['Data de Conclusão'] = pd.to_datetime(base_resolvidos['Data de Conclusão']).dt.date

# Remover registros resolvidos de base_em_andamento
base_em_andamento = base_em_andamento[base_em_andamento['Status'] != 'Resolvido']
base_em_andamento['Data Inclusão'] = pd.to_datetime(base_em_andamento['Data Inclusão']).dt.date
base_novo['Data Inclusão'] = pd.to_datetime(base_novo['Data Inclusão']).dt.date

# Salvar os dados no Excel sem modificar outras guias
output_path = data_path + 'Base.xlsx'
with pd.ExcelWriter(output_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
    base_novo.to_excel(writer, sheet_name='Novo Arquivo', index=False)
    base_em_andamento.to_excel(writer, sheet_name='Benner - Processado Erro 0', index=False)
    base_resolvidos.to_excel(writer, sheet_name='Resolvidos', index=False)

print(f"\033[92m\n- O arquivo foi salvo com sucesso em:\033[0m {output_path}")
