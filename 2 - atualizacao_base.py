import pandas as pd
from datetime import datetime

# Caminho dos arquivos
data_path = 'PROCESSADO ERRO/Analise de Dados/'

# Leitura dos arquivos
relatorio_dash = pd.read_excel(
    data_path + 'Relatorio - Dash.xlsx', sheet_name='Processado Erro - BASE')
base_novo = pd.read_excel(data_path + 'Base.xlsx', sheet_name='Novo Arquivo')
base_em_andamento = pd.read_excel(
    data_path + 'Base.xlsx', sheet_name='Benner - Processado Erro 0')
base_resolvidos = pd.read_excel(
    data_path + 'Base.xlsx', sheet_name='Resolvidos')

# Data atual
data_atual = datetime.today().strftime('%m/%d/%Y')

# 1. Excluir todas as linhas do processado_erro
base_novo = pd.DataFrame()

# 2. Criar as novas colunas e preenchÃª-las com as informaÃ§Ãµes do relatorio_dash
novas_colunas = [
    'Status', 'Handle PNR', 'Handle ACC', 'Localizadora', 'RequisiÃ§Ã£o',
    'OBTS', 'Grupo Empresarial', 'ServiÃ§o', 'Mensagem Erro', 'TIPO DE ERRO',
    'RESPONSÃVEL', 'Data InclusÃ£o'
]

# Cria o DataFrame com as novas colunas
base_novo = pd.DataFrame(columns=novas_colunas)

# Preenche as novas colunas com os dados do relatorio_dash
for coluna in novas_colunas[1:]:  # Excluindo 'Status'
    if coluna in relatorio_dash.columns:
        base_novo[coluna] = relatorio_dash[coluna]
    else:
        base_novo[coluna] = pd.NA

# 3. Aplica o processo anterior com o DataFrame atualizado
# 3.1 Verificar se o "Handle ACC" do processado_erro estÃ¡ em base_em_andamento ou base_resolvidos


def verificar_status(handle_acc):
    if handle_acc in base_em_andamento['Handle ACC'].values:
        return 'Em Andamento'
    elif handle_acc in base_resolvidos['Handle ACC'].values:
        return 'Resolvido'

    else:
        return 'Novo'


# Aplica a funÃ§Ã£o para determinar o status
base_novo['Status'] = base_novo['Handle ACC'].apply(verificar_status)

# 3.2 Print de tudo que Ã© novo em processado_erro antes de copiar para base_em_andamento
novos_registros = base_novo[base_novo['Status'] == 'Novo'].copy()
print(f'Casos novos: {len(novos_registros)}')

# Atualiza o status dos novos registros para "Em Andamento" e copia para base_em_andamento
novos_registros['Status'] = 'Em Andamento'
base_em_andamento = pd.concat(
    [base_em_andamento, novos_registros], ignore_index=True)

# 3.3 Verificar o base_em_andamento


def atualizar_status(handle_acc):
    if handle_acc in base_novo['Handle ACC'].values:
        return 'Em Andamento'
    else:
        return 'Resolvido'


# Aplica a funÃ§Ã£o para determinar o status
base_em_andamento['Status'] = base_em_andamento['Handle ACC'].apply(
    atualizar_status)

# Filtra os registros resolvidos
registros_resolvidos = base_em_andamento[base_em_andamento['Status'] == 'Resolvido'].copy(
)
registros_resolvidos['Data de ConclusÃ£o'] = data_atual

# 3.4 Print do total em andamento em base_em_andamento
qtd_em_andamento = len(relatorio_dash)
print(f'Total Processado Erro Hoje: {qtd_em_andamento}')

# 3.5 Print dos registros resolvidos com a data de hoje em base_resolvidos
print(f'Casos resolvidos hoje: {len(registros_resolvidos)}')

# Move os registros resolvidos para base_resolvidos
base_resolvidos = pd.concat(
    [base_resolvidos, registros_resolvidos], ignore_index=True)
base_resolvidos['Data de ConclusÃ£o'] = pd.to_datetime(
    base_resolvidos['Data de ConclusÃ£o']).dt.date

# Remove os registros resolvidos de base_em_andamento
base_em_andamento = base_em_andamento[base_em_andamento['Status'] != 'Resolvido']
base_em_andamento['Data InclusÃ£o'] = pd.to_datetime(
    base_em_andamento['Data InclusÃ£o']).dt.date
base_novo['Data InclusÃ£o'] = pd.to_datetime(
    base_novo['Data InclusÃ£o']).dt.date

# Salva o DataFrame atualizado em um novo arquivo Excel, mantendo as outras guias intactas
output_path = data_path + 'Base.xlsx'

with pd.ExcelWriter(output_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    base_novo.to_excel(writer, sheet_name='Novo Arquivo', index=False)
    base_em_andamento.to_excel(
        writer, sheet_name='Benner - Processado Erro 0', index=False)
    base_resolvidos.to_excel(writer, sheet_name='Resolvidos', index=False)

print(f"\nO arquivo foi salvo com sucesso em {output_path}")
