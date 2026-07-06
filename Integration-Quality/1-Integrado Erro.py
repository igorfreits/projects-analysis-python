import pandas as pd
import os

usuario = os.getlogin()
data_path = f'C:\\Users\\{usuario}\\Desktop\\DOCS\\data-analysis-python\\Integration-Quality\\'

arquivo_excel = pd.ExcelFile(data_path + 'Relatorio - Integratour.xlsx')
base_novo = arquivo_excel.parse('Novo Arquivo')
base_andamento = arquivo_excel.parse('Integrado Erro')
base_resolvidos = arquivo_excel.parse('Resolvidos')

data_hoje = pd.to_datetime('today').date()

# Preencher campos obrigatórios em branco
base_andamento['MOTIVO DO ERRO'] = base_andamento['MOTIVO DO ERRO'].fillna('Erro não identificado')
base_andamento['DETALHES DO ERRO'] = base_andamento['DETALHES DO ERRO'].fillna('Não identificado')
base_andamento['CATEGORIA DO ERRO'] = base_andamento['CATEGORIA DO ERRO'].fillna('Sistêmico')

# Pré-computar coluna MENSAGEM em minúsculas uma única vez para todas as comparações
msg = base_andamento['MENSAGEM'].str.lower().fillna('')

#---Tratamento de Erros---

# Cliente não identificado
base_andamento.loc[
    msg.str.contains('código de cliente não identificado', regex=False) |
    msg.str.contains('código do cliente não informado!', regex=False),
    ['MOTIVO DO ERRO', 'DETALHES DO ERRO', 'CATEGORIA DO ERRO']
] = ['Cliente não identificado', 'DK de cliente não preenchido/cliente não configurado no OBT', 'Processo Operacional']

base_andamento.loc[
    msg.str.contains('converterpagamentosremark', regex=False) |
    msg.str.contains('input string was not in a correct format', regex=False) |
    msg.str.contains('verificando anexos - could not convert variant of type (null) into type (olestr)', regex=False),
    ['MOTIVO DO ERRO', 'DETALHES DO ERRO', 'CATEGORIA DO ERRO']
] = ['Formato de texto inválido', 'Texto fora do padrão aceito no campo', 'Qualidade dos Dados']

# Cancelamento de reserva
base_andamento.loc[
    msg.str.contains('cancelamento de venda', regex=False),
    ['MOTIVO DO ERRO', 'DETALHES DO ERRO', 'CATEGORIA DO ERRO']
] = ['Cancelamento de reserva', 'Reserva cancelada no OBT', 'Sistêmico']

base_andamento.loc[
    msg.str.contains('emissor de código  não encontrado', regex=False) |
    msg.str.contains('consultando agente-> consultando agente-> o agente', regex=False),
    ['MOTIVO DO ERRO', 'DETALHES DO ERRO', 'CATEGORIA DO ERRO']
] = ['Emissor não encontrado/Inativo', 'Código do emissor não cadastrado no OBT', 'Sistêmico']

# Codigo do fornecedor
base_andamento.loc[
    msg.str.contains('não foi possível localizar o contrato do fornecedor', regex=False) |
    msg.str.contains('fornecedor com o código', regex=False) |
    msg.str.contains('fornecedor com o apelido', regex=False),
    ['MOTIVO DO ERRO', 'DETALHES DO ERRO', 'CATEGORIA DO ERRO']
] = ['Contrato do Fornecedor', 'Contrato inativo ou divergente', 'Sistêmico']

# Codigo do cliente
base_andamento.loc[
    msg.str.contains('não localizado cliente código', regex=False),
    ['MOTIVO DO ERRO', 'DETALHES DO ERRO', 'CATEGORIA DO ERRO']
] = ['Código do cliente divergente', 'Código diferente do cadasto no Benner (Cliente X Grupo)', 'Qualidade dos Dados']

# Remark enviado errado
base_andamento.loc[
    msg.str.contains('não localizado cliente com cnpj', regex=False),
    ['MOTIVO DO ERRO', 'DETALHES DO ERRO', 'CATEGORIA DO ERRO']
] = ['Remark enviado errado', 'Remark de CNPJ enviado em formato inválido', 'Qualidade dos Dados']

# Assento/bagagem
base_andamento.loc[
    msg.str.contains('não foi informado o localizador do assento/bagagem!', regex=False) |
    msg.str.contains('não encontrada para importar assento/bagagem!', regex=False),
    ['MOTIVO DO ERRO', 'DETALHES DO ERRO', 'CATEGORIA DO ERRO']
] = ['Assento/Bagagem não informado', 'Campo OBS preenchido incorretamente', 'Qualidade dos Dados']

# RLOC não informado
base_andamento.loc[
    msg.str.contains('foi identificado um localizador sem código', regex=False),
    ['MOTIVO DO ERRO', 'DETALHES DO ERRO', 'CATEGORIA DO ERRO']
] = ['RLOC não informado', 'RLOC não preenchido na reserva', 'Qualidade dos Dados']

# Dados do Cartão
base_andamento.loc[
    msg.str.contains('não foi possível determinar a validade do cartão', regex=False) |
    msg.str.contains("'long' does not contain a definition for 'tostring2'", regex=False) |
    msg.str.contains('erro: administradora não encontrada para bandeira', regex=False),
    ['MOTIVO DO ERRO', 'DETALHES DO ERRO', 'CATEGORIA DO ERRO']
] = ['Dados do Cartão', 'TAG de validade não enviada', 'Sistêmico']

# Tag de serviço não enviada
base_andamento.loc[
    msg.str.contains('não foi encontrado nenhum item veículo para o localizador', regex=False) |
    msg.str.contains('não foi encontrado nenhum item hotel para o localizador', regex=False) |
    msg.str.contains('não foi encontrado nenhum item aéreo para o localizador', regex=False),
    ['MOTIVO DO ERRO', 'DETALHES DO ERRO', 'CATEGORIA DO ERRO']
] = ['Tag de Serviço não enviada', 'TAG de serviço não enviada', 'Sistêmico']

# Número de VOO não informado/Inválido
base_andamento.loc[
    msg.str.contains('número de voo informado é inválido:', regex=False),
    ['MOTIVO DO ERRO', 'DETALHES DO ERRO', 'CATEGORIA DO ERRO']
] = ['Número de VOO não informado/Inválido', 'Número de VOO não preenchido ou inválido', 'Qualidade dos Dados']

# Centro de custo não informado
base_andamento.loc[
    msg.str.contains('centro de custo não informado!', regex=False),
    ['MOTIVO DO ERRO', 'DETALHES DO ERRO', 'CATEGORIA DO ERRO']
] = ['Centro de custo não informado', 'Campo centro de custo não preenchido', 'Qualidade dos Dados']

# Número de caracter excedido
base_andamento.loc[
    msg.str.contains('o tamanho máximo do campo "código" é 60 caracteres.', regex=False),
    ['MOTIVO DO ERRO', 'DETALHES DO ERRO', 'CATEGORIA DO ERRO']
] = ['Número de caracter excedido', 'Campo preenchido com mais de 60 caracteres', 'Qualidade dos Dados']

# Canal de venda não informado
base_andamento.loc[
    msg.str.contains('canal de venda com descrição', regex=False),
    ['MOTIVO DO ERRO', 'DETALHES DO ERRO', 'CATEGORIA DO ERRO']
] = ['Canal de venda não encontrado', 'Campo canal de venda não encontrado', 'Qualidade dos Dados']

# Erro de XML e/ou SQL
base_andamento.loc[
    msg.str.contains('the update statement conflicted with the foreign key constraint', regex=False) |
    msg.str.contains('the delete statement conflicted with the reference constraint', regex=False) |
    msg.str.contains('is specified more than once in the set clause or column list of an insert', regex=False) |
    msg.str.contains('inserindo accounting - (execsql) - list index out of bounds (0)', regex=False),
    ['MOTIVO DO ERRO', 'DETALHES DO ERRO', 'CATEGORIA DO ERRO']
] = ['Erro de XML e/ou SQL', 'Erro de processamento da reserva', 'Sistêmico']


# Garantir colunas extras no base_novo
for col in ['MOTIVO DO ERRO', 'DETALHES DO ERRO', 'CATEGORIA DO ERRO', 'DATA INCLUSÃO']:
    if col not in base_novo.columns:
        base_novo[col] = pd.NA

base_novo['DATA INCLUSÃO'] = data_hoje

# Classificar status
base_novo['STATUS'] = 'Novo'
base_novo.loc[base_novo['HANDLE'].isin(base_andamento['HANDLE']), 'STATUS'] = 'Em Andamento'
base_novo.loc[base_novo['HANDLE'].isin(base_resolvidos['HANDLE']), 'STATUS'] = 'Resolvido'

# Identificar registros novos e mover para andamento
novos = base_novo[base_novo['STATUS'] == 'Novo'].copy()
novos['STATUS'] = 'Em Andamento'
base_andamento = pd.concat([base_andamento, novos], ignore_index=True)

# Atualizar status de andamento com base no relatório atual
# Itens ausentes do relatório atual foram resolvidos; 'Novo' recém-adicionados ficam em andamento
status_temp = base_novo[['HANDLE', 'STATUS']]
base_andamento = base_andamento.merge(status_temp, on='HANDLE', how='left', suffixes=('', '_NOVO'))
resolved_mask = base_andamento['STATUS_NOVO'].isna() | (base_andamento['STATUS_NOVO'] == 'Resolvido')
base_andamento['STATUS'] = 'Em Andamento'
base_andamento.loc[resolved_mask, 'STATUS'] = 'Resolvido'
base_andamento.drop(columns=['STATUS_NOVO'], inplace=True)

# Identificar resolvidos
resolvidos = base_andamento[base_andamento['STATUS'] == 'Resolvido'].copy()
resolvidos['DATA CONCLUSÃO'] = data_hoje

# Adiciona aos resolvidos
base_resolvidos = pd.concat([base_resolvidos, resolvidos], ignore_index=True)
base_resolvidos['DATA CONCLUSÃO'] = pd.to_datetime(base_resolvidos['DATA CONCLUSÃO']).dt.date

# Remove resolvidos de andamento
base_andamento = base_andamento[base_andamento['STATUS'] != 'Resolvido']

# Formata datas
base_andamento['DATA INCLUSÃO'] = pd.to_datetime(base_andamento['DATA INCLUSÃO']).dt.date
base_novo['DATA INCLUSÃO'] = pd.to_datetime(base_novo['DATA INCLUSÃO']).dt.date

# Exporta planilha
with pd.ExcelWriter(data_path + 'Relatorio - Integratour.xlsx', engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
    base_novo.to_excel(writer, sheet_name='Novo Arquivo', index=False)
    base_andamento.to_excel(writer, sheet_name='Integrado Erro', index=False)
    base_resolvidos.to_excel(writer, sheet_name='Resolvidos', index=False)

# Avisos
print(f'\033[1;33m- Identificamos {qtd} novos erros não categorizados\033[m') if (qtd := len(base_andamento[base_andamento["MOTIVO DO ERRO"] == "Erro não identificado"])) > 1 else None
print(f'\033[94m- Total Integratour Hoje:\033[0m {len(base_novo)}')
print(f'\033[94m- Casos resolvidos hoje:\033[0m {len(resolvidos)}')
print(f'\033[94m- Casos novos:\033[0m {len(novos)}')

print(f'\033[1;32m\n- Relatório Integratour atualizado com sucesso!\033[m')
