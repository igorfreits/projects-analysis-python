import pandas as pd
from datetime import datetime

# Caminho dos arquivos
data_path = 'data-analysis-python/Integration Quality/'

# Carregar Excel
arquivo_excel = pd.ExcelFile(data_path + 'Relatorio - Integratour.xlsx')
base_novo = arquivo_excel.parse('Novo Arquivo')
base_andamento = arquivo_excel.parse('Integrado Erro')
base_resolvidos = arquivo_excel.parse('Resolvidos')

# Data atual
data_hoje = pd.to_datetime('today').date()

# Preencher campos obrigatórios em branco
base_andamento['MOTIVO DO ERRO'].fillna('Erro não identificado', inplace=True)
base_andamento['DETALHES DO ERRO'].fillna('Não identificado', inplace=True)
base_andamento['CATEGORIA DO ERRO'].fillna('Sistêmico', inplace=True)

# Aplicar classificação de erro com base na coluna "MENSAGEM"
mensagem = base_andamento['MENSAGEM'].fillna('')

#---Tratamento de Erros---

# Cliente não identificado
base_andamento.loc[base_andamento['MENSAGEM'].str.contains('Código de cliente não identificado', case=False) |
                                base_andamento['MENSAGEM'].str.contains('Código do cliente não informado!', case=False),
                             ['MOTIVO DO ERRO', 'DETALHES DO ERRO', 'CATEGORIA DO ERRO']] = ['Cliente não identificado', 'DK de cliente não preenchido/cliente não configurado no OBT', 'Processo Operacional']

base_andamento.loc[base_andamento['MENSAGEM'].str.contains('ConverterPagamentosRemark', case=False) |
                                base_andamento['MENSAGEM'].str.contains('Input string was not in a correct format', case=False) |
                                base_andamento['MENSAGEM'].str.contains('Verificando Anexos - Could not convert variant of type (Null) into type (OleStr)', case=False),
                                ['MOTIVO DO ERRO', 'DETALHES DO ERRO', 'CATEGORIA DO ERRO']] = ['Formato de texto inválido', 'Texto fora do padrão aceito no campo', 'Qualidade dos Dados']

# Cancelamento de reserva
base_andamento.loc[base_andamento['MENSAGEM'].str.contains('Cancelamento de Venda', case=False), 
                             ['MOTIVO DO ERRO', 'DETALHES DO ERRO', 'CATEGORIA DO ERRO']] = ['Cancelamento de reserva', 'Reserva cancelada no OBT', 'Sistêmico']

base_andamento.loc[base_andamento['MENSAGEM'].str.contains('Emissor de código  não encontrado', case=False) |
                                base_andamento['MENSAGEM'].str.contains('Consultando Agente-> Consultando Agente-> O agente', case=False),
                             ['MOTIVO DO ERRO', 'DETALHES DO ERRO', 'CATEGORIA DO ERRO']] = ['Emissor não encontrado/Inativo', 'Código do emissor não cadastrado no OBT', 'Sistêmico']

# Codigo do fornecedor
base_andamento.loc[base_andamento['MENSAGEM'].str.contains('Não foi possível localizar o contrato do fornecedor', case=False) |
                                base_andamento['MENSAGEM'].str.contains('Fornecedor com o código', case=False) | 
                                base_andamento['MENSAGEM'].str.contains('Fornecedor com o apelido', case=False), 
                             ['MOTIVO DO ERRO', 'DETALHES DO ERRO', 'CATEGORIA DO ERRO']] = ['Contrato do Fornecedor', 'Contrato inativo ou divergente', 'Sistêmico']

# Codigo do cliente
base_andamento.loc[base_andamento['MENSAGEM'].str.contains('Não localizado cliente código', case=False), 
                             ['MOTIVO DO ERRO', 'DETALHES DO ERRO', 'CATEGORIA DO ERRO']] = ['Código do cliente divergente', 'Código diferente do cadasto no Benner (Cliente X Grupo)', 'Qualidade dos Dados']

# Remark enviado errado
base_andamento.loc[base_andamento['MENSAGEM'].str.contains('Não localizado cliente com CNPJ', case=False), 
                             ['MOTIVO DO ERRO', 'DETALHES DO ERRO', 'CATEGORIA DO ERRO']] = ['Remark enviado errado', 'Remark de CNPJ enviado em formato inválido', 'Qualidade dos Dados']

# Assento/bagagem
base_andamento.loc[base_andamento['MENSAGEM'].str.contains('Não foi informado o Localizador do Assento/Bagagem!', case=False)|
                                base_andamento['MENSAGEM'].str.contains('não encontrada para importar assento/bagagem!', case=False), 
                             ['MOTIVO DO ERRO', 'DETALHES DO ERRO', 'CATEGORIA DO ERRO']] = ['Assento/Bagagem não informado', 'Campo OBS preenchido incorretamente', 'Qualidade dos Dados']


# RLOC não informado
base_andamento.loc[base_andamento['MENSAGEM'].str.contains('Foi identificado um localizador sem código', case=False), 
                             ['MOTIVO DO ERRO', 'DETALHES DO ERRO', 'CATEGORIA DO ERRO']] = ['RLOC não informado', 'RLOC não preenchido na reserva', 'Qualidade dos Dados']

# Dados do Cartão
base_andamento.loc[base_andamento['MENSAGEM'].str.contains('Não foi possível determinar a validade do cartão', case=False) |
                                base_andamento['MENSAGEM'].str.contains("'long' does not contain a definition for 'toString2'", case=False) |
                                base_andamento['MENSAGEM'].str.contains('Erro: Administradora não encontrada para Bandeira', case=False),
                             ['MOTIVO DO ERRO', 'DETALHES DO ERRO', 'CATEGORIA DO ERRO']] = ['Dados do Cartão', 'TAG de validade não enviada', 'Sistêmico']

# Tag de serviço não enviada
base_andamento.loc[base_andamento['MENSAGEM'].str.contains('Não foi encontrado nenhum item veículo para o localizador', case=False)|
                                base_andamento['MENSAGEM'].str.contains('Não foi encontrado nenhum item hotel para o localizador', case=False)|
                                base_andamento['MENSAGEM'].str.contains('Não foi encontrado nenhum item aéreo para o localizador', case=False),
                             ['MOTIVO DO ERRO', 'DETALHES DO ERRO', 'CATEGORIA DO ERRO']] = ['Tag de Serviço não enviada', 'TAG de serviço não enviada', 'Sistêmico']

# Número de VOO não informado/Inválido
base_andamento.loc[base_andamento['MENSAGEM'].str.contains('Número de voo informado é inválido:', case=False),
                             ['MOTIVO DO ERRO', 'DETALHES DO ERRO', 'CATEGORIA DO ERRO']] = ['Número de VOO não informado/Inválido', 'Número de VOO não preenchido ou inválido', 'Qualidade dos Dados']

# Centro de custo não informado
base_andamento.loc[base_andamento['MENSAGEM'].str.contains('Centro de custo não informado!', case=False), 
                             ['MOTIVO DO ERRO', 'DETALHES DO ERRO', 'CATEGORIA DO ERRO']] = ['Centro de custo não informado', 'Campo centro de custo não preenchido', 'Qualidade dos Dados']

# Número de caracter excedido
base_andamento.loc[base_andamento['MENSAGEM'].str.contains('O tamanho máximo do campo "Código" é 60 caracteres.', case=False), 
                             ['MOTIVO DO ERRO', 'DETALHES DO ERRO', 'CATEGORIA DO ERRO']] = ['Número de caracter excedido', 'Campo preenchido com mais de 60 caracteres', 'Qualidade dos Dados']

# Canal de venda não informado
base_andamento.loc[base_andamento['MENSAGEM'].str.contains('Canal de venda com descrição', case=False), 
                             ['MOTIVO DO ERRO', 'DETALHES DO ERRO', 'CATEGORIA DO ERRO']] = ['Canal de venda não encontrado', 'Campo canal de venda não encontrado', 'Qualidade dos Dados']

# Erro de XML e/ou SQL
base_andamento.loc[base_andamento['MENSAGEM'].str.contains('The UPDATE statement conflicted with the FOREIGN KEY constraint', case=False) |
                                base_andamento['MENSAGEM'].str.contains('The DELETE statement conflicted with the REFERENCE constraint', case=False) |
                                base_andamento['MENSAGEM'].str.contains('is specified more than once in the SET clause or column list of an INSERT', case=False) |
                                base_andamento['MENSAGEM'].str.contains('Inserindo Accounting - (ExecSQL) - List index out of bounds (0)', case=False),
                             ['MOTIVO DO ERRO', 'DETALHES DO ERRO', 'CATEGORIA DO ERRO']] = ['Erro de XML e/ou SQL', 'Erro de processamento da reserva', 'Sistêmico']


# Garantir colunas extras no base_novo
colunas_adicionais = ['MOTIVO DO ERRO', 'DETALHES DO ERRO', 'CATEGORIA DO ERRO', 'DATA INCLUSÃO']
for col in colunas_adicionais:
    if col not in base_novo.columns:
        base_novo[col] = pd.NA

base_novo['DATA INCLUSÃO'] = data_hoje

# Classificar status
base_novo['STATUS'] = 'Novo'
base_novo.loc[base_novo['HANDLE'].isin(base_andamento['HANDLE']), 'STATUS'] = 'Em Andamento'
base_novo.loc[base_novo['HANDLE'].isin(base_resolvidos['HANDLE']), 'STATUS'] = 'Resolvido'

# Identificar registros novos
novos = base_novo[base_novo['STATUS'] == 'Novo'].copy()
# Atualizar status e mover para andamento
novos['STATUS'] = 'Em Andamento'
base_andamento = pd.concat([base_andamento, novos], ignore_index=True)

# Atualizar status de andamento com a base_novo
status_temp = base_novo[['HANDLE', 'STATUS']]
base_andamento = base_andamento.merge(status_temp, on='HANDLE', how='left', suffixes=('', '_NOVO'))
base_andamento['STATUS'] = base_andamento['STATUS_NOVO'].fillna('Resolvido')
base_andamento.drop(columns=['STATUS_NOVO'], inplace=True)

# Identificar resolvidos
resolvidos = base_andamento[base_andamento['STATUS'] == 'Resolvido'].copy()
resolvidos['DATA CONLCUSÃO'] = data_hoje

# Adiciona aos resolvidos
base_resolvidos = pd.concat([base_resolvidos, resolvidos], ignore_index=True)
base_resolvidos['DATA CONLCUSÃO'] = pd.to_datetime(base_resolvidos['DATA CONLCUSÃO']).dt.date

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