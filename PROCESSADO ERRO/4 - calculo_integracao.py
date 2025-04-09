import pandas as pd
from datetime import datetime

# Caminho dos arquivos
data_path = 'data-analysis-python/PROCESSADO ERRO/'

# Leitura dos arquivos
base_xls = pd.ExcelFile(data_path + 'Base.xlsx')
base_resolvidos = base_xls.parse(sheet_name='Resolvidos')

# Converte a coluna de data para datetime (caso ainda não esteja)
base_resolvidos['Data Inclusão'] = pd.to_datetime(base_resolvidos['Data Inclusão'])

# Pega o ano e mês atuais
hoje = datetime.today()
ano_atual = hoje.year
mes_atual = hoje.month

# Filtra casos do SABRE no mês atual
casos_sabre_carro = base_resolvidos[
    (base_resolvidos['OBTS'] == 'SABRE') &
    (base_resolvidos['Serviço'] == 'Carro') &
    (base_resolvidos['Data Inclusão'].dt.year == ano_atual) &
    (base_resolvidos['Data Inclusão'].dt.month == mes_atual)
]

casos_sabre_hotel = base_resolvidos[
    (base_resolvidos['OBTS'] == 'SABRE') &
    (base_resolvidos['Serviço'] == 'Hotel') &
    (base_resolvidos['Data Inclusão'].dt.year == ano_atual) &
    (base_resolvidos['Data Inclusão'].dt.month == mes_atual)
]

# Conta quantos casos trouxe
quantidade_casos_sabre_mes = casos_sabre_carro.shape[0]
quantidade_casos_sabre_hotel = casos_sabre_hotel.shape[0]
print(f"Quantidade de casos SABRE no mês: {quantidade_casos_sabre_mes}")
print(f"Quantidade de casos SABRE Hotel no mês: {quantidade_casos_sabre_hotel}")
