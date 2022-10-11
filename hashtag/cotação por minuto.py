import pandas as pd
from datetime import datetime
import requests
import time

for i in range(10):
    requisicao = requests.get(
        'https://economia.awesomeapi.com.br/all/USD-BRL,EUR-BRL,BTC-BRL')
    requisicao_dic = requisicao.json()

    cotacao_dolar = float(requisicao_dic['USD']['bid'])
    cotacao_euro = float(requisicao_dic['EUR']['bid'])
    cotacao_bitcoin = float(requisicao_dic['BTC']['bid'])

    tabela = pd.read_excel('data/Cotações.xlsx')
    tabela.loc[0, 'Cotação'] = (cotacao_dolar)
    tabela.loc[1, 'Cotação'] = (cotacao_euro)
    tabela.loc[2, 'Cotação'] = (cotacao_bitcoin) * 1000

    tabela.loc[0, 'Última atualização'] = datetime.now()

    tabela.to_excel('data/Cotações.xlsx', index=False)
    print(
        f'Cotações atualizadas com sucesso! {datetime.strftime(datetime.now(), "%d/%m/%Y - %H:%M:%S")}')
    time.sleep(60*60)
