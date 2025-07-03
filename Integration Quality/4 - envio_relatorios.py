import os
import win32com.client as win32
import pandas as pd
from datetime import datetime

# Obtendo o nome do usuário atual
usuario = os.getlogin()
# Importando os arquivos
novo_arquivo_resolvido = pd.read_excel('data-analysis-python/Integration Quality/Base.xlsx', sheet_name='Novo Arquivo')
base_resolvido = pd.read_excel('data-analysis-python/Integration Quality/Base.xlsx', sheet_name='Resolvidos')
integra_tour_base = pd.read_excel('data-analysis-python/Integration Quality/Relatorio - Integratour.xlsx', sheet_name='Integrado Erro')
caminho_dashboard = f'C:\\Users\\{usuario}\\Desktop\\DOCS\\data-analysis-python\\Integration Quality\\Relatorio - Dash.xlsx'

data_hoje = datetime.now().strftime('%d.%m.%Y')
nome_pdf = f'Relatorio - {data_hoje}.pdf'
caminho_saida_pdf = os.path.join(f'C:\\Users\\{usuario}\\Desktop\\DOCS\\data-analysis-python\\Integration Quality\\PDFs', nome_pdf)

# Inicia o Excel
excel = win32.Dispatch('Excel.Application')
excel.Visible = False  # Mantenha oculto durante execução

try:
    # Abre o workbook
    wb = excel.Workbooks.Open(caminho_dashboard)
    
    # Acessa a aba "Dashboard"
    aba_dashboard = wb.Sheets("Dashboard")
    # Atualiza todas as tabelas dinâmicas da aba "Dashboard"
    for pt in aba_dashboard.PivotTables():
        pt.RefreshTable()
    
    # Salva e fecha
    wb.Save()
    # Exporta a aba "Dashboard" como PDF
    aba_dashboard.ExportAsFixedFormat(0, caminho_saida_pdf)
    wb.Close()
    print('\033[1;36m- Guia "Dashboard" atualizada e exportada com sucesso!\033[m')

except Exception as e:
    print(f'\033[1;31mErro ao atualizar a guia Dashboard: {e}\033[m')

finally:
    excel.Quit()


# Empresa - KONTIK BUSINESS TRAVEL
emails_corp = {
    'envio': [
        # Lista de emails para envio - KONTIK BUSINESS TRAVEL
    ],    
    'copia': [
        # Lista de emails para cópia - KONTIK BUSINESS TRAVEL
    ]}

#Empresa - ZUPPER VIAGENS
emails_zupper = {
    'envio': [
        # Lista de emails para envio - ZUPPER VIAGENS
    ],
    'copia': [
        # Lista de emails para cópia - ZUPPER VIAGENS
        ]}

# Empresa - KONTRIP VIAGENS
emails_kontrip = {
    'envio': [
        # Lista de emails para envio - KONTRIP VIAGENS
        ],
    'copia': [
        # Lista de emails para cópia - KONTRIP VIAGENS
        ]}

# Empresa - GRUPO KONTIK
emails_grpkontik = {
    'envio' : [
        # Lista de emails para envio - GRUPO KONTIK
            ],
    'copia': [
        # Lista de emails para cópia - GRUPO KONTIK
              ]}

# Empresa - KTK
emails_ktk = {
    'envio' : [
        # Lista de emails para envio - KTK
        ],
    'copia': [
        # Lista de emails para cópia - KTK
        ]}

# Empresa - INOVENTS
emails_inovents = {
    'envio' : [
        # Lista de emails para envio - INOVENTS
        ],
    'copia': [
        # Lista de emails para cópia - INOVENTS
        ]}


def geracao_email(empresa='GRUPO KONTIK', email_envio=emails_grpkontik['envio'], email_copia=emails_grpkontik['copia'], relatorio=None):

    if empresa == 'ZUPPER VIAGENS': 
        caminho = 'data-analysis-python/Integration Quality/EMPRESAS/Relatorio - ZUPPER VIAGENS.xlsx'
    elif empresa == 'KONTIK BUSINESS TRAVEL':
        caminho = 'data-analysis-python/Integration Quality/EMPRESAS/Relatorio - KONTIK BUSINESS TRAVEL.xlsx'
    elif empresa == 'KONTRIP VIAGENS':
        caminho = 'data-analysis-python/Integration Quality/EMPRESAS/Relatorio - KONTRIP VIAGENS.xlsx'
    elif empresa == 'INOVENTS':
        caminho = 'data-analysis-python/Integration Quality/EMPRESAS/Relatorio - INOVENTS.xlsx'
    elif empresa == 'GRUPO KONTIK':
        caminho = 'data-analysis-python/Integration Quality/EMPRESAS/Relatorio - GRUPO KONTIK.xlsx'
    else:
        print(f'\033[1;31m- Empresa {empresa} não encontrada!\033[m')
        return

    if not os.path.exists(caminho):
        print(f'\033[1;33m- Arquivo não encontrado para {empresa}\033[m')
        return

    # Se chegou aqui, o arquivo existe
    if empresa ==  'GRUPO KONTIK':
        relatorio= pd.read_excel('data-analysis-python/Integration Quality/Relatorio - Dash.xlsx', sheet_name='Processado Erro - BASE')
    else:
        relatorio = pd.read_excel(caminho)
    
    # Total de casos - Processado Erro
    total_casos = len(relatorio)
    
    # Top 5 Grupos Empresariais
    top_5_grp_emp = ', '.join(relatorio['Grupo Empresarial'].value_counts().head(5).index)

    # Aging Acima de 15 Dias
    soma_aging_alteracao = len(relatorio.loc[relatorio['Aging Alteração'].str.contains(
        '16 a 23 dias|24 a 31 dias|31 dias ou +')])
    
    soma_aging_inclusao = len(relatorio.loc[relatorio['Aging Inclusão'].str.contains(
        '16 a 23 dias|24 a 31 dias|31 dias ou +')
        ])
    
    # integratour
    integratour_hoje = integra_tour_base.loc[integra_tour_base['DATAENVIO'] == datetime.now().strftime('%d/%m/%Y')]
    integratour_ontem = integra_tour_base.loc[integra_tour_base['DATAENVIO'] == (datetime.now() - pd.Timedelta(days=1)).strftime('%d/%m/%Y')]
    print(f'\033[1;36m- Integratour ontem: {len(integratour_ontem)}\033[m')
    
    # Casos que retornaram
    handles_resolvidos = novo_arquivo_resolvido.loc[novo_arquivo_resolvido['Status'] == 'Resolvido', 'Handle PNR'].tolist()

    casos_retornados = {
        'ANALISE': [],
        'ZUPPER VIAGENS': [],
        'KONTRIP VIAGENS': [],
        'INOVENTS': [],
        'GRUPO KONTIK': [],
        'KONTIK BUSINESS TRAVEL': []
    }

    # Corrigindo o loop para adicionar casos retornados
    
    for row in range(len(relatorio)):
        for handle in handles_resolvidos:
            if str(handle) in str(relatorio['Handle PNR'][row]):
                casos_retornados[empresa].append(relatorio['Localizadora'][row])
    
    casos_retornados[empresa] = list(set(casos_retornados[empresa]))
    casos_formatados = casos_retornados[empresa]

    # porcentagem categoria de erro qualidade de dados
    porcentagem_qualidade_dados = (relatorio['CATEGORIA DE ERRO'] == 'Qualidade dos dados').sum() / total_casos * 100
    porcentagem_sistemico = (relatorio['CATEGORIA DE ERRO'] == 'Sistêmico').sum() / total_casos * 100
    porcentagem_operacional = (relatorio['CATEGORIA DE ERRO'] == 'Processo Operacional').sum() / total_casos * 100

    # Maiores Ofensores do Relatório - categoria
    maior_ofensor = relatorio['CAMPO'].value_counts().head(1).index.tolist()[0]

    # Maiores Ofensores do Relatório - quantidade
    qtd_maior_ofensor = relatorio['CAMPO'].value_counts().head(1).values[0]

    # Maiores Ofensores do Relatório - OBT
    obt_maior_ofensor = relatorio['OBTS'].value_counts().head(1).index.tolist()[0]

    # Maiores Ofensores do Relatório - quantidade
    qtd_obt_maior_ofensor = ((relatorio['OBTS'] == obt_maior_ofensor) & (relatorio['CAMPO'] == maior_ofensor)).sum()

    if empresa == 'KONTIK BUSINESS TRAVEL' or empresa == 'GRUPO KONTIK':
        # maior ofensor por obt
        maior_ofensor_argo = relatorio.loc[relatorio['OBTS'] == 'ARGO(TMS)', 'CAMPO'].value_counts().head(1).index.tolist()[0]
        qtd_maior_ofensor_argo = relatorio.loc[relatorio['OBTS'] == 'ARGO(TMS)', 'CAMPO'].value_counts().head(1).values[0]
        porcentagem_maior_ofensor_argo = (qtd_maior_ofensor_argo / total_casos) * 100

        maior_ofensor_sabre = relatorio.loc[relatorio['OBTS'] == 'SABRE', 'CAMPO'].value_counts().head(1).index.tolist()[0]
        qtd_maior_ofensor_sabre = relatorio.loc[relatorio['OBTS'] == 'SABRE', 'CAMPO'].value_counts().head(1).values[0]
        porcentagem_maior_ofensor_sabre = (qtd_maior_ofensor_sabre / total_casos) * 100

        maior_ofensor_gover = relatorio.loc[relatorio['OBTS'] == 'GOVER', 'CAMPO'].value_counts().head(1).index.tolist()[0]
        qtd_maior_ofensor_gover = relatorio.loc[relatorio['OBTS'] == 'GOVER', 'CAMPO'].value_counts().head(1).values[0]
        porcentagem_maior_ofensor_gover = (qtd_maior_ofensor_gover / total_casos) * 100

        try:
            maior_ofensor_lemontech = relatorio.loc[relatorio['OBTS'] == 'LEMONTECH', 'CAMPO'].value_counts().head(1).index.tolist()[0]
            qtd_maior_ofensor_lemontech = relatorio.loc[relatorio['OBTS'] == 'LEMONTECH', 'CAMPO'].value_counts().head(1).values[0]
            porcentagem_maior_ofensor_lemontech = (qtd_maior_ofensor_lemontech / total_casos) * 100
        except (IndexError, ZeroDivisionError):
            maior_ofensor_lemontech = '-'
            qtd_maior_ofensor_lemontech = 0
            porcentagem_maior_ofensor_lemontech = 0

        primeiro_ofensor ={'obt': '-','campo': '-','qtd': '-','porcentagem': '-'}
        segundo_ofensor = {'obt': '-','campo': '-','qtd': '-','porcentagem': '-'}
        terceiro_ofensor = {'obt': '-','campo': '-','qtd': '-','porcentagem': '-'}
        quarto_ofensor = {'obt': '-','campo': '-','qtd': '-','porcentagem': '-'}

        # Ranqueando os ofensores por OBT
        ofensores_por_obt = {
            'ARGO(TMS)': porcentagem_maior_ofensor_argo,
            'SABRE': porcentagem_maior_ofensor_sabre,
            'GOVER': porcentagem_maior_ofensor_gover,
            'LEMONTECH': porcentagem_maior_ofensor_lemontech
        }

        # Ordenando os ofensores por OBT do maior para o menor
        ofensores_ordenados = sorted(ofensores_por_obt.items(), key=lambda x: x[1], reverse=True)

        # Preenchendo os dados dos ofensores no ranking
        for i, (obt, porcentagem) in enumerate(ofensores_ordenados[:4]):
            if i == 0:
                primeiro_ofensor['obt'] = obt
                primeiro_ofensor['porcentagem'] = porcentagem
            elif i == 1:
                segundo_ofensor['obt'] = obt
                segundo_ofensor['porcentagem'] = porcentagem
            elif i == 2:
                terceiro_ofensor['obt'] = obt
                terceiro_ofensor['porcentagem'] = porcentagem
            elif i == 3:
                quarto_ofensor['obt'] = obt
                quarto_ofensor['porcentagem'] = porcentagem

        # Preenchendo os campos de quantidade e campo para os ofensores
        for obt, campo_ofensor in [(primeiro_ofensor['obt'], primeiro_ofensor), 
                                (segundo_ofensor['obt'], segundo_ofensor), 
                                (terceiro_ofensor['obt'], terceiro_ofensor), 
                                (quarto_ofensor['obt'], quarto_ofensor)]:
            try:                    
                campo_ofensor['campo'] = relatorio.loc[relatorio['OBTS'] == obt, 'CAMPO'].value_counts().head(1).index[0]
                campo_ofensor['qtd'] = relatorio.loc[(relatorio['OBTS'] == obt) & (relatorio['CAMPO'] == campo_ofensor['campo'])].shape[0]
            except (IndexError):
                campo_ofensor['campo'] = ''
                campo_ofensor[ 'qtd'] = 0
    #================================================================================================

    # Criando o email
    outlook = win32.Dispatch('outlook.application')

    # Remetente
    remetente = 'suporte.benner@kontik.com.br'

    # 
    email = outlook.CreateItem(0)

    # Configurações do email
    email.to = ';'.join(email_envio)
    email.cc = ';'.join(email_copia)

    # Assunto
    email.Subject = f'📊 Análise Diária - Qualidade de Integração | {datetime.now().strftime("%d/%m/%Y")} | {empresa}' 

    # Links
    link_sd = 'Inserir link do Service Desk aqui'
    link_bi = 'Inserir link do Power BI aqui'

    # Corpo do email 1
    corpo_email_1 = f"""
    <style>
        p,ul,li {{
            font-size: 11pt;
        }}
    </style>
    <p>Bom dia, pessoal!</p>

    <p>Segue abaixo a análise detalhada do <strong>Processado Erro</strong>, com base no arquivo recebido hoje.</p>
    
    <p><strong>📌 Para solicitações ao Suporte Benner, é imprescindível a abertura de chamado via Service Desk <a href="{link_sd}" style="color: #007bff;">aqui</a> ou no caminho:
    <br>➡️ Portal Benner → Contabilização → Pendentes (Processado Erro)</strong></p>

    <p>🔗<a href="{link_bi}" style="color: #007bff;"><strong>Clique aqui para acessar o Power Bi</strong></a></p>
    
    <p><strong>🔍 Pontos de Atenção:</strong></p>
        <ul>
            <li> <strong>Grupos empresariais que mais impactam:</strong> {top_5_grp_emp}</li>    
            <li> <strong>Aging Alteração acima de 15 Dias:</strong> {soma_aging_alteracao} casos, indicando a necessidade de atenção especial</li>
            <li> <strong>Aging Inclusão acima de 15 Dias:</strong> {soma_aging_inclusao} casos, indicando a necessidade de atenção especial</li>    
            <li> <strong>Casos que retornaram:</strong> Identificamos {len(casos_retornados[empresa])}: {casos_formatados}</li>
            <li> <strong>Porcentagem de Erros:</strong> 
                <ul>    
                    <li>{porcentagem_qualidade_dados:.2f}% – Qualidade dos Dados</li>
                    <li>{porcentagem_sistemico:.2f}% – Sistêmico</li>
                    <li>{porcentagem_operacional:.2f}% – Processo Operacional</li>
                </ul>
            </li>
    """

    if empresa == 'KONTIK BUSINESS TRAVEL' or empresa == 'GRUPO KONTIK':

        corpo_email_2 = f"""

            <li> <strong>Relatório do Quero Passagem:</strong> responsabilidade (coluna A) e Tipo de Erro (coluna B), sendo: 
                <ul>
                    <li>Fornecedor: alterar para o CNPJ da viação (já contabilizado, apenas ajustar o fornecedor)</li>      
                    <li>Não contabilizada:</strong> seguir com a contabilização manual.</li>
                </ul>
            </li>
        </ul>
    
    <p><strong>🔥 Maiores Ofensores por OBT:</strong></p>    
        <ul>
            <li><strong>{primeiro_ofensor['obt']}:</strong> {primeiro_ofensor['qtd']} casos de {primeiro_ofensor['campo']} sendo {primeiro_ofensor['porcentagem']:.2f}% do total de casos</li>           
            <li><strong>{segundo_ofensor['obt']}:</strong> {segundo_ofensor['qtd']} casos de {segundo_ofensor['campo']} sendo {segundo_ofensor['porcentagem']:.2f}% do total de casos</li>            
            <li><strong>{terceiro_ofensor['obt']}:</strong> {terceiro_ofensor['qtd']} casos de {terceiro_ofensor['campo']} sendo {terceiro_ofensor['porcentagem']:.2f}% do total de casos</li>               
            <li><strong>{quarto_ofensor['obt']}:</strong> {quarto_ofensor['qtd']} casos de {quarto_ofensor['campo']} sendo {quarto_ofensor['porcentagem']:.2f}% do total de casos</li>
        
    """
    
    corpo_email_3 = """
    </ul>  
    <p><strong>✅ Ações Recomendadas:</strong></p>
        <ul>
            <li><strong>Priorização:</strong> Foco na resolução dos casos relacionados aos grupos empresariais com maior impacto.</li>
            <li><strong>Aging > 15 dias: </strong> Monitorar com atenção os aging mais antigos para evitar atrasos no processo.</li>   
            <li><strong>Casos Recorrentes:</strong> Investigar a fundo quaisquer casos reincidentes para evitar novas ocorrências.</li>    
            <li><strong>Power BI:</strong> Utilize o painel para análises visuais complementares e tomada de decisão.</li>    
        </ul>
    <br>
    <p><strong>📣 Lembrete importante:</strong><br> Em caso de dúvidas, dificuldades ou necessidade de apoio, <strong>abra um chamado conforme instruções acima</strong>. 
        Isso garante um atendimento ágil e rastreável por parte da equipe de suporte.</p>
    <p>Ficamos à disposição para quaisquer esclarecimentos ou ações adicionais necessárias.</p>
    <br>
    <br>
    """

    # Assinatura do email
    assinatura_nome = "Igor F. Santos (igorsantos@kontik.com.br)"

    assinatura_path = os.path.expandvars(rf"%APPDATA%\Microsoft\Signatures\{assinatura_nome}.htm")
    
    with open(assinatura_path, 'r', encoding='latin-1') as f:
        assinatura_html = f.read()

    # Definição de corpo dos e-mails
    if empresa == 'KONTIK BUSINESS TRAVEL' or empresa == 'GRUPO KONTIK':
        email.HTMLBody = corpo_email_1 + corpo_email_2 + corpo_email_3 + assinatura_html
    else:
        email.HTMLBody = corpo_email_1 + corpo_email_3 + assinatura_html

    # Anexos
    dashboard_pdf = f'C:\\Users\\{usuario}\\Desktop\\DOCS\\data-analysis-python\\Integration Quality\\PDFs\\Relatorio - {datetime.now().strftime("%d.%m.%Y")}.pdf'

    quero_passagem = f'C:\\Users\\{usuario}\\Desktop\\DOCS\\data-analysis-python\\Integration Quality\\Quero Passagem.xlsx'

    # integra_tour = f'C:\\Users\\{usuario}\\Desktop\\DOCS\\data-analysis-python\\Integration Quality\\Contabilização Manual - IntegraTur.xlsx'
    relatorio = f'C:\\Users\\{usuario}\\Desktop\\DOCS\\data-analysis-python\\Integration Quality\\EMPRESAS\\Relatorio - {empresa}.xlsx'

    relatorio_dash = f'C:\\Users\\{usuario}\\Desktop\\DOCS\\data-analysis-python\\Integration Quality\\Relatorio - Dash.xlsx'


    if empresa == 'GRUPO KONTIK' or empresa == 'KONTIK BUSINESS TRAVEL':
        # email.Attachments.Add(quero_passagem)
        email.Attachments.Add(dashboard_pdf)
        # email.Attachments.Add(integra_tour)
    
    if empresa == 'GRUPO KONTIK':
        email.Attachments.Remove(1)
        email.Attachments.Add(relatorio_dash)
    else:
        email.Attachments.Add(relatorio)

    email.SentOnBehalfOfName = remetente
    email.Save()

    if total_casos == 0:
        print(f'\033[1;31m- Não há casos para envio do email para {empresa}!\033[m')
        email.Delete()
    print(f'E-mail da empresa {empresa} criado com sucesso!')




geracao_email()
geracao_email('ZUPPER VIAGENS', emails_zupper['envio'], emails_zupper['copia'])
geracao_email('KONTIK BUSINESS TRAVEL', emails_corp['envio'],emails_corp['copia'])
geracao_email('KONTRIP VIAGENS', emails_kontrip['envio'], emails_kontrip['copia'])
geracao_email('INOVENTS', emails_inovents['envio'],emails_inovents['copia'])

print('\033[1;32m-Emails criado com sucesso!\033[m')