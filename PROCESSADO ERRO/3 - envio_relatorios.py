import os
import win32com.client as win32
import pandas as pd
from datetime import datetime

# Importando os arquivos
relatorio_erro = pd.read_excel('data-analysis-python/PROCESSADO ERRO/Relatorio - Dash.xlsx', sheet_name='Processado Erro - BASE')
relatorio_erro_zupper = pd.read_excel('data-analysis-python/PROCESSADO ERRO/EMPRESAS/Relatorio - ZUPPER VIAGENS.xlsx')
relatorio_erro_corp = pd.read_excel('data-analysis-python/PROCESSADO ERRO/EMPRESAS/Relatorio - KONTIK BUSINESS TRAVEL.xlsx')
relatorio_erro_kontrip = pd.read_excel('data-analysis-python/PROCESSADO ERRO/EMPRESAS/Relatorio - KONTRIP VIAGENS.xlsx')
relatorio_erro_inovents = pd.read_excel('data-analysis-python/PROCESSADO ERRO/EMPRESAS/Relatorio - INOVENTS.xlsx')
relatorio_erro_grpktk = pd.read_excel('data-analysis-python/PROCESSADO ERRO/EMPRESAS/Relatorio - GRUPO KONTIK.xlsx')

novo_arquivo_resolvido = pd.read_excel('data-analysis-python/PROCESSADO ERRO/Base.xlsx', sheet_name='Novo Arquivo')
base_resolvido = pd.read_excel('data-analysis-python/PROCESSADO ERRO/Base.xlsx', sheet_name='Resolvidos')

# Emails
emails_corp = {
    'envio': [
        'wagneyoliveira@kontik.com.br','yurirodrigues@kontik.com.br',
        'wellingtonribeiro@kontik.com.br','michellysilva@kontik.com.br','eduardomanso@kontik.com.br',
        'vanessadias@kontik.com.br','giselecarmo@kontik.com.br','nucleonabr@kontik.com.br','cartaoaereo@kontik.com.br','jackelinenascimento@kontik.com.br',
        'andreajorge@kontik.com.br','adailtonsantos@kontik.com.br','reinildosantos@kontik.com.br',
        'andreiaalves@kontik.com.br','herbertsantana@kontik.com.br','camilasilva@kontik.com.br','robertobento@kontik.com.br','jacquelinesantos@kontik.com.br',
        'anafeitosa@kontik.com.br','mylenasilva@kontik.com.br','samsung@kontik.com.br','giseledenck@kontik.com.br','leticiapinheiro@kontik.com.br','andressasilva@kontik.com.br'
    ],    
    'copia': [
        'alexandrecastro@kontik.com.br','lanatakuma@kontik.com.br','thiagobatello@kontik.com.br','danielacoelho@kontik.com.br',
        'rafaelzizzi@kontik.com.br','luisvasquez@kontik.com.br','pliniocarvalho@kontik.com.br'

        ]}

#Empresa - ZUPPER VIAGENS
emails_zupper = {
    'envio': [
        'higorlima@zupper.com.br'],
    'copia': [
        'angelasilva@zupper.com.br','pliniocarvalho@kontik.com.br']}

# Empresa - KONTRIP VIAGENS
emails_kontrip = {
    'envio': [
        'laylaoliveira@kontrip.com.br', 'emillysantos@kontrip.com.br'],
    'copia': [
        'alexandreberbel@kontrip.com.br','pliniocarvalho@kontik.com.br']}

# Empresa - GRUPO KONTIK
emails_grpkontik = {
    'envio' : ['mylenasilva@kontik.com.br','icaroxavier@kontik.com.br',
            'conciliacao_aereo@kontik.com.br','suporte.benner@kontik.com.br','thiagobatello@kontik.com.br',
            'victorbazogli@kontik.com.br','wellingtonribeiro@kontik.com.br'],
    'copia': ['biancasantos@kontik.com.br','luisvasquez@kontik.com.br',
              'pliniocarvalho@kontik.com.br', 'williancardoso@kontik.com.br']}

# Empresa - KTK
emails_ktk = {
    'envio' : ['mariatrindade@kontik.com.br'],
    'copia': ['girlacarneiro@kontik.com.br','pliniocarvalho@kontik.com.br']}

# Empresa - INOVENTS
emails_inovents = {
    'envio' : ['mariatrindade@kontik.com.br','flaviomazzola@inovents.com.br'],
    'copia': ['alexandrecastro@kontik.com.br','administrativo@inovents.com.br','lucianagarcez@inovents.com.br','pliniocarvalho@kontik.com.br']}


def geracao_email(relatorio=relatorio_erro, empresa='GRUPO KONTIK', email_envio=emails_grpkontik['envio'], email_copia=emails_grpkontik['copia']):

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
    link_bi = 'https://app.powerbi.com/view?r=eyJrIjoiMzM2NDgzZDEtNmE3Yi00MTMxLWEzNTMtYjM5NTUxYTcwZTMwIiwidCI6IjcwZGU1YWJlLTk2YzgtNDU2MS05Nzg0LThhYWQ1NTBlZDI2MCJ9'
    link_sd = 'https://servicedesk.kontikspo.com.br/WorkOrder.do?woMode=newWO&reqTemplate=1505'

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
    dashboard_pdf = r'C:\Users\igorsantos\Desktop\DOCS\data-analysis-python\PROCESSADO ERRO\PDFs\Relatorio - ' + \
        f'{datetime.now().strftime("%d.%m.%Y")}.pdf'
    
    quero_passagem = r'C:\Users\igorsantos\Desktop\DOCS\data-analysis-python\PROCESSADO ERRO\Quero Passagem.xlsx'
    
    # integra_tour = r'C:\Users\igorsantos\Desktop\DOCS\data-analysis-python\PROCESSADO ERRO\Contabilização Manual - IntegraTur.xlsx'       
    relatorio = r'C:\Users\igorsantos\Desktop\DOCS\data-analysis-python\PROCESSADO ERRO\EMPRESAS\Relatorio - ' + \
    f'{empresa}.xlsx'

    relatorio_dash = r'C:\Users\igorsantos\Desktop\DOCS\data-analysis-python\PROCESSADO ERRO\Relatorio - Dash.xlsx'

    email.Attachments.Add(relatorio)

    # email.Attachments.Add(dashboard_pdf)
    if empresa == 'GRUPO KONTIK' or empresa == 'KONTIK BUSINESS TRAVEL':
        email.Attachments.Add(quero_passagem)
        email.Attachments.Add(dashboard_pdf)
        # email.Attachments.Add(integra_tour)
    
    if empresa == 'GRUPO KONTIK':
        email.Attachments.Remove(1)
        email.Attachments.Add(relatorio_dash)

    email.SentOnBehalfOfName = remetente
    email.Save()

    if total_casos == 0:
        print(f'\033[1;31m- Não há casos para envio do email para {empresa}!\033[m')
        email.Delete()
    print(f'E-mail da empresa {empresa} criado com sucesso!')

geracao_email()
geracao_email(relatorio_erro_zupper, 'ZUPPER VIAGENS', emails_zupper['envio'], emails_zupper['copia'])
geracao_email(relatorio_erro_corp, 'KONTIK BUSINESS TRAVEL', emails_corp['envio'],emails_corp['copia'])
geracao_email(relatorio_erro_kontrip,'KONTRIP VIAGENS', emails_kontrip['envio'], emails_kontrip['copia'])
geracao_email(relatorio_erro_inovents, 'INOVENTS', emails_inovents['envio'],emails_inovents['copia'])
# geracao_email(relatorio_erro_grpktk, 'GRUPO KONTIK', emails_grpkontik['envio'], emails_grpkontik['copia'])

print('\033[1;32m-Emails criado com sucesso!\033[m')