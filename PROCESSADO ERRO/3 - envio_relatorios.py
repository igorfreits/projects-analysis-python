import win32com.client as win32
import pandas as pd
from datetime import datetime

data_path = 'data-analysis-python/PROCESSADO ERRO/'

# Importando os arquivos
relatorio_erro_zupper = pd.read_excel(data_path + 'EMPRESAS/Relatorio - ZUPPER VIAGENS.xlsx')
relatorio_erro_corp = pd.read_excel(data_path + 'EMPRESAS/Relatorio - KONTIK BUSINESS TRAVEL.xlsx')
relatorio_erro_kontrip = pd.read_excel(data_path + 'EMPRESAS/Relatorio - KONTRIP VIAGENS.xlsx')
#relatorio_erro_inovents = pd.read_excel(data_path + 'EMPRESAS/Relatorio - INOVENTS.xlsx')
relatorio_erro_grpktk = pd.read_excel(data_path + 'EMPRESAS/Relatorio - GRUPO KONTIK.xlsx')

novo_arquivo_resolvido = pd.read_excel(data_path + 'Base.xlsx', sheet_name='Novo Arquivo')
base_resolvido = pd.read_excel(data_path + 'Base.xlsx', sheet_name='Resolvidos')
relatorio_erro = pd.read_excel(data_path + 'Relatorio - Dash.xlsx', sheet_name='Processado Erro - BASE')


# Emails
emails_corp = {
    'envio': [
    'robertasilva@kontik.com.br','brunomoreira@kontik.com.br','camilasilva@kontik.com.br','elaineoliveira@kontik.com.br',
    'marcelovieira@kontik.com.br','najlavieira@kontik.com.br','rodrigoberton@kontik.com.br','ceciliaherculino@kontik.com.br',
    'giovannapereira@kontik.com.br','henriquetolentino@kontik.com.br','karinaxavier@kontik.com.br',
    'lucieneleal@kontik.com.br','margaretokura@kontik.com.br','michellysilva@kontik.com.br','nayaraoliveira@kontik.com.br',
    'wagneyoliveira@kontik.com.br','yurirodrigues@kontik.com.br','reinildosantos@kontik.com.br',
    'nucleonabr@kontik.com.br','taisdiniz@kontik.com.br','alicasantos@kontik.com.br','joaovillar@kontik.com.br','pamelasilva@kontik.com.br'],    
    'copia': [
        'faturamentocliente@kontik.com.br','jackelinenascimento@kontik.com.br','rafaelzizzi@kontik.com.br','vanessadias@kontik.com.br',
        'adailtonsantos@kontik.com.br','lanatakuma@kontik.com.br','alexandrecastro@kontik.com.br','cartaoaereo@kontik.com.br',
        'andreajorge@kontik.com.br','raquelmonteiro@kontik.com.br','elidealtran@kontik.com.br','conciliacao_aereo@kontik.com.br','pliniocarvalho@kontik.com.br'
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
    'copia': ['alexandrecastro@kontik.com.br','josysilva@inovents.com.br','administrativo@inovents.com.br','lucianagarcez@inovents.com.br','pliniocarvalho@kontik.com.br']}


def geracao_email(relatorio_empresa=relatorio_erro, empresa='GRUPO KONTIK', email_envio=emails_grpkontik['envio'], email_copia=emails_grpkontik['copia']):

    # Total de casos - Processado Erro
    total_casos = len(relatorio_empresa)
    
    # Top 5 Grupos Empresariais
    top_5_grp_emp = ', '.join(relatorio_empresa['Grupo Empresarial'].value_counts().head(5).index)

    # Aging Acima de 15 Dias
    soma_aging_alteracao = len(relatorio_empresa.loc[relatorio_empresa['Aging Alteração'].str.contains(
        '16 a 23 dias|24 a 31 dias|31 dias ou +')])
    
    soma_aging_inclusao = len(relatorio_empresa.loc[relatorio_empresa['Aging Inclusão'].str.contains(
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
    for row in range(len(relatorio_empresa)):
        for handle in handles_resolvidos:
            if str(handle) in str(relatorio_empresa['Handle PNR'][row]):
                casos_retornados[empresa].append(relatorio_empresa['Localizadora'][row])
    
    casos_retornados[empresa] = list(set(casos_retornados[empresa]))
    casos_formatados = ", ".join(casos_retornados[empresa])

    # porcentagem categoria de erro qualidade de dados
    porcentagem_qualidade_dados = (relatorio_empresa['CATEGORIA DE ERRO'] == 'Qualidade dos dados').sum() / total_casos * 100
    porcentagem_sistemico = (relatorio_empresa['CATEGORIA DE ERRO'] == 'Sistêmico').sum() / total_casos * 100

    # Maiores Ofensores do Relatório - categoria
    maior_ofensor = relatorio_empresa['CAMPO'].value_counts().head(1).index.tolist()[0]
    # Maiores Ofensores do Relatório - quantidade
    qtd_maior_ofensor = relatorio_empresa['CAMPO'].value_counts().head(1).values[0]

    # Maiores Ofensores do Relatório - OBT
    obt_maior_ofensor = relatorio_empresa['OBTS'].value_counts().head(1).index.tolist()[0]
    # Maiores Ofensores do Relatório - quantidade
    qtd_obt_maior_ofensor = ((relatorio_empresa['OBTS'] == obt_maior_ofensor) & (relatorio_empresa['CAMPO'] == maior_ofensor)).sum()

    if empresa == 'KONTIK BUSINESS TRAVEL' or empresa == 'GRUPO KONTIK':
        # maior ofensor por obt
        maior_ofensor_argo = relatorio_empresa.loc[relatorio_empresa['OBTS'] == 'ARGO(TMS)', 'CAMPO'].value_counts().head(1).index.tolist()[0]
        qtd_maior_ofensor_argo = relatorio_empresa.loc[relatorio_empresa['OBTS'] == 'ARGO(TMS)', 'CAMPO'].value_counts().head(1).values[0]
        porcentagem_maior_ofensor_argo = (qtd_maior_ofensor_argo / total_casos) * 100

        maior_ofensor_sabre = relatorio_empresa.loc[relatorio_empresa['OBTS'] == 'SABRE', 'CAMPO'].value_counts().head(1).index.tolist()[0]
        qtd_maior_ofensor_sabre = relatorio_empresa.loc[relatorio_empresa['OBTS'] == 'SABRE', 'CAMPO'].value_counts().head(1).values[0]
        porcentagem_maior_ofensor_sabre = (qtd_maior_ofensor_sabre / total_casos) * 100

        maior_ofensor_gover = relatorio_empresa.loc[relatorio_empresa['OBTS'] == 'GOVER', 'CAMPO'].value_counts().head(1).index.tolist()[0]
        qtd_maior_ofensor_gover = relatorio_empresa.loc[relatorio_empresa['OBTS'] == 'GOVER', 'CAMPO'].value_counts().head(1).values[0]
        porcentagem_maior_ofensor_gover = (qtd_maior_ofensor_gover / total_casos) * 100

        try:
            maior_ofensor_lemontech = relatorio_empresa.loc[relatorio_empresa['OBTS'] == 'LEMONTECH', 'CAMPO'].value_counts().head(1).index.tolist()[0]
            qtd_maior_ofensor_lemontech = relatorio_empresa.loc[relatorio_empresa['OBTS'] == 'LEMONTECH', 'CAMPO'].value_counts().head(1).values[0]
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
                campo_ofensor['campo'] = relatorio_empresa.loc[relatorio_empresa['OBTS'] == obt, 'CAMPO'].value_counts().head(1).index[0]
                campo_ofensor['qtd'] = relatorio_empresa.loc[(relatorio_empresa['OBTS'] == obt) & (relatorio_empresa['CAMPO'] == campo_ofensor['campo'])].shape[0]
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
    email.Subject = f'Análise de Erros - {datetime.now().strftime("%d/%m/%Y")} - {empresa}' 

    # Link do Power BI
    link = 'https://app.powerbi.com/view?r=eyJrIjoiN2I3Zjk5ZDgtMzQ3ZS00ZDcwLWJlOTgtNTA2NGI2Y2RlOGRkIiwidCI6IjcwZGU1YWJlLTk2YzgtNDU2MS05Nzg0LThhYWQ1NTBlZDI2MCJ9'

    # Corpo do email 1
    corpo_email_1 = f"""
    <p>Olá, equipe!</p>

    Espero que estejam bem, segue análise detalhada do Processado Erro com base no arquivo recebido hoje:
    <p></p>
    <br><br>
    <p><em><strong style="background-color:yellow;color:red">PARA SOLICITAÇÕES AO SUPORTE BENNER, FAVOR ABRIR CHAMADO VIA SERVICE DESK: PORTAL BENNER -> CONTABILIZAÇÃO -> PENDENTES (PROCESSADO ERRO).</strong></em></p>

    <br><strong>Link do Power BI:</strong> <a href="{link}">Clique aqui</a></br>
    <p></p>
    <strong><p><u>Pontos importantes:</u></p></strong>


    <blockquote>
        <p> <strong>1. Grupos empresariais que mais impactam:</strong> {top_5_grp_emp};</p>
    </blockquote>

    <blockquote>
        <p> <strong>2. Aging Alteração acima de 15 Dias:</strong> {soma_aging_alteracao} casos, indicando a necessidade de atenção especial;
        <p> <strong>3. Aging Inclusão acima de 15 Dias:</strong> {soma_aging_inclusao} casos, indicando a necessidade de atenção especial;
        </p>
    </blockquote>

    <blockquote>
        <p> <strong>4. Casos que retornaram:</strong> Identificamos {len(casos_retornados[empresa])}: {casos_formatados};</p>
        </p>
    </blockquote>

    <blockquote>
        <p> <strong>5. Porcentagem de Erros:</strong> {porcentagem_qualidade_dados:.2f}% de Qualidade dos Dados e {porcentagem_sistemico:.2f}% Sistêmico;
        </p>
    </blockquote>
    """

    if empresa == 'KONTIK BUSINESS TRAVEL' or empresa == 'GRUPO KONTIK':
        corpo_email_2 = f"""
    <blockquote>
        <p> <strong>6. IntegraTur:</strong> Relatório com vendas que falharam no processo automático de
            integração e precisam ser contabilizadas manualmente no Portal Wes;</p>
    </blockquote>

    <blockquote>
        <p> <strong>7. Relatório do Quero Passagem:</strong> Considere a coluna "A" para responsabilidade e a "B" para tipo
            de erro, sendo que:</p>
    </blockquote>

    <blockquote>
        <p>• Se correção, mudar o fornecedor para “Cia Rodoviaria” invés de "Quero Passagem" (Venda já contabilizada, apenas mudar o fornecedor) </p>
    </blockquote>
    <blockquote>
        <p>• Se contabilização, seguir com contabilização manual </p>
    </blockquote>
    <blockquote>
        <p>Obs.: O critério de pesquisa usado foi o campo de confirmação da Quero passagem e o rloc. </p>
    </blockquote>

    <br>
    <br>
    <br>
  
    <strong><u>Maiores Ofensores por OBT:</u></strong>

    <blockquote>
        <p> <strong>{primeiro_ofensor['obt']}:</strong> {primeiro_ofensor['qtd']} casos de {primeiro_ofensor['campo']} sendo {primeiro_ofensor['porcentagem']:.2f}% do total de casos</p>
        <p> Causa: </p>
        <p> Responsável: Suporte KCS - Operações - Suporte Benner  - Central de Emissão </p>
    </blockquote>
    <br>
    <blockquote>
        <p> <strong>{segundo_ofensor['obt']}:</strong> {segundo_ofensor['qtd']} casos de {segundo_ofensor['campo']} sendo {segundo_ofensor['porcentagem']:.2f}% do total de casos</p>
        <p> Causa: </p>
        <p> Responsável: Suporte KCS - Operações - Suporte Benner  - Central de Emissão </p>
    </blockquote>
    <br>
    <blockquote>
        <p> <strong>{terceiro_ofensor['obt']}:</strong> {terceiro_ofensor['qtd']} casos de {terceiro_ofensor['campo']} sendo {terceiro_ofensor['porcentagem']:.2f}% do total de casos</p>
        <p> Causa: </p>
        <p> Responsável: Suporte KCS - Operações - Suporte Benner  - Central de Emissão </p>
    </blockquote>
    <br>
    <blockquote>
        <p> <strong>{quarto_ofensor['obt']}:</strong> {quarto_ofensor['qtd']} casos de {quarto_ofensor['campo']} sendo {quarto_ofensor['porcentagem']:.2f}% do total de casos</p>
        <p> Causa: </p>
        <p> Responsável: Suporte KCS - Operações - Suporte Benner  - Central de Emissão </p>
    </blockquote>

    <br>
    """
    
    corpo_email_3 = """



    <br>
    <br>
    <br>
   
    <strong><u>Ações Recomendadas:</u></strong>

    <blockquote>
        <p><strong>Priorização:</strong> Recomendo priorizar a resolução dos casos nos grupos empresariais de maior impacto
            para otimizar o processo.</p>
    </blockquote>
    <blockquote>
        <p><strong>Aging Superior a 15 Dias:</strong> Uma atenção especial deve ser dada aos casos com Aging acima de 15
            dias para evitar possíveis atrasos.</p>
    </blockquote>
    <blockquote>
        <p><strong>Casos Recorrentes:</strong> Os casos que reapareceram merecem uma investigação mais aprofundada para
            evitar recorrências futuras.</p>
    </blockquote>
    <blockquote>
        <p><strong>Power BI:</strong> Utilize o link fornecido para acessar o Power BI e obter insights visuais adicionais.
        </p>
    </blockquote>

    <p>Ficamos à disposição para discutir qualquer ação adicional que possa ser necessária para abordar esses pontos.</p>

    <p></p>
    <p></p>
    """

    if empresa == 'KONTIK BUSINESS TRAVEL' or empresa == 'GRUPO KONTIK':
        email.HTMLBody = corpo_email_1 + corpo_email_2 + corpo_email_3
    else:
        email.HTMLBody = corpo_email_1 + corpo_email_3

    # Anexos
    dashboard_pdf = r'C:\Users\igorsantos\Desktop\DOCS\data-analysis-python\PROCESSADO ERRO\PDFs\Relatorio - ' + \
        f'{datetime.now().strftime("%d.%m.%Y")}.pdf'
    
    quero_passagem = r'C:\Users\igorsantos\Desktop\DOCS\data-analysis-python\PROCESSADO ERRO\Quero Passagem.xlsx'
    
    integra_tour = r'C:\Users\igorsantos\Desktop\DOCS\data-analysis-python\PROCESSADO ERRO\Contabilização Manual - IntegraTur.xlsx'
       
    relatorio_empresa = r'C:\Users\igorsantos\Desktop\DOCS\data-analysis-python\PROCESSADO ERRO\EMPRESAS\Relatorio - ' + \
    f'{empresa}.xlsx'

    relatorio_dash = r'C:\Users\igorsantos\Desktop\DOCS\data-analysis-python\PROCESSADO ERRO\Relatorio - Dash.xlsx'

    email.Attachments.Add(relatorio_empresa)
    # email.Attachments.Add(dashboard_pdf)

    if empresa == 'GRUPO KONTIK' or empresa == 'KONTIK BUSINESS TRAVEL':
        email.Attachments.Add(quero_passagem)
        email.Attachments.Add(integra_tour)
    
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
# geracao_email(relatorio_erro_inovents, 'INOVENTS', emails_inovents['envio'],emails_inovents['copia'])


print('\033[1;32m\n-Emails criado com sucesso!\033[m')