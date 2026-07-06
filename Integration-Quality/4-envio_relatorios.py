import os
import re
import base64
from urllib.parse import unquote
import win32com.client as win32
import pandas as pd
from datetime import datetime


usuario = os.getlogin()
data_path = f'C:\\Users\\{usuario}\\Desktop\\DOCS\\data-analysis-python\\Integration-Quality'

# --- Carga global de dados ---
try:
    novo_arquivo_resolvido = pd.read_excel(f'{data_path}\\Base.xlsx', sheet_name='Novo Arquivo')
    relatorio_base = pd.read_excel(f'{data_path}\\Relatorio - Dash.xlsx', sheet_name='Processado Erro - BASE')
except FileNotFoundError as e:
    raise RuntimeError(f'Arquivo não encontrado: {e}') from None
except Exception as e:
    raise RuntimeError(f'Erro ao carregar arquivos base: {e}') from e

caminho_dashboard = f'{data_path}\\Relatorio - Dash.xlsx'
data_hoje = datetime.now().strftime('%d.%m.%Y')
pasta_pdf = os.path.join(data_path, 'PDFs')
os.makedirs(pasta_pdf, exist_ok=True)
caminho_saida_pdf = os.path.join(pasta_pdf, f'Relatorio - {data_hoje}.pdf')

# ---Exportação do Dashboard como PDF---
# Abre o Excel via COM, atualiza todas as tabelas dinâmicas da aba "Dashboard",
# salva e exporta como PDF para ser usado como anexo nos e-mails.
excel = win32.Dispatch('Excel.Application')
excel.Visible = False

try:
    wb = excel.Workbooks.Open(caminho_dashboard)
    aba_dashboard = wb.Sheets("Dashboard")
    for pt in aba_dashboard.PivotTables():
        pt.RefreshTable()
    wb.Save()
    aba_dashboard.ExportAsFixedFormat(0, caminho_saida_pdf)
    wb.Close()
    print('\033[1;36m- Guia "Dashboard" atualizada e exportada com sucesso!\033[m')
except Exception as e:
    raise RuntimeError(f'Erro ao atualizar a guia Dashboard: {e}') from e
finally:
    excel.Quit()

# --- Listas de e-mails por empresa ---
emails_corp = {
    'envio': [
        'wagneyoliveira@kontik.com.br','relatoriosgi@kontik.com.br','wellingtonribeiro@kontik.com.br',
          'michellysilva@kontik.com.br', 'eduardomanso@kontik.com.br', 'giselecarmo@kontik.com.br', 
          'nucleonabr@kontik.com.br', 'cartaoaereo@kontik.com.br', 'jackelinenascimento@kontik.com.br',
          'reinildosantos@kontik.com.br', 'andreiaalves@kontik.com.br', 'herbertsantana@kontik.com.br',
          'anafeitosa@kontik.com.br', 'alinemarinho@kontik.com.br', 'mylenasilva@kontik.com.br',
          'giseledenck@kontik.com.br', 'andressasilva@kontik.com.br'
    ],
    'copia': [
        'alexandrecastro@kontik.com.br', 'lanatakuma@kontik.com.br', 'thiagobatello@kontik.com.br',
        'danielacoelho@kontik.com.br', 'rafaelzizzi@kontik.com.br', 'pliniocarvalho@kontik.com.br'
    ]
}

emails_zupper = {
    'envio': ['higorlima@zupper.com.br'],
    'copia': ['angelasilva@zupper.com.br', 'pliniocarvalho@kontik.com.br', 'financeiro@zupper.com.br']
}

emails_kontrip = {
    'envio': ['administrativo@kontrip.com.br'],
    'copia': ['pliniocarvalho@kontik.com.br']
}

emails_grpkontik = {
    'envio': [
        'mylenasilva@kontik.com.br', 'icaroxavier@kontik.com.br',
        'conciliacao_aereo@kontik.com.br', 'suporte.benner@kontik.com.br',
        'thiagobatello@kontik.com.br', 'wellingtonribeiro@kontik.com.br'
    ],
    'copia': ['pliniocarvalho@kontik.com.br', 'williancardoso@kontik.com.br']
}

emails_ktk = {
    'envio': ['girlacarneiro@kontik.com.br'],
    'copia': ['pliniocarvalho@kontik.com.br']
}

emails_inovents = {
    'envio': ['flaviomazzola@inovents.com.br'],
    'copia': [
        'alexandrecastro@kontik.com.br', 'administrativo@inovents.com.br',
        'lucianagarcez@inovents.com.br', 'pliniocarvalho@kontik.com.br'
    ]
}

vendas_integratour = 3000

# Mapeamento empresa → caminho do relatório Excel individual gerado no script anterior
EMPRESA_CAMINHOS = {
    'ZUPPER VIAGENS':           f'{data_path}\\EMPRESAS\\Relatorio - ZUPPER VIAGENS.xlsx',
    'KONTIK BUSINESS TRAVEL':   f'{data_path}\\EMPRESAS\\Relatorio - KONTIK BUSINESS TRAVEL.xlsx',
    'KONTRIP VIAGENS':          f'{data_path}\\EMPRESAS\\Relatorio - KONTRIP VIAGENS.xlsx',
    'INOVENTS':                 f'{data_path}\\EMPRESAS\\Relatorio - INOVENTS.xlsx',
    'GRUPO KONTIK':             f'{data_path}\\EMPRESAS\\Relatorio - GRUPO KONTIK.xlsx',
}

EMPRESAS_BASE_DASH = {'GRUPO KONTIK', 'KONTIK BUSINESS TRAVEL'}
AGING_ALTO = ['16 a 23 dias', '24 a 31 dias', '31 dias ou +']
OBTS_PRINCIPAIS = ['ARGO(TMS)', 'SABRE', 'GOVER', 'LEMONTECH','WOOBA']


def _get_assinatura():
    """Lê a assinatura .htm mais recente do Outlook e embute imagens como base64."""
    sig_dir = os.path.join(os.environ['APPDATA'], 'Microsoft', 'Signatures')
    if not os.path.isdir(sig_dir):
        return ''
    arquivos = [f for f in os.listdir(sig_dir) if f.lower().endswith('.htm')]
    if not arquivos:
        return ''
    mais_recente = max(arquivos, key=lambda f: os.path.getmtime(os.path.join(sig_dir, f)))
    caminho_htm = os.path.join(sig_dir, mais_recente)
    with open(caminho_htm, encoding='utf-8', errors='ignore') as f:
        html = f.read()

    # Extrai só o conteúdo do <body> para não aninhar documentos HTML completos
    body_match = re.search(r'<body[^>]*>(.*?)</body>', html, re.IGNORECASE | re.DOTALL)
    fragmento = body_match.group(1) if body_match else html

    mime_map = {'png': 'image/png', 'jpg': 'image/jpeg', 'jpeg': 'image/jpeg',
                'gif': 'image/gif', 'bmp': 'image/bmp'}

    def substituir_src(match):
        src = match.group(1)
        if src.startswith(('http', 'data:', 'cid:')):
            return match.group(0)
        img_path = os.path.join(os.path.dirname(caminho_htm), unquote(src).replace('/', os.sep))
        if not os.path.exists(img_path):
            return match.group(0)
        ext = os.path.splitext(img_path)[1].lower().lstrip('.')
        mime = mime_map.get(ext, 'image/png')
        with open(img_path, 'rb') as f:
            b64 = base64.b64encode(f.read()).decode()
        return f'src="data:{mime};base64,{b64}"'

    return re.sub(r'src="([^"]+)"', substituir_src, fragmento)


def _set_sender(outlook, email, smtp_address):
    """Configura o remetente para caixa compartilhada Exchange."""
    smtp_lower = smtp_address.lower()
    try:
        for acc in outlook.Session.Accounts:
            if acc.SmtpAddress.lower() == smtp_lower:
                email.SendUsingAccount = acc
                return
    except Exception:
        pass
    try:
        for acc in outlook.Session.Accounts:
            if acc.AccountType == 0:  # olExchange
                email.SendUsingAccount = acc
                break
        email.SentOnBehalfOfName = smtp_address
    except Exception:
        pass


def _top_ofensor_obt(df, obt, total_casos):
    """Retorna (campo, qtd, porcentagem) para um dado OBT, ou valores nulos se ausente."""
    try:
        contagem = df.loc[df['OBTS'] == obt, 'CAMPO'].value_counts()
        campo = contagem.index[0]
        qtd = int(contagem.iloc[0])
        return campo, qtd, (qtd / total_casos) * 100
    except (IndexError, ZeroDivisionError):
        return '-', 0, 0.0


# ---Geração do e-mail analítico por empresa---
# Para cada empresa, carrega seu relatório, calcula métricas (aging crítico, top grupos,
# % por categoria de erro, maiores ofensores por OBT) e monta um e-mail HTML completo
# com anexo e assinatura. O e-mail é salvo como rascunho no Outlook (não enviado diretamente).
def geracao_email(empresa='GRUPO KONTIK', email_envio=None, email_copia=None):
    if email_envio is None:
        email_envio = emails_grpkontik['envio']
    if email_copia is None:
        email_copia = emails_grpkontik['copia']

    if empresa not in EMPRESA_CAMINHOS:
        print(f'\033[1;31m- Empresa "{empresa}" não encontrada!\033[m')
        return

    caminho_empresa = EMPRESA_CAMINHOS[empresa]
    if not os.path.exists(caminho_empresa):
        print(f'\033[1;33m- Arquivo não encontrado para {empresa}\033[m')
        return

    try:
        if empresa in EMPRESAS_BASE_DASH:
            df = pd.read_excel(f'{data_path}\\Relatorio - Dash.xlsx', sheet_name='Processado Erro - BASE')
        else:
            df = pd.read_excel(caminho_empresa)
    except Exception as e:
        print(f'\033[1;31m- Erro ao ler relatório de {empresa}: {e}\033[m')
        return

    total_casos = len(df)
    if total_casos == 0:
        print(f'\033[1;31m- Não há casos para {empresa}, e-mail não será gerado.\033[m')
        return

    top_5_grp_emp = ', '.join(df['Grupo Empresarial'].value_counts().head(5).index)

    soma_aging_alteracao = df['Aging Alteração'].isin(AGING_ALTO).sum()
    soma_aging_inclusao  = df['Aging Inclusão'].isin(AGING_ALTO).sum()

    # Casos que retornaram — busca vetorizada O(n) em vez de loop aninhado O(n*m)
    handles_resolvidos = set(
        str(h) for h in novo_arquivo_resolvido.loc[
            novo_arquivo_resolvido['Status'] == 'Resolvido', 'Handle PNR'
        ]
    )
    mask_retornados = df['Handle PNR'].astype(str).apply(
        lambda pnr: any(h in pnr for h in handles_resolvidos)
    )
    casos_retornados = df.loc[mask_retornados, 'Localizadora'].unique().tolist()

    pct_qualidade   = (df['CATEGORIA DE ERRO'] == 'Qualidade dos dados').sum() / total_casos * 100
    pct_sistemico   = (df['CATEGORIA DE ERRO'] == 'Sistêmico').sum()          / total_casos * 100
    pct_operacional = (df['CATEGORIA DE ERRO'] == 'Processo Operacional').sum()/ total_casos * 100

    # Bloco exclusivo KONTIK BUSINESS TRAVEL / GRUPO KONTIK
    corpo_email_2 = ''
    if empresa in EMPRESAS_BASE_DASH:
        ofensores = sorted(
            [{'obt': obt, 'campo': campo, 'qtd': qtd, 'pct': pct}
             for obt in OBTS_PRINCIPAIS
             for campo, qtd, pct in [_top_ofensor_obt(df, obt, total_casos)]],
            key=lambda x: x['pct'], reverse=True
        )

        itens_obt = ''.join(
            f'<li><strong>{o["obt"]}:</strong> {o["qtd"]} casos de {o["campo"]} '
            f'sendo {o["pct"]:.2f}% do total de casos</li>'
            for o in ofensores
        )

        corpo_email_2 = f"""
        </ul>

    <p><strong>🔥 Maiores Ofensores por OBT:</strong></p>
        <ul>
            {itens_obt}
        """

    link_sd = 'https://grupokontik.atlassian.net/servicedesk/customer/portal/4/group/111'
    link_bi = 'Inserir link do Power BI aqui'

    corpo_email_1 = f"""
    <style>
        p,ul,li {{
            font-size: 11pt;
        }}
    </style>
    <p>Bom dia, pessoal!</p>

    <p>Segue abaixo a análise detalhada do <strong>Processado Erro</strong>, com base no arquivo recebido hoje.</p>

    <p><strong>📌 Para solicitações ao Suporte Benner, é imprescindível a abertura de chamado via Jira <a href="{link_sd}" style="color: #007bff;">aqui</a> ou no caminho:
    <br>➡️ Portal Benner → Contabilização → Pendentes (Processado Erro)</strong></p>

    <p>🔗<a href="{link_bi}" style="color: #007bff;"><strong>Clique aqui para acessar o Power Bi</strong></a></p>

    <p><strong>🔍 Pontos de Atenção:</strong></p>
        <ul>
            <li><strong>Grupos empresariais que mais impactam:</strong> {top_5_grp_emp}</li>
            <li><strong>Aging Alteração acima de 15 Dias:</strong> {soma_aging_alteracao} casos, indicando a necessidade de atenção especial</li>
            <li><strong>Aging Inclusão acima de 15 Dias:</strong> {soma_aging_inclusao} casos, indicando a necessidade de atenção especial</li>
            <li><strong>Casos que retornaram:</strong> Identificamos {len(casos_retornados)}: {casos_retornados}</li>
            <li><strong>Porcentagem de Erros:</strong>
                <ul>
                    <li>{pct_qualidade:.2f}% – Qualidade dos Dados</li>
                    <li>{pct_sistemico:.2f}% – Sistêmico</li>
                    <li>{pct_operacional:.2f}% – Processo Operacional</li>
                </ul>
            </li>
    """

    corpo_email_3 = """
    </ul>
    <p><strong>✅ Ações Recomendadas:</strong></p>
        <ul>
            <li><strong>Priorização:</strong> Foco na resolução dos casos relacionados aos grupos empresariais com maior impacto.</li>
            <li><strong>Aging > 15 dias:</strong> Monitorar com atenção os aging mais antigos para evitar atrasos no processo.</li>
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

    try:
        outlook = win32.Dispatch('outlook.application')
        email = outlook.CreateItem(0)
    except Exception as e:
        print(f'\033[1;31m- Erro ao iniciar Outlook para {empresa}: {e}\033[m')
        return

    email.to = ';'.join(email_envio)
    email.cc = ';'.join(email_copia)
    email.Subject = f'📊 Análise Diária - Qualidade de Integração | {datetime.now().strftime("%d/%m/%Y")} | {empresa}'

    _set_sender(outlook, email, 'suporte.benner@kontik.com.br')

    corpo_completo = corpo_email_1 + corpo_email_2 + corpo_email_3
    assinatura = _get_assinatura()
    email.HTMLBody = f'<html><body style="font-family:Calibri,Arial;font-size:11pt">{corpo_completo}{assinatura}</body></html>'

    dashboard_pdf         = os.path.join(data_path, 'PDFs', f'Relatorio - {data_hoje}.pdf')
    caminho_relat_empresa = os.path.join(data_path, 'EMPRESAS', f'Relatorio - {empresa}.xlsx')
    relatorio_dash        = os.path.join(data_path, 'Relatorio - Dash.xlsx')

    # GRUPO KONTIK recebe o Dash completo + PDF; demais empresas recebem apenas seu relatório
    if empresa == 'GRUPO KONTIK':
        if os.path.exists(relatorio_dash):
            email.Attachments.Add(relatorio_dash)
        if os.path.exists(dashboard_pdf):
            email.Attachments.Add(dashboard_pdf)
        else:
            print(f'\033[1;33m- PDF do dashboard não encontrado: {dashboard_pdf}\033[m')
    else:
        if os.path.exists(caminho_relat_empresa):
            email.Attachments.Add(caminho_relat_empresa)
        else:
            print(f'\033[1;33m- Anexo não encontrado para {empresa}: {caminho_relat_empresa}\033[m')

    email.Send()
    print(f'\033[1;32m- E-mail da empresa {empresa} enviado com sucesso!\033[m')


# ---Envio do arquivo Base.xlsx para o Suporte Benner---
# Cria um e-mail com o Base.xlsx em anexo, mini resumo de erros e lista de e-mails gerados.
def envio_base(resumo_emails=None):
    caminho_base = os.path.join(data_path, 'Base.xlsx')
    if not os.path.exists(caminho_base):
        print(f'\033[1;33m- Base.xlsx não encontrado: {caminho_base}\033[m')
        return

    try:
        outlook = win32.Dispatch('outlook.application')
        email = outlook.CreateItem(0)
    except Exception as e:
        print(f'\033[1;31m- Erro ao iniciar Outlook para envio da Base: {e}\033[m')
        return

    email.to = 'suporte.benner@kontik.com.br'
    email.Subject = f'📎 Base de Dados | {datetime.now().strftime("%d/%m/%Y")}'

    total_casos = len(relatorio_base)
    nao_identificados = int(
        relatorio_base['CAMPO'].str.strip().str.lower().eq('não identificado').sum()
    )

    resumo_metricas = f"""
<p><strong>📋 Resumo — Processado Erro:</strong></p>
<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;font-family:Arial;font-size:11pt">
    <tr style="background:#f2f2f2">
        <td><strong>Total de casos</strong></td>
        <td>{total_casos}</td>
    </tr>
    <tr>
        <td><strong>Erros não identificados</strong></td>
        <td>{nao_identificados}</td>
    </tr>
</table>
"""

    resumo_envios = ''
    if resumo_emails:
        itens = ''.join(f'<li>{e["empresa"]}</li>' for e in resumo_emails)
        resumo_envios = f"""
<hr style="border:none;border-top:1px solid #ccc;margin:16px 0">
<p><strong>📧 E-mails gerados neste ciclo:</strong></p>
<ul style="font-family:Arial;font-size:11pt">{itens}</ul>
"""

    corpo = resumo_metricas + resumo_envios
    email.HTMLBody = f'<html><body style="font-family:Arial;font-size:11pt">{corpo}</body></html>'

    email.Attachments.Add(caminho_base)
    _set_sender(outlook, email, 'suporte.benner@kontik.com.br')
    email.Send()
    print('\033[1;32m- E-mail com Base.xlsx enviado com sucesso!\033[m')


# ---Disparo dos e-mails para todas as empresas---
emails_enviados = [
    {'empresa': 'GRUPO KONTIK',          'envio': emails_grpkontik['envio'], 'copia': emails_grpkontik['copia']},
    {'empresa': 'ZUPPER VIAGENS',         'envio': emails_zupper['envio'],    'copia': emails_zupper['copia']},
    {'empresa': 'KONTIK BUSINESS TRAVEL', 'envio': emails_corp['envio'],      'copia': emails_corp['copia']},
    {'empresa': 'KONTRIP VIAGENS',        'envio': emails_kontrip['envio'],   'copia': emails_kontrip['copia']},
    {'empresa': 'INOVENTS',               'envio': emails_inovents['envio'],  'copia': emails_inovents['copia']},
]

geracao_email()
geracao_email('ZUPPER VIAGENS',          emails_zupper['envio'],   emails_zupper['copia'])
geracao_email('KONTIK BUSINESS TRAVEL',  emails_corp['envio'],     emails_corp['copia'])
geracao_email('KONTRIP VIAGENS',         emails_kontrip['envio'],  emails_kontrip['copia'])
geracao_email('INOVENTS',                emails_inovents['envio'], emails_inovents['copia'])
envio_base(emails_enviados)

print()
print('\033[1;32m- Emails enviados com sucesso!\033[m')
