import subprocess
import os
import re
import html
import ctypes
import time
import win32com.client


usuario = os.getlogin()
data_path = f'C:\\Users\\{usuario}\\Desktop\\DOCS\\data-analysis-python\\Integration-Quality\\'

MAIL = 'suporte.benner@kontik.com.br'

SUGESTOES = {
    'FileNotFoundError': 'Verifique se o arquivo existe no caminho indicado e se o nome está correto.',
    'PermissionError': 'Feche o arquivo no Excel e tente novamente.',
    'KeyError': 'Verifique se a coluna existe na planilha — confira o nome exato.',
    'ValueError': 'Verifique o formato dos dados (datas, números, textos) na planilha.',
    'AttributeError': 'Verifique se a coluna ou objeto está sendo acessado corretamente.',
    'RuntimeError': 'Verifique as condições iniciais: e-mail recebido hoje, anexo presente, arquivos no caminho correto e Outlook aberto.',
    'IndexError': 'Verifique se há dados no e-mail ou na planilha de origem.',
    'TypeError': 'Verifique se os tipos de dados (texto, número, data) estão coerentes.',
    'com_error': 'Erro de comunicação com o Outlook (MAPI). Verifique se a caixa compartilhada está montada no perfil, se o nome da pasta está correto e se você tem permissão de acesso a ela.',
    'KeyboardInterrupt': 'O processo foi interrompido manualmente (Ctrl+C) ou encerrado pelo sistema operacional.',
}


def _enviar_email_erro(script, stderr_text):
    linhas = stderr_text.strip().splitlines()

    arquivo = script
    linha = '?'
    for l in reversed(linhas):
        m = re.match(r'\s*File "(.+)", line (\d+)', l)
        if m:
            arquivo = os.path.basename(m.group(1))
            linha = m.group(2)
            break

    tipo_erro = 'Erro desconhecido'
    mensagem = linhas[-1] if linhas else '?'
    if ':' in mensagem:
        tipo_erro = mensagem.split(':')[0].strip()
        mensagem = ':'.join(mensagem.split(':')[1:]).strip()

    sugestao = SUGESTOES.get(tipo_erro, 'Revise o traceback abaixo para identificar a causa raiz.')

    corpo_html = f"""
<h3 style="color:#cc0000;font-family:Arial">&#9888; Erro na execução: {arquivo}</h3>
<table border="1" cellpadding="8" cellspacing="0" style="border-collapse:collapse;font-family:Arial;font-size:13px">
  <tr style="background:#f2f2f2"><td><b>Tipo do erro</b></td><td>{tipo_erro}</td></tr>
  <tr><td><b>Mensagem</b></td><td>{html.escape(mensagem)}</td></tr>
  <tr style="background:#f2f2f2"><td><b>Arquivo</b></td><td>{arquivo}</td></tr>
  <tr><td><b>Linha</b></td><td>{linha}</td></tr>
  <tr style="background:#fff3cd"><td><b>Como tratar</b></td><td>{sugestao}</td></tr>
</table>
<br>
<b style="font-family:Arial">Traceback completo:</b>
<pre style="background:#f8f8f8;padding:12px;font-size:12px;border:1px solid #ddd">{html.escape(stderr_text)}</pre>
"""

    try:
        outlook_app = win32com.client.Dispatch("Outlook.Application")
        mail = outlook_app.CreateItem(0)
        mail.To = MAIL
        mail.Subject = f'[ERRO] Envio Relatório Analise de Erros — {arquivo} (linha {linha})'
        mail.SentOnBehalfOfName = MAIL
        mail.Importance = 2  # olImportanceHigh
        mail.HTMLBody = corpo_html
        mail.Send()
        print('\033[93mE-mail de erro enviado.\033[0m')
    except Exception as e_mail:
        print(f'\033[91mFalha ao enviar e-mail de erro: {e_mail}\033[0m')


def _sincronizar_outlook():
    print('\033[1;36mAbrindo Outlook para sincronizar e-mails...\033[0m')

    outlook = win32com.client.Dispatch("Outlook.Application")
    mapi = outlook.GetNamespace("MAPI")
    mapi.Logon("", "", False)

    try:
        mapi.SendAndReceive(False)
        time.sleep(15)
        print('\033[92mSincronização concluída.\033[0m')
    except Exception:
        print('\033[93mSincronização indisponível — prosseguindo com e-mails já baixados.\033[0m')



def _salvar_e_fechar_excel():
    ctypes.windll.user32.MessageBoxW(
        0,
        'O Excel será fechado e todos os arquivos abertos serão salvos automaticamente.',
        'Aviso — Fechamento do Excel',
        0x30,  # MB_ICONWARNING
    )
    try:
        xl = win32com.client.GetActiveObject("Excel.Application")
        xl.DisplayAlerts = False
        for wb in xl.Workbooks:
            wb.Save()
        xl.Quit()
        print('\033[92mArquivos Excel salvos e Excel fechado.\033[0m')
    except Exception:
        print('\033[93mNenhuma instância do Excel aberta encontrada.\033[0m')


scripts = [
    "2-processamento_erros.py",
    "3-atualizacao_status.py",
    "4-envio_relatorios.py"
]

try:
    _sincronizar_outlook()
    _salvar_e_fechar_excel()
except RuntimeError as e:
    print(f'\033[1;31m{e}\033[0m')
    raise SystemExit(1)

for script in scripts:
    script_path = os.path.join(data_path, script)
    print(f"\033[1;36mExecutando {script_path}...\033[0m\n")

    result = subprocess.run(["python", script_path], stderr=subprocess.PIPE, text=True)

    if result.returncode != 0:
        print(f"\033[1;31mErro ao rodar {script}.\033[0m\n")
        if result.stderr:
            print(result.stderr)
            _enviar_email_erro(script, result.stderr)
        break
    else:
        print()
