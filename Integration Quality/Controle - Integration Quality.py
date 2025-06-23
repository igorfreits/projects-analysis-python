import subprocess
import os

data_path = 'data-analysis-python/Integration Quality/'

scripts = [
    "0 - Integrado Erro.py",
    "1 - processamento_erros.py",
    "2 - atualizacao_status.py",
    "3 - envio_relatorios.py"
]

for script in scripts:
    script_path = os.path.join(data_path, script)
    print(f"\033[1;36mExecutando {script_path}...\033[0m\n")
    
    result = subprocess.run(["python", script_path])
    
    if result.returncode != 0:
        print(f"\033[1;31mErro ao rodar {script}.\033[0m\n") 
        break
    else:
        print() 
