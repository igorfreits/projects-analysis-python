import subprocess
import os

data_path = 'data-analysis-python/Integration Quality/'

scripts = [
    "1 - Integrado Erro.py",
    "2 - processamento_erros.py",
    "3 - atualizacao_status.py",
    "4 - envio_relatorios.py"
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
