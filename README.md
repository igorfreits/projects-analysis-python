# 🚀 Análise e Processamento de Erros com Python 📊

Este repositório contém scripts Python para análise, categorização e processamento de erros a partir de arquivos Excel, gerando relatórios detalhados e automatizando o envio de e-mails com os dados analisados. Além disso, os dados processados podem ser visualizados no **Power BI** para facilitar a análise 📈.

## 🛠️ Estrutura dos Scripts

### 1. `processamento_erros.py` 📝

* 🔄 Converte arquivos `.xls` para `.xlsx`.
* 🔍 Processa arquivos de erro, realizando a formatação dos dados e categorizando os erros por tipo, origem e responsabilidade.
* 📊 Aplica filtros e cria colunas adicionais para melhor organização.
* 🏢 Segmenta os dados processados e gera relatórios organizados por empresa.
* 🎨 Formata e salva os dados em arquivos Excel, aplicando estilos personalizados.

### 2. `atualizacao_status.py` 🔄

* 📌 Atualiza o status dos erros identificados como "Novo", "Em Andamento" ou "Resolvido".
* ✅ Identifica registros resolvidos e os move para uma planilha de resoluções.
* 🗑️ Remove registros resolvidos da base de erros em andamento.
* 💾 Salva os dados atualizados em um arquivo Excel sem alterar outras abas.

### 3. `envio_relatorios.py` ✉️

* 📑 Lê os relatórios processados e segmentados por empresa.
* 🔎 Identifica padrões e categorias de erro para compilar insights.
* 📨 Gera e-mails automáticos formatados com análises detalhadas.
* 📎 Anexa relatórios e outros documentos relevantes.
* 📤 Envia os e-mails para listas predefinidas de destinatários.

## 📚 Bibliotecas Utilizadas

Os scripts utilizam as seguintes bibliotecas Python:

* 🐼 `pandas`: Para manipulação e análise de dados.
* 📂 `openpyxl`: Para leitura e escrita de arquivos Excel no formato `.xlsx`.
* 📑 `xlrd`: Para leitura de arquivos `.xls` (necessário para conversão para `.xlsx`).
* 📧 `win32com.client`: Para integração com o Microsoft Outlook e envio automatizado de e-mails.
* ⏳ `datetime`: Para manipulação de datas nos relatórios.
* 🗂️ `os`: Para manipulação de diretórios e arquivos.

## 🔧 Requisitos

Para rodar os scripts, instale as bibliotecas necessárias:

```bash
pip install pandas openpyxl xlrd pywin32
```

## ▶️ Como Usar

1. **Processamento de Erros:**

   ```bash
   python processamento_erros.py
   ```

   Esse script irá converter arquivos, processar os dados e gerar relatórios segmentados.
2. **Atualização de Status:**

   ```bash
   python atualizacao_status.py
   ```

   Ele atualiza o status dos registros de erro e salva os dados atualizados no Excel.
3. **Envio de Relatórios:**

   ```bash
   python envio_relatorios.py
   ```

   O script gera e-mails formatados com análises e relatórios anexados.

## 📂 Estrutura de Diretórios

```
/
|-- data-analysis-python/
|   |-- PROCESSADO ERRO/
|   |   |-- Base.xlsx
|   |   |-- Relatorio - Dash.xlsx
|   |   |-- EMPRESAS/
|   |   |   |-- Relatorio - ZUPPER VIAGENS.xlsx
|   |   |   |-- Relatorio - KONTIK BUSINESS TRAVEL.xlsx
|   |   |   |-- Relatorio - KONTRIP VIAGENS.xlsx
|   |   |   |-- Relatorio - GRUPO KONTIK.xlsx
|-- processamento_erros.py
|-- atualizacao_status.py
|-- envio_relatorios.py
```

## 🤝 Contribuição

Se quiser contribuir, sinta-se à vontade para abrir um pull request com melhorias ou correções.

## ⚖️ Licença

Este projeto está sob a licença MIT. Consulte o arquivo LICENSE para mais detalhes.

🔎 **Visualização dos dados no Power BI disponível!** 📊
