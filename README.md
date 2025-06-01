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

### 4. `pokemon_api_populate.py` 🎮

* 🎯 Integra com a API pública do Pokémon para popular um banco de dados PostgreSQL com dados estruturados.
* 🗃️ Cria e mantém as tabelas de `pokemons`, `tipos`, `regioes`, `imagens` e `evolucoes`.
* 🔄 Atualiza informações de pokémons, suas evoluções, tipos e regiões de forma automatizada.
* 🖼️ Gerencia URLs das imagens oficiais para consulta e uso em visualizações BI.
* 🚀 Exemplo prático de automação de coleta e transformação de dados para análise e dashboards.

### 5. `ecommerce_data_generator.py` 🛒

* 📦 Gera dados fictícios para um e-commerce focado em serviços de TI, incluindo vendas, clientes, produtos e vendedores.
* 🔢 Popula uma base PostgreSQL com ao menos 1000 registros, permitindo testes, simulações e análises.
* 📊 Auxilia no desenvolvimento de relatórios, dashboards e estudos preditivos baseados em dados realistas.

## 📚 Bibliotecas Utilizadas

Os scripts utilizam as seguintes bibliotecas Python:

* 🐼 `pandas`
* 📂 `openpyxl`
* 📑 `xlrd`
* 📧 `win32com.client`
* ⏳ `datetime`
* 🗂️ `os`
* 🌐 `requests`
* 🐍 `psycopg2`
* ⏲️ `time`

## 🔧 Requisitos

Para instalar as bibliotecas necessárias, rode:

pip install pandas openpyxl xlrd pywin32 requests psycopg2-binary

.

## ▶️ Como Usar

1. **Processamento de Erros:**

   * python processamento_erros.py
2. **Atualização de Status:**

   * python atualizacao_status.py
3. **Envio de Relatórios:**

   * python envio_relatorios.py
4. **População da base Pokémon:**

   * python pokemon_api_populate.py
5. **Geração de dados fictícios para e-commerce:**
   * python pokemon_api_populate.py

## 📂 Estrutura do Repositório

   |-- data-analysis-python/
   |   |-- PROCESSADO ERRO/
   |   |   |-- Base.xlsx
   |   |   |-- Relatorio - Dash.xlsx
   |   |   |-- EMPRESAS/
   |   |       |-- Relatorio - ZUPPER VIAGENS.xlsx
   |   |       |-- Relatorio - KONTIK BUSINESS TRAVEL.xlsx
   |   |       |-- Relatorio - KONTRIP VIAGENS.xlsx
   |   |       |-- Relatorio - GRUPO KONTIK.xlsx
   |-- processamento_erros.py
   |-- atualizacao_status.py
   |-- envio_relatorios.py
   |-- pokemon_api_populate.py
   |-- ecommerce_data_generator.py

## 🤝 Contribuição

   Contribuições são muito bem-vindas! Abra um pull request para melhorias, correções ou sugestões.

## ⚖️ Licença

   Este projeto está sob a licença MIT. Consulte o arquivo LICENSE para mais detalhes.
