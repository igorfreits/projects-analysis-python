# 🛠️ Processamento de Erros e Geração de Relatórios

Este repositório contém scripts desenvolvidos em Python para a manipulação, análise e distribuição de relatórios de erros processados. São utilizados dados oriundos de arquivos Excel, que são processados e consolidados em dashboards e enviados via email para as equipes responsáveis.

## 📁 Conteúdo

- **processamento_dados.py**  
  Realiza a leitura, tratamento e formatação dos dados do arquivo "Processado Erro.xlsx". Entre as atividades executadas estão:  
  - Limpeza e padronização de colunas e valores nulos.  
  - Criação de novas colunas e cálculos (por exemplo, "Dias Parados no Erro" e "Mês Alteração").  
  - Realocações de registros conforme regras definidas (por exemplo, atribuição de responsáveis, empresas e categorias de erro).  
  - Geração de relatórios em Excel com dashboards customizados para cada empresa.

- **atualizacao_base.py**  
  Atualiza as bases de dados de erros, gerenciando registros novos, em andamento e resolvidos. As principais funções deste script são:  
  - Limpeza da base de novos registros e criação de uma nova base estruturada.  
  - Verificação se um registro já consta na base "Em Andamento" ou "Resolvidos", alterando o status conforme necessário.  
  - Consolidação dos registros resolvidos e atualização das datas de conclusão.  
  - Salvamento dos dados atualizados de volta em um arquivo Excel com múltiplas guias.

- **geracao_email.py**  
  Gera e envia e-mails personalizados com a análise dos erros, utilizando o Outlook via `win32com.client`. As principais funcionalidades são:  
  - Leitura dos relatórios gerados para cada empresa.  
  - Extração de métricas (total de casos, grupos empresariais de maior impacto, aging acima de 15 dias, principais ofensores etc.).  
  - Montagem de um email com corpo em HTML contendo as informações consolidadas e anexando os relatórios relevantes.  
  - Envio dos e-mails para destinatários específicos (configurados via dicionários de emails) com cópias, conforme a empresa.

## 📦 Dependências

Para executar os scripts, certifique-se de ter instaladas as seguintes bibliotecas Python:

- [pandas](https://pandas.pydata.org/)
- [openpyxl](https://openpyxl.readthedocs.io/)
- [pywin32](https://github.com/mhammond/pywin32) (necessário para integração com o Outlook)

Você pode instalá-las utilizando o `pip`:

```bash
pip install pandas openpyxl pywin32
🗂️ Estrutura de Pastas e Arquivos
Copiar
Editar
├── PROCESSADO ERRO
│   └── Analise de Dados
│       ├── Relatorio - Dash.xlsx
│       ├── Base.xlsx
│       ├── EMPRESAS
│           ├── Relatorio - ZUPPER VIAGENS.xlsx
│           ├── Relatorio - KONTIK BUSINESS TRAVEL.xlsx
│           ├── Relatorio - KONTRIP VIAGENS.xlsx
│           ├── Relatorio - INOVENTS.xlsx
│           └── Relatorio - GRUPO KONTIK.xlsx
├── processamento_dados.py
├── atualizacao_base.py
└── geracao_email.py
💡 Observação:

Verifique se os arquivos Excel estão organizados conforme o esperado e se as planilhas (sheets) possuem os nomes corretos.
Alguns textos podem apresentar problemas de codificação (ex.: "SistÃªmico" em vez de "Sistêmico"). Recomenda-se utilizar UTF-8 ao salvar e ler os arquivos para evitar inconsistências.
🚀 Como Utilizar
Processamento de Dados e Geração de Relatórios
Execute o script processamento_dados.py para processar os dados dos arquivos Excel, aplicar as regras de tratamento e gerar os relatórios (incluindo a criação dos dashboards e planilhas por empresa).

Atualização das Bases de Dados
Após a geração dos relatórios, execute o script atualizacao_base.py para atualizar os status dos registros (Novo, Em Andamento, Resolvido) e consolidar as bases de dados em um único arquivo Excel.

Geração e Envio de Emails
Por fim, execute o script geracao_email.py para gerar os emails com a análise dos erros e enviá-los aos destinatários configurados.
⚠️ Atenção: O script utiliza o Outlook instalado na máquina para envio dos emails. Verifique as configurações e permissões do Outlook para automação.

⚙️ Configurações Específicas
Dados de Entrada:
Os scripts assumem que os arquivos Excel estão localizados na pasta PROCESSADO ERRO/Analise de Dados/ e que as planilhas possuem os nomes conforme especificados nos códigos.

Envio de Emails:

As listas de destinatários (envio e cópia) estão definidas nos dicionários emails_corp, emails_zupper, emails_kontrip, emails_grpkontik, emails_ktk e emails_inovents.
Certifique-se de ajustar ou atualizar os emails conforme a necessidade do seu ambiente.
💬 Considerações Finais
Testes:
Antes de executar os scripts em produção, recomenda-se testá-los em um ambiente de desenvolvimento para garantir que as regras de negócio e o fluxo de dados estejam corretos.

Suporte e Contribuições:
Se você encontrar algum problema ou tiver sugestões de melhorias, sinta-se à vontade para abrir uma issue ou enviar um pull request.

📧 Contato
Para dúvidas ou mais informações, entre em contato com o responsável pelo projeto ou abra uma issue neste repositório.
