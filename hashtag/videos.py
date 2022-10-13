import seaborn as sns
import pandas as pd
import matplotlib.pyplot as plt

# Load the data
df_videos = pd.read_excel('data/videosYT.xlsx', 'videos')

# Gráfico de dispersão
sns.scatterplot(x='Nº de Views', y='Nº de Likes',
                data=df_videos, hue='Categoria', style='Responsável', palette=['red', 'cyan', 'blue', 'green'])
# hue -  categoriza os pontos
# style -  altera o estilo dos pontos
# palette -  altera a cor dos pontos
plt.show()

# Gráfico de dispersão relacional
grafico_relacional = sns.relplot(
    data=df_videos, x='Nº de Views', y='Nº de Likes', hue='Categoria', col='Responsável')
# col -  coloca os gráficos lado a lado
grafico_relacional.set_titles('Responsável: {col_name}')
# col_name - Nome da coluna
plt.show()

# Gráfico de linas
df_inscritos = pd.read_excel('data/videosYT.xlsx', 'Inscritos')

grafico_linha = sns.lineplot(data=df_inscritos, x='Mês/Ano', y='Inscritos')
grafico_linha.set_title('Inscritos por mês')
plt.show()

# Histogramas
grafico_histograma = sns.displot(
    data=df_videos, x='Nº de Views', hue='Responsável', col='Categoria', rug=True)
# kind='ecdf' - Empirical cumulative distribution function - Função de distribuição acumulada empírica
# kind='hist' - Histograma - Empirical cumulative distribution function - Função de distribuição acumulada empírica(padrão)
# kind='kde' - Kernel density estimation - Estimativa de densidade do núcleo

# rug=True - Mostra a distribuição dos dados
grafico_histograma.set_titles('Categoria: {col_name}')
plt.show()

# Regressão linear
grafico_regressao = sns.regplot(
    data=df_videos, x='Nº de Views', y='Nº de Likes')
plt.show()

grafico_regressao_duplo = sns.lmplot(
    data=df_videos, x='Nº de Views', y='Nº de Likes', hue='Responsável', markers=['o', 'x'])
# markers -  altera o estilo dos pontos
plt.show()
