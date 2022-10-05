import pandas as pd
import plotly.express as px

dados_x = ['2018', '2019', '2020', '2021']
dados_y = [10, 20, 5, 35]

# Gráfico de linhas
fig = px.line(x=dados_x, y=dados_y, title='Vendas por ano',
              width=600, height=300, line_shape='spline')

# -Configurando o gráfico - line_shape
# spline - curva suave
# hvn - horizontal vertical normal
# hv - horizontal vertical
# vh - vertical horizontal
# linear - linha reta
fig.update_yaxes(title='Vendas', title_font_color='red')
fig.show()  # Mostrando o gráfico

# Gráfico de Pizza
data_x = ['2018', '2019', '2020', '2021']
data_y = [10, 20, 5, 35]

fig = px.pie(names=data_x, values=data_y,
             title='Vendas por ano', width=500, height=500)
fig.update_traces(title_text='Pizza', title_position='bottom left')
fig.show()

# Gráfico de barras
data_x = ['2018', '2019', '2020', '2021']
data_y = [10, 20, 5, 35]

fig = px.bar(x=data_x, y=data_y, title='Vendas por ano', width=500, height=500)
fig.update_layout(title_text='Barras', title_x=0.5)
fig.show()

# Gráfico de dispersão
dados_x = [1, 4, 6, 7, 8, 4, 3, 2, 1, 5]
dados_y = [10, 20, 5, 35, 2, 3, 40, 25, 16, 27]

fig = px.scatter(x=dados_x, y=dados_y, width=600, height=600)
fig.show()

# Gráfico de Gantt com dataframe
tarefas = pd.read_excel('data/tarefas.xlsx')

fig = px.timeline(tarefas, x_start='Início',
                  x_end='Fim', y='Tarefa', width=1200, height=600)

fig.update_yaxes(autorange='reversed')
fig.show()
