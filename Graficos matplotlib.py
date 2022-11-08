import matplotlib.pyplot as plt

x = [1, 2, 3, 4]
y = [2, 3, 4, 3]

# grafico de linhas
plt.plot(x, y, label='dados', linestyle='dashed',
         color='purple', lw=3.0)

# Parâmetro x and y
# label - nome que aparece na legenda
# linestyle - estilo do  grafico de linha(padrão:solid)
# color - cor da linha(hex,RGB,RGBA string)
# lw - tamanho da linha(largura)

# xlabel e ylabel - nomes para os eixos
plt.ylabel('Eixo Y')
plt.xlabel('Eixo X')

# title - Titulo do grafico
plt.title('Titulo do grafico')

# xticks e yticks - Escala
plt.xticks([0, 2, 4, 6, 8, 9])
plt.yticks([-1, 3, 5, 9, 11])

# legend - Legenda
plt.legend()

plt.axis(xmin=-1, xmax=10, ymin=0, ymax=12)
# axis - limites dos eixos
plt.show()

# Grafico de dispersão
plt.scatter(x, y, marker='s')
# marker - tipo de marcador
plt.show()

# Gráfico de Barra
plt.bar(x, y)
plt.show()

# Subplots
valores_x = [1, 2, 3, 4]
valores_y = [1, 4, 2, 3]
# figsize - tamanho da figura
figura = plt.figure(figsize=(15, 4))

# subtitle - Subtitulo
figura.suptitle('Titulo Geral')

# add_subplot - Adiciona um subgrafico(linha, colunas, grafico)
figura.add_subplot(131)
plt.plot(valores_x, valores_y, label='Um dado qualquer')
plt.title('Grafico 1')
plt.legend()

figura.add_subplot(133)
plt.scatter(valores_x, valores_y, label='Outro dado qualquer')
plt.title('Grafico 2')

figura.add_subplot(132)
plt.bar(valores_x, valores_y, label='Mais um dado qualquer')
plt.title('Grafico 3')

plt.save
plt.show()
