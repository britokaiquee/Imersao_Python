# Continuação do arquivo analise_dados
# Data: 10/05 - aulas 4 e 5 da Imersão Python.
# Obs: a imersão foi entre 25 a 29 de março, mas resolvi passar esses códigos
# que fiz no Coolab pro VS Code só hoje.

# Instalando as bibliotecas (módulos)

# Importa a biblioteca pandas para manipulação de dados
import pandas as pd 
# Importa a biblioteca matplotlib para visualização de dados
# (usada para criar gráficos como candlesticks)
import matplotlib.pyplot as plt
# Importa as funcionalidades de datas do matplotlib
import matplotlib.dates as mdates
# Importa a biblioteca mplfinance para visualização de gráficos financeiros
# (extensão do matplotlib, mas que não tem como ajustar detalhes específicos)
import mplfinance as mpf
# Importa a biblioteca yfinance para baixar dados financeiros
import yfinance as yf
# Importa a biblioteca plotly para visualização interativa de dados (gráficos)
import plotly.graph_objects as go
# Importa funcionalidades para criar subplots no Plotly
from plotly.subplots import make_subplots

dados = yf.download('PETR4.SA', start='2023-01-01',
                    end='2023-12-31')  # Baixar a base dos dados

dados.columns = ['Abertura', 'Maximo', 'Minimo', 'Fechamento',
                 'Fech_Ajust', 'Volume']  # Definir os nomes da coluna

dados = dados.rename_axis('Data')  # Renomear

dados['Fechamento'].plot(figsize=(10, 6))  # Gerar gráfico
# Adicionar título ao gráfico e alterar tamanho da fonte
plt.title('Variação do preço por data', fontsize=16)
plt.legend(['Fechamento'])  # Adicionar legenda

# Limitando a quantidade de linhas:
df = dados.head(60).copy()

# Convertendo o índice em uma coluna de data:
df['Data'] = df.index

# Convertendo as datas para o formato numérico de matplotlib:
df['Data'] = df['Data'].apply(mdates.date2num)
# Isso é necessário para que o Matplotlib possa plotar as datas corretamente
# no gráfico.

# Gera um gráfico em branco, os números entre parênteses são as medidas, que
# podem ser alteradas
fig, ax = plt.subplots(figsize=(15, 8))

fig, ax = plt.subplots(figsize=(15, 8))

# Definindo a largura dos candles no gráfico
width = 0.7

for i in range(len(df)):
    '''
    Determinando a cor do candle
    Se o preço de fechamento for maior que o de abertura, o candle é verde
    (a ação valorizou nesse dia).
    Se for menor, o candle é vermelho (a ação desvalorizou).
    '''
    if df['Fechamento'].iloc[i] > df['Abertura'].iloc[i]:
        color = 'green'
    else:
        color = 'red'

    '''
    Desenhando a linha vertical do candle (mecha)
    Essa linha mostra os preços máximo (topo da linha) e mínimo
    (base da linha) do dia.
    Usamos ax.plot para desenhar uma linha vertical.
    ([df['Data'].iloc[i], df['Data'].iloc[i]] define o ponto x da linha
    (a data), e [df['Mínimo'].iloc[i], df['Máximo'].iloc[i]] define a
    altura da mecha.
    '''
    ax.plot([df['Data'].iloc[i], df['Data'].iloc[i]],
            [df['Minimo'].iloc[i], df['Maximo'].iloc[i]],
            color=color,
            linewidth=1)

    ax.add_patch(plt.Rectangle((df['Data'].iloc[i] - width/2, \
                        min(df['Abertura'].iloc[i], df['Fechamento'].iloc[i])),
                               width,
                               abs(df['Fechamento'].iloc[i] -
                                   df['Abertura']. iloc[i]),
                               facecolor=color))
df['MA7'] = df['Fechamento'].rolling(window=7).mean()
df['MA14'] = df['Fechamento'].rolling(window=14).mean()

# Plotando as médias móveis
ax.plot(df['Data'], df['MA7'], color='lime',
        label='Média Móvel 7 Dias')  # Média de 7 dias
ax.plot(df['Data'], df['MA14'], color='skyblue',
        label='Média Móvel 14 Dias')  # Média de 14 dias

# Adicionando legendas para as médias móveis
ax.legend()

# Formatando o eixo x para mostrar as datas
# Configuramos o formato da data e a rotação para melhor legibilidade

# O método xaxis_date() é usado para dizer ao Matplotlib que as datas
# estão sendo usadas no eixo x
ax.xaxis_date() 
ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m-%d'))
plt.xticks(rotation=45)

# Adicionando título e rótulos para os eixos x e y
plt.title("Gráfico de Candlestick - PETR4.SA com matplotlib")
plt.xlabel("Data")
plt.ylabel("Preço")

# Adicionando uma grade para facilitar a visualização dos valores
plt.grid(True)

# Exibindo o gráfico
plt.show()

# Criando subplots
'''
"Primeiro, criamos uma figura que conterá nossos gráficos usando make_subplots.
Isso nos permite ter múltiplos gráficos em uma única visualização.
Aqui, teremos dois subplots: um para o gráfico de candlestick e outro para o
volume de transações."
'''
fig = make_subplots(rows=2, cols=1, shared_xaxes=True,
                    vertical_spacing=0.1,
                    subplot_titles=('Candlesticks', 'Volume Transacionado'),
                    row_width=[0.2, 0.7])

'''
"No gráfico de candlestick, cada candle representa um dia de negociação,
mostrando o preço de abertura, fechamento, máximo e mínimo. Vamos adicionar
este gráfico à nossa figura."
'''
# Adicionando o gráfico de candlestick
fig.add_trace(go.Candlestick(x=df.index,
                             open=df['Abertura'],
                             high=df['Maximo'],
                             low=df['Minimo'],
                             close=df['Fechamento'],
                             name='Candlestick'),
              row=1, col=1)

# Adicionando as médias móveis
# Adicionamos também médias móveis ao mesmo subplot para análise de tendências
fig.add_trace(go.Scatter(x=df.index,
                         y=df['MA7'],
                         mode='lines',
                         name='MA7 - Média Móvel 7 Dias'),
              row=1, col=1)

fig.add_trace(go.Scatter(x=df.index,
                         y=df['MA14'],
                         mode='lines',
                         name='MA14 - Média Móvel 14 Dias'),
              row=1, col=1)

# Adicionando o gráfico de barras para o volume
# Em seguida, criamos um gráfico de barras para o volume de transações, que nos
# dá uma ideia da atividade de negociação naquele dia
fig.add_trace(go.Bar(x=df.index,
                     y=df['Volume'],
                     name='Volume'),
              row=2, col=1)

# Atualizando layout
# Finalmente, configuramos o layout da figura, ajustando títulos, formatos de
# eixo e outras configurações para tornar o gráfico claro e legível.
fig.update_layout(yaxis_title='Preço',
                  xaxis_rangeslider_visible=False,  # Desativa o range slider
                  width=1100, height=600)

# Mostrando o gráfico
fig.show()

# Desafio da aula 4
dados = yf.download('AAPL', start='2024-01-01', end='2024-03-31')

mpf.plot(dados.head(50), type='candle', figsize=(
    16, 8), volume=True, mav=(7, 15), style='yahoo')

dados = yf.download('GOOGL', start='2023-01-01', end='2023-12-31')  # Extra

mpf.plot(dados.head(50), type='candle', figsize=(16, 8),
         volume=True, mav=(7, 15, 21), style='starsandstripes')
