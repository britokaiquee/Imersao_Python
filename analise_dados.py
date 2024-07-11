# Criação do arquivo: dia 02/04/2024 - aulas 2 e 3 da Imersão Python.

# ↓↓↓ Importando as bibliotecas ↓↓↓

# Biblioteca usada para manipulação e análise de dados (tabelas)
import pandas as pd

# Biblioteca usada para fazer gráficos
import plotly.express as px

# "plotly.graph_objects" é uma parte dessa biblioteca
# "go" é o alias, que é usado como prefixo (assim como pd, px, etc)
import plotly.graph_objects as go

# Biblioteca para colocar bordas nas tabelas
from tabulate import tabulate

# Para tirar a notação científica ("e+" no meio dos números grandes)
pd.options.display.float_format = '{:.2f}'.format

# Definindo o caminho do arquivo
caminho_arquivo = r"c:\Users\Kaique\Downloads\desafios _imersao_python.xlsx"
''' Alternartiva: 
caminho_arquivo = "c:\\Users\\Kaique\\Downloads\\desafios _imersao_python.xlsx"
'''

# Carregando os DataFrames
df_principal = pd.read_excel(caminho_arquivo, sheet_name="principal")
df_acoes = pd.read_excel(caminho_arquivo, sheet_name="total_de_açoes")
df_ticker = pd.read_excel(caminho_arquivo, sheet_name="ticker")
df_chatgpt = pd.read_excel(caminho_arquivo, sheet_name="chatgpt")

# ↓↓↓ Ajustando o DataFrame principal ↓↓↓

# Faz aparecer colunas específicas
# Sem precisar ser na mesma ordem da planilha
df_principal = df_principal[["Ativo", "Data",
                             "Último (R$)", "Var. Dia (%)"]].copy()

# A função "rename" serve para renomear colunas
df_principal = df_principal.rename(columns={
    'Último (R$)': 'Valor_final', 'Var. Dia (%)': 'Var_dia',
    'Data': 'Data', 'Resultado': 'Resultado'}).copy()

# Cria uma nova coluna com base em outra e dividindo os valores
df_principal['Var_pct'] = df_principal['Var_dia'] / 100

# Mesma coisa, só que com outro operador
df_principal['Valor_inicial'] = df_principal['Valor_final'] / \
    (df_principal['Var_pct'] + 1)

''' "merge" combina os DataFrames, usando 'Ativo' como chave e adicionando
as colunas. Semelhante ao PROCV do Excel, só que mais eficaz '''
df_principal = df_principal.merge(
    df_acoes, left_on='Ativo', right_on='Código', how='left')

df_principal = df_principal.rename(
    columns={'Qtde. Teórica': 'Qtde_teorica'}).copy()

# A função "drop" serve para dropar coluna(s)
df_principal = df_principal.drop(columns=['Código'])

# Para criar mais outra coluna, com base em outras três
df_principal['Var_rs'] = (df_principal['Valor_final'] -
                          df_principal['Valor_inicial']) * \
                            df_principal['Qtde_teorica']

# Deixa os números da coluna como int (inteiros)
df_principal['Qtde_teorica'] = df_principal['Qtde_teorica'].astype(int)

# Mesmo objetivo da função IF do Excel
df_principal['Resultado'] = df_principal['Var_dia'].apply(
    lambda x: "Subiu" if x > 0 else ("Desceu" if x < 0 else "Estável"))

df_principal = df_principal.merge(
    df_ticker, left_on='Ativo', right_on='Ticker', how='left')

df_principal = df_principal.drop(columns=['Ticker'])

df_principal = df_principal.merge(
    df_chatgpt, left_on='Nome', right_on='Nome da empresa', how='left')

df_principal = df_principal.drop(columns=['Nome da empresa'])

df_principal = df_principal.rename(
    columns={'Idade (anos)': 'Idade'}).copy()

# Tira as linhas repetidas
df_principal = df_principal.drop_duplicates().reset_index(drop=True)

# Definindo os limites das faixas etárias
faixas_etarias = [0, 20, 40, 60, 80, float('inf')]
labels_faixas = ['Entre 0-20 anos', 'Entre 21-40 anos',
                 'Entre 41-60 anos', 'Entre 61-80 anos', 'Mais de 80 anos']

# Cria a coluna com as faixas etárias
df_principal['Faixa_etaria'] = pd.cut(
    df_principal['Idade'], bins=faixas_etarias, 
    labels=labels_faixas, right=False)

''' Exemplo de segmentação por resultado:
df_principal_subiu = df_principal[df_principal['Resultado'] == 'Subiu'].copy()
Esse código filtra as empresas que tiveram como resultado subiu e dá para fazer
o mesmo com estável e desceu '''

# ↓↓↓ Contagens ↓↓↓

# Empresas por resultado
df_contagem_resultado = df_principal['Resultado'].value_counts().reset_index()
df_contagem_resultado.columns = ['Resultado', 'Quantidade']
df_contagem_resultado.index = df_contagem_resultado.index + 1

# Empresas por segmento
df_contagem_segmento = df_principal['Segmento'].value_counts().reset_index()
df_contagem_segmento.columns = ['Segmento', 'Quantidade']
df_contagem_segmento.index = df_contagem_segmento.index + 1

# Empresas por faixa etária
df_empresas_por_faixa_etaria = df_principal.groupby(
    'Faixa_etaria', observed=False).size().reset_index(name='Qtde_de_empresas')
# Para o índice começar pelo nº1 ↓
df_empresas_por_faixa_etaria.index = df_empresas_por_faixa_etaria.index + 1

# ↓↓↓ Análises ↓↓↓

''' "Groupby" cria um resumo dos dados por meio de uma fórmula e em ordem
alfabética, semelhante a função "UNIQUE" do Excel, e "sum" a função "SUMIF".
Já "reset_index()" é um método utilizado para redefinir o índice do DataFrame,
sendo que os parênteses vazios são para deixa-lo em sua forma padrão (tabela)'''
df_segmento_subiu = df_principal[df_principal['Resultado'] 
                                 == 'Subiu'].groupby(
'Segmento')['Var_rs'].sum().reset_index()

df_segmento_desceu = df_principal[df_principal['Resultado'] 
                                  == 'Desceu'].groupby(
    'Segmento')['Var_rs'].sum().reset_index()

df_segmento_saldo = df_principal.groupby(
    'Segmento')['Var_rs'].sum().reset_index()

df_analise_saldo = df_principal.groupby(
    'Resultado')['Var_rs'].sum().reset_index()

# Juntar os DataFrames em um único DataFrame
df_combinado = df_segmento_subiu.merge(
    df_segmento_desceu, on='Segmento', how='outer').merge(
    df_segmento_saldo, on='Segmento', how='outer')

df_combinado.fillna(0, inplace=True)  # Substituir NaN por 0
df_combinado.index += 1  # Fazer o índice começar por 1

# ↓↓↓ Gerando gráficos ↓↓↓

''' Gera um gráfico de colunas para resultado
Sendo "x" horizontal e "y" vertical as legendas,
"tittle" usado parar definir o título do gráfico, e
",.2f" é para ter apenas duas casas decimais '''
fig_resultado = px.bar(df_analise_saldo, x='Resultado', y='Var_rs',
                       text=df_analise_saldo['Var_rs'].apply(
    lambda x: f'{x:.2f}'), title='Variação R$ por Resultado',
    labels={'Var_rs': 'Variação R$'})
fig_resultado.update_layout(yaxis_tickformat=".2f")

# Gráfico de pizza/torta para segmento
fig_segmento = px.pie(df_segmento_subiu, values='Var_rs', names='Segmento',
                      title='Distribuição do Valor por Segmento',
                      labels={'Var_rs': 'Variação R$'})

# Gráfico de barras para faixa etária
fig_faixa_etaria = go.Figure(
    [go.Bar(y=df_empresas_por_faixa_etaria['Faixa_etaria'], 
            x=df_empresas_por_faixa_etaria['Qtde_de_empresas'],
            orientation='h', marker_color='gold')])

''' "orientation" é para definir a direção das barras, Sendo "h' horizontal e
"v" vertical. Já "marker_color" é para definir a cor, sendo necessário
escrever em inglês '''
fig_faixa_etaria.update_layout(xaxis=dict(title='Número de Empresas'),
                               yaxis=dict(title='Faixa Etária'),
                               title='Número de Empresas por Faixa Etária')
# "xaxis" seria para colocar a legenda na horizontal e "yaxis" na vertical

# ↓↓↓ Calculando estatísticas ↓↓↓

# Calcular as estatísticas gerais
maior = df_principal['Var_rs'].max()   # <- Calculando o maior valor
menor = df_principal['Var_rs'].min()   # <- Calculando o menor valor
media = df_principal['Var_rs'].mean()  # <- Calculando a média

# Calcular as estatísticas para as empresas que subiram
maior_subiu = df_principal[df_principal['Resultado']
                           == 'Subiu']['Var_rs'].max()
menor_subiu = df_principal[df_principal['Resultado']
                           == 'Subiu']['Var_rs'].min()
media_subiu = df_principal[df_principal['Resultado']
                           == 'Subiu']['Var_rs'].mean()

# Calcular as estatísticas para as empresas que desceram
maior_desceu = df_principal[df_principal['Resultado']
                            == 'Desceu']['Var_rs'].max()
menor_desceu = df_principal[df_principal['Resultado']
                            == 'Desceu']['Var_rs'].min()
media_desceu = df_principal[df_principal['Resultado']
                            == 'Desceu']['Var_rs'].mean()

# Formatando as estatísticas como strings
estatisticas_formatadas = [
    ["Estatísticas", "Geral", "Empresas que subiram", "Empresas que desceram"],
    ["Maior", f"R$ {maior:,.2f}", f"R$ {maior_subiu:,.2f}", 
     f"R$ {maior_desceu:,.2f}"],
    ["Menor", f"R$ {menor:,.2f}", f"R$ {menor_subiu:,.2f}", 
     f"R$ {menor_desceu:,.2f}"],
    ["Média", f"R$ {media:,.2f}", f"R$ {media_subiu:,.2f}", 
     f"R$ {media_desceu:,.2f}"]
]

dataframes = [df_principal, df_combinado, df_analise_saldo]

# Ordenar os DataFrames pelo valor da variação em ordem decrescente
for df in dataframes:
    if 'Var_rs' in df.columns:
        df.sort_values(by='Var_rs', ascending=False, inplace=True)
        # Redefinir o índice sem criar uma nova coluna de índice ↓
        df.reset_index(drop=True, inplace=True)
        df.index += 1  # <- Adicionar 1 ao índice de cada DataFrame

# ↓↓↓ Renomeando as colunas para torná-las legíveis ↓↓↓

# Definir o dicionário com os novos nomes das colunas
novos_nomes = {
    'Ativo': 'Ativo',
    'Data': 'Data',
    'Valor_final': 'Valor Final (R$)',
    'Var_dia': 'Variação do dia',
    'Var_pct': 'Variação (%)',
    'Valor_inicial': 'Valor Inicial (R$)',
    'Qtde_teorica': 'Quantidade Teórica',
    'Var_rs': 'Variação (R$)',
    'Resultado': 'Resultado',
    'Nome': 'Nome da Empresa',
    'Segmento': 'Segmento',
    'Idade': 'Idade',
    'Faixa_etaria': 'Faixa Etária'
}

# Renomear as colunas dos DataFrames
for df in dataframes:
    if 'Var_rs' in df.columns:
        df.rename(columns=novos_nomes, inplace=True)

df_empresas_por_faixa_etaria = df_empresas_por_faixa_etaria.rename(
    columns=novos_nomes)
df_empresas_por_faixa_etaria = df_empresas_por_faixa_etaria.rename(
    columns={'Qtde_de_empresas': 'Quantidade'}).copy()
df_combinado.rename(columns={'Var_rs_x': 'Subiu', 'Var_rs_y': 'Desceu', 
                             'Variação (R$)': 'Saldo'}, inplace=True)
                    

# Convertendo a coluna de data para o tipo datetime e formatando-a
df_principal['Data'] = pd.to_datetime(df_principal['Data']).dt.date

''' Tirar a notação científica das colunas que envolvem Variação (R$)
nos DataFrames ↓↓↓↓ '''

def format_float(x):
    return '{:>10,.2f}'.format(x)

for df in dataframes:
    if 'Variação (R$)' in df.columns:
        df['Variação (R$)'] = df['Variação (R$)'].apply(format_float)

def format_float(x):
    return '{:>10,.2f}'.format(x)

for df in dataframes:
    if 'Subiu' in df.columns:
        df['Subiu'] = df['Subiu'].apply(format_float)

def format_float(x):
    return '{:>10,.2f}'.format(x)

for df in dataframes:
    if 'Desceu' in df.columns:
        df['Desceu'] = df['Desceu'].apply(format_float)

def format_float(x):
    return '{:>10,.2f}'.format(x)

for df in dataframes:
    if 'Saldo' in df.columns:
        df['Saldo'] = df['Saldo'].apply(format_float)

# ↓↓↓ CENTRAL DE EXIBIÇÃO ↓↓↓

''' Obs: execute com o Run Python File.
A contagem de titulos são apenas para os DataFrames.
Coloque ou tire como comentário para ligar/desligar tais exibições e para
baixar ou não as tabelas como arquivo Excel. '''

# Exibir tabelas no terminal ↓↓

# DataFrame principal ajustado
titulo0 = "Principal:"
# Para substituir o '\n' por ele ficar bugado no ínicio da função print abaixo ↓
print()
print(tabulate([['1 ', titulo0]], tablefmt="fancy_grid"))
print(tabulate(df_principal, headers='keys', tablefmt='fancy_grid',
      numalign='center', stralign='center'), '\n')

# Estatísticas e DataFrame dos totais
print(tabulate(estatisticas_formatadas, headers="firstrow",
      tablefmt='fancy_grid', numalign='center', stralign='center'), '\n')

titulo1 = "Variação (R$) total por resultado:"
print(tabulate([[titulo1]], tablefmt="fancy_grid"))
print(tabulate(df_analise_saldo, headers='keys', tablefmt='fancy_grid',
      numalign='center', stralign='center'), '\n'*5)

# DataFrames das contagens
titulo2 = "Quantidade de empresas por:"
print(tabulate([['2 ', titulo2]], tablefmt="fancy_grid"), '\n'*2)
print(tabulate(df_contagem_resultado, headers='keys',
      tablefmt='fancy_grid', numalign='center', stralign='center'), '\n')
print(tabulate(df_contagem_segmento, headers='keys',
      tablefmt='fancy_grid', numalign='center', stralign='center'), '\n')
print(tabulate(df_empresas_por_faixa_etaria, headers='keys',
      tablefmt='fancy_grid', numalign='center', stralign='center'), '\n'*5)

# DataFrames das análises
titulo3 = "Análise por segmento (R$):"
print(tabulate([['3', titulo3]], tablefmt="fancy_grid"))
print(tabulate(df_combinado, headers='keys', tablefmt='fancy_grid',
      numalign='center', stralign='center'), '\n')

# Exibir gráficos na web ↓↓

# Gráfico de colunas
fig_resultado.show()

# Gráfico de pizza/torta
fig_segmento.show()

# Gráfico de barras
fig_faixa_etaria.show()

''' Dica: as opções como a de baixar imagem como png estão no
Canto superior direito dos gráficos '''

# # Criar um Excel com várias páginas para cada DataFrame ↓↓
# with pd.ExcelWriter('dados_analise.xlsx') as writer:
#     # index=False é para não incluir o indíce do DataFrame ↓
#     df_principal.to_excel(writer, sheet_name='Principal', index=False)
#     df_segmento_subiu.to_excel(
#         writer, sheet_name='Segmentos positivos', index=False)
#     df_segmento_desceu.to_excel(
#         writer, sheet_name='Segmentos negativos', index=False)
#     df_segmento_saldo.to_excel(
#         writer, sheet_name='Saldo dos segmentos', index=False)
#     df_analise_saldo.to_excel(
#         writer, sheet_name='Saldo das empresas', index=False)
#     df_empresas_por_faixa_etaria.to_excel(
#         writer, sheet_name='Faixa Etária', index=False)
