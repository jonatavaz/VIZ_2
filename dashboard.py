import pandas as pd
import plotly.express as px
import dash
from dash import dcc, html, Input, Output

# -----------------------------
# Carregamento e Tratamento
# -----------------------------

# Vendas
df_vendas_total = pd.concat([
    pd.read_excel("vendas2020.xlsx"),
    pd.read_excel("vendas2021.xlsx"),
    pd.read_excel("vendas2022.xlsx")
], ignore_index=True)

# Clientes
df_clientes = pd.read_excel("cadcli.xlsx")
df_clientes['Nome Completo'] = df_clientes.get('Primeiro Nome', '') + ' ' + df_clientes.get('Sobrenome', '')
df_clientes.drop(columns=['Primeiro Nome', 'Sobrenome'], errors='ignore', inplace=True)

# Produtos e Lojas
df_produtos = pd.read_excel("cadprod.xlsx")
df_lojas = pd.read_excel("cadloj.xlsx")

# Enriquecimento dos dados
df_vendas_total = df_vendas_total.merge(df_produtos[['SKU', 'Tipo do Produto', 'Preço Unitario', 'Marca']], on='SKU', how='left')
df_vendas_total['Valor Total'] = df_vendas_total['Qtd Vendida'] * df_vendas_total['Preço Unitario']
df_vendas_total = df_vendas_total.merge(df_lojas[['ID Loja', 'Nome da Loja']], on='ID Loja', how='left')
df_vendas_total = df_vendas_total.merge(df_clientes[['ID Cliente', 'Nome Completo']], on='ID Cliente', how='left')

# Datas
df_vendas_total['Data da Venda'] = pd.to_datetime(df_vendas_total['Data da Venda'])
df_vendas_total['Ano'] = df_vendas_total['Data da Venda'].dt.year
df_vendas_total['Mes'] = df_vendas_total['Data da Venda'].dt.month

# -----------------------------
# Inicialização do App
# -----------------------------

app = dash.Dash(__name__)
server = app.server  # Necessário para o Render

# -----------------------------
# Layout com Tabs
# -----------------------------

df_lojas_vendas = df_vendas_total['Nome da Loja'].value_counts().reset_index()
df_lojas_vendas.columns = ['Nome da Loja', 'Numero de Vendas']

app.layout = html.Div([
    html.H1("Dashboard de Vendas", style={'textAlign': 'center'}),
    dcc.Tabs([
        dcc.Tab(label='Vendas por Loja', children=[
            dcc.Graph(
                figure=px.bar(
                    df_lojas_vendas,
                    x='Nome da Loja',
                    y='Numero de Vendas',
                    title='Número de Vendas por Loja'
                )
            )
        ]),
        dcc.Tab(label='Top 10 Lojas por Produto', children=[
            html.Div([
                html.Label("Selecione o Tipo de Produto:"),
                dcc.Dropdown(
                    id='tipo-produto-top-lojas',
                    options=[{'label': tipo, 'value': tipo} for tipo in ['Todos'] + sorted(df_vendas_total['Tipo do Produto'].dropna().unique())],
                    value='Todos'
                ),
                dcc.Graph(id='grafico-top-lojas')
            ], style={'padding': '20px'})
        ]),
        dcc.Tab(label='Participação por Produto (Loja)', children=[
            html.Div([
                html.Label("Selecione a Loja:"),
                dcc.Dropdown(
                    id='loja-pizza',
                    options=[{'label': l, 'value': l} for l in sorted(df_vendas_total['Nome da Loja'].dropna().unique())],
                    value=sorted(df_vendas_total['Nome da Loja'].dropna().unique())[0]
                ),
                dcc.Graph(id='grafico-pizza')
            ], style={'padding': '20px'})
        ]),
        dcc.Tab(label='Top Clientes por Ano', children=[
            html.Div([
                html.Label("Selecione o Ano:"),
                dcc.Dropdown(
                    id='ano-clientes',
                    options=[{'label': str(a), 'value': a} for a in sorted(df_vendas_total['Ano'].dropna().unique())],
                    value=sorted(df_vendas_total['Ano'].dropna().unique())[0]
                ),
                dcc.Graph(id='grafico-clientes')
            ], style={'padding': '20px'})
        ]),
        dcc.Tab(label='Lojas que Mais Venderam (Marca)', children=[
            html.Div([
                html.Label("Tipo de Produto:"),
                dcc.Dropdown(
                    id='tipo-marca',
                    options=[{'label': t, 'value': t} for t in sorted(df_vendas_total['Tipo do Produto'].dropna().unique())],
                    value=sorted(df_vendas_total['Tipo do Produto'].dropna().unique())[0]
                ),
                html.Label("Marca:"),
                dcc.Dropdown(id='marca-dropdown'),
                dcc.Graph(id='grafico-marca')
            ], style={'padding': '20px'})
        ]),
        dcc.Tab(label='Evolução Anual por Marca', children=[
            html.Div([
                html.Label("Selecione o Tipo de Produto:"),
                dcc.Dropdown(
                    id='tipo-evolucao',
                    options=[{'label': t, 'value': t} for t in sorted(df_vendas_total['Tipo do Produto'].dropna().unique())],
                    value=sorted(df_vendas_total['Tipo do Produto'].dropna().unique())[0]
                ),
                dcc.Graph(id='grafico-evolucao')
            ], style={'padding': '20px'})
        ])
    ])
])

# -----------------------------
# Callbacks
# -----------------------------

@app.callback(
    Output('grafico-top-lojas', 'figure'),
    Input('tipo-produto-top-lojas', 'value')
)
def top_lojas(tipo):
    df = df_vendas_total if tipo == 'Todos' else df_vendas_total[df_vendas_total['Tipo do Produto'] == tipo]
    top = df.groupby('Nome da Loja')['Valor Total'].sum().nlargest(10).reset_index()
    return px.bar(top, x='Nome da Loja', y='Valor Total', title=f'Top 10 Lojas - {tipo}', color='Valor Total')

@app.callback(
    Output('grafico-pizza', 'figure'),
    Input('loja-pizza', 'value')
)
def pizza_loja(loja):
    df = df_vendas_total[df_vendas_total['Nome da Loja'] == loja]
    tipo_vendas = df.groupby('Tipo do Produto')['Valor Total'].sum().reset_index()
    return px.pie(tipo_vendas, names='Tipo do Produto', values='Valor Total', hole=0.4, title=f'Participação - {loja}')

@app.callback(
    Output('grafico-clientes', 'figure'),
    Input('ano-clientes', 'value')
)
def top_clientes(ano):
    df = df_vendas_total[df_vendas_total['Ano'] == ano]
    top = df.groupby('Nome Completo')['Valor Total'].sum().nlargest(10).reset_index()
    return px.bar(top, x='Valor Total', y='Nome Completo', orientation='h', title=f'Top 10 Clientes - {ano}')

@app.callback(
    Output('marca-dropdown', 'options'),
    Output('marca-dropdown', 'value'),
    Input('tipo-marca', 'value')
)
def atualizar_marcas(tipo):
    marcas = sorted(df_vendas_total[df_vendas_total['Tipo do Produto'] == tipo]['Marca'].dropna().unique())
    options = [{'label': m, 'value': m} for m in marcas]
    return options, marcas[0] if marcas else None

@app.callback(
    Output('grafico-marca', 'figure'),
    Input('tipo-marca', 'value'),
    Input('marca-dropdown', 'value')
)
def lojas_mais_venderam(tipo, marca):
    df = df_vendas_total[(df_vendas_total['Tipo do Produto'] == tipo) & (df_vendas_total['Marca'] == marca)]
    top = df.groupby('Nome da Loja')['Valor Total'].sum().reset_index().sort_values(by='Valor Total', ascending=False)
    return px.bar(top, x='Valor Total', y='Nome da Loja', orientation='h', title=f'Lojas - {tipo}/{marca}')

@app.callback(
    Output('grafico-evolucao', 'figure'),
    Input('tipo-evolucao', 'value')
)
def evolucao_tipo(tipo):
    df = df_vendas_total[df_vendas_total['Tipo do Produto'] == tipo]
    evolucao = df.groupby(['Ano', 'Marca'])['Valor Total'].sum().reset_index()
    return px.line(evolucao, x='Ano', y='Valor Total', color='Marca', title=f'Evolução por Marca - {tipo}', markers=True)

# -----------------------------
# Run
# -----------------------------

if __name__ == '__main__':
    app.run(debug=True)  # Para desenvolvimento local
