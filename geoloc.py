import openpyxl
from dash import dcc, html, Input, Output, State
import dash
import plotly.express as px
import pandas as pd

# Carregar dados do arquivo Excel
wb = openpyxl.load_workbook("rampa.xlsx", data_only=True)  # data_only=True para obter valores calculados
sheet = wb.active

# Lista de coordenadas e estados
coordenadas = []
for row in sheet.iter_rows(min_row=2):
    try:
        latitude = float(row[6].value)
        longitude = float(row[7].value)
    except (ValueError, TypeError):
        continue

    coordenadas.append({
        "LOTE": str(row[0].value),
        "UF": str(row[1].value),
        "Município": str(row[2].value),
        "Código INEP": str(row[3].value),
        "Nome da Escola": str(row[4].value),
        "Endereço": str(row[5].value),
        "Latitude": latitude,
        "Longitude": longitude,
        "Kit Wi-Fi (estimado)": str(row[8].value),
        "AP adicional (estimado)": str(row[9].value),
        "Nobreak": str(row[10].value)
    })

# Criar DataFrame com as colunas na ordem desejada
df = pd.DataFrame(coordenadas, columns=[
    "LOTE", "UF", "Município", "Código INEP", "Nome da Escola", "Endereço",
    "Latitude", "Longitude", "Kit Wi-Fi (estimado)", "AP adicional (estimado)", "Nobreak"
])

# Inicializar o aplicativo Dash
app = dash.Dash(__name__)

# Layout do aplicativo
app.layout = html.Div([
    html.H1("Mapa Interativo de Escolas"),
    dcc.Tabs(id="tabs", value='tab-1', children=[
        dcc.Tab(label='Mapa', value='tab-1'),
        dcc.Tab(label='Pesquisar Escola', value='tab-2'),
    ]),
    html.Div(id='tabs-content'),
])

@app.callback(Output('tabs-content', 'children'),
              Input('tabs', 'value'))
def render_content(tab):
    if tab == 'tab-1':
        return html.Div([
            dcc.Dropdown(
                id='estado-dropdown',
                options=[{'label': estado, 'value': estado} for estado in sorted(df['UF'].unique())],
                value=sorted(df['UF'].unique())[0],  # Valor inicial
                multi=False
            ),
            dcc.Graph(id='map'),
            html.Div(id='school-info', style={'margin-top': '20px'}),
            html.Button("Salvar Mapa como HTML", id="save-button", n_clicks=0),
            dcc.Slider(
                id='zoom-slider',
                min=1,
                max=20,
                step=1,
                value=6,  # Valor inicial
                marks={i: str(i) for i in range(1, 21)},
                tooltip={"placement": "bottom", "always_visible": True}
            )
        ])
    elif tab == 'tab-2':
        return html.Div([
            dcc.Input(id='search-input', type='text', placeholder='Digite o nome da escola...'),
            html.Button('Pesquisar', id='search-button', n_clicks=0),
            html.Div(id='search-results')
        ])

# Callback para atualizar o mapa com base no estado selecionado e zoom
@app.callback(
    Output('map', 'figure'),
    Output('school-info', 'children'),
    Input('estado-dropdown', 'value'),
    Input('zoom-slider', 'value')
)
def update_map(selected_estado, selected_zoom):
    # Filtrar dados para o estado selecionado
    filtered_df = df[df['UF'] == selected_estado]

    # Criar o mapa
    fig = px.scatter_mapbox(
        filtered_df,
        lat='Latitude',
        lon='Longitude',
        hover_name='Nome da Escola',
        hover_data={
            'Código INEP': True,
            'Endereço': True,
            'Kit Wi-Fi (estimado)': True,
            'AP adicional (estimado)': True,
            'Nobreak': True,
            'Latitude': False,  # Remover Latitude do hover
            'Longitude': False  # Remover Longitude do hover
        },
        color_discrete_sequence=["blue"],
        zoom=selected_zoom,
        height=600  # Ajustar a altura do mapa
    )
    
    fig.update_traces(marker=dict(size=8))  # Ajustar o tamanho dos marcadores

    fig.update_layout(
        mapbox_style="carto-positron",
        title=f"Escolas no Estado {selected_estado}",
        title_x=0.5,
        margin={"r":0,"t":30,"l":0,"b":0},
        legend=dict(
            yanchor="top",
            y=0.99,
            xanchor="left",
            x=0.01
        ),
        hoverlabel=dict(
            bgcolor="white",
            font_size=12,
            font_family="Arial"
        )
    )

    return fig, ""

# Callback para salvar o mapa como HTML
@app.callback(
    Output('save-button', 'children'),
    Input('save-button', 'n_clicks'),
    State('estado-dropdown', 'value'),
    State('map', 'figure')
)
def save_map(n_clicks, selected_estado, figure):
    if n_clicks > 0:
        fig = px.scatter_mapbox()
        fig.update_traces(marker=dict(size=10))
        fig.update_layout(figure['layout'])
        fig.update(data=figure['data'])
        fig.write_html(f"map_{selected_estado}.html")
        return "Mapa Salvo!"
    return "Salvar Mapa como HTML"

# Callback para pesquisar escolas
@app.callback(
    Output('search-results', 'children'),
    Input('search-button', 'n_clicks'),
    State('search-input', 'value')
)
def search_school(n_clicks, search_value):
    if n_clicks > 0 and search_value:
        results = df[df['Nome da Escola'].str.contains(search_value, case=False, na=False)]
        if not results.empty:
            return html.Ul([html.Li(f"{row['Nome da Escola']} - {row['Endereço']}") for index, row in results.iterrows()])
        else:
            return "Nenhuma escola encontrada."
    return ""

# Executar o servidor
if __name__ == '__main__':
    app.run_server(debug=True)