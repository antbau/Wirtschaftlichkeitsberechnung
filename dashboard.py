import base64
import io
import pandas as pd
import dash
from dash import dcc, html, dash_table
from dash.dependencies import Input, Output, State
import plotly.express as px

# =============================================================================
# Initial Setup & Data Loading
# =============================================================================

# Initialize the Dash app
app = dash.Dash(__name__)
app.title = "PV Revenue Calculator"

# Load historical spot market price data
try:
    df_2021 = pd.read_csv('data/Spotmarktpreis2021.csv', sep=';')
    df_2022 = pd.read_csv('data/Spotmarktpreis2022.csv', sep=';')
    df_2023 = pd.read_csv('data/Spotmarktpreis2023.csv', sep=';')
    df_2024 = pd.read_csv('data/Spotmarktpreis2024.csv', sep=';')
    price_dfs = {2021: df_2021, 2022: df_2022, 2023: df_2023, 2024: df_2024}
except FileNotFoundError:
    print("Error: Market price CSV files not found. Make sure they are in a 'data' subfolder.")
    price_dfs = {}

# Market value data
mv_2021 = [5.543, 4.499, 4.105, 4.551, 4.187, 6.864, 7.409, 7.681, 11.715, 12.804, 18.307, 27.075]
mv_2022 = [17.838, 11.871, 20.712, 14.566, 15.132, 18.940, 26.093, 39.910, 31.673, 12.904, 15.374, 24.661]
mv_2023 = [12.291, 12.343, 8.883, 8.002, 5.356, 7.124, 5.173, 7.533, 7.447, 6.763, 8.525, 6.592]
mv_2024 = [7.535, 5.875, 4.965, 3.795, 3.161, 4.635, 3.554, 4.263, 4.512, 6.752, 10.076, 11.171]
market_values = {2021: mv_2021, 2022: mv_2022, 2023: mv_2023, 2024: mv_2024}
ANZULEGENDER_WERT = 6.72

# =============================================================================
# Helper Functions
# =============================================================================

def preprocess_price_data(df, year):
    """Cleans and prepares the historical price data for a given year."""
    df_processed = df.copy()
    expected_columns = ['Datum', 'von', 'Zeitzone von', 'bis', 'Zeitzone bis', 'Spotmarktpreis in ct/kWh']
    df_processed.columns = expected_columns
    df_processed['Spotmarktpreis in ct/kWh'] = df_processed['Spotmarktpreis in ct/kWh'].str.replace(',', '.').astype(float)
    df_processed["Time (CET)"] = pd.to_datetime(df_processed["Datum"] + ' ' + df_processed["von"], format='%d.%m.%Y %H:%M')
    df_processed['Market Value (ct/kWh)'] = df_processed['Time (CET)'].dt.month.map(lambda x: market_values[year][x-1])
    df_processed['Anzulegender Wert (ct/kWh)'] = ANZULEGENDER_WERT
    return df_processed

def preprocess_pv_data(df):
    """Cleans and prepares the uploaded PV production data."""
    df_processed = df.copy()
    df_processed.rename(columns={df_processed.columns[0]: 'Time (UTC)', df_processed.columns[1]: 'Yield (kwH)'}, inplace=True)
    df_processed["Yield (kwH)"] = df_processed["Yield (kwH)"].apply(lambda x: max(x, 0.0))
    df_processed["Time (UTC)"] = pd.to_datetime(df_processed["Time (UTC)"])
    df_processed.set_index("Time (UTC)", inplace=True)
    df_processed = df_processed.resample('h').sum().reset_index()
    df_processed["Time (CET)"] = df_processed["Time (UTC)"] + pd.Timedelta(hours=1)
    # Filter for the relevant years
    df_processed = df_processed[df_processed['Time (CET)'].dt.year.isin([2021, 2022, 2023, 2024])]
    return df_processed

# Pre-process all loaded price dataframes
for year, df in price_dfs.items():
    price_dfs[year] = preprocess_price_data(df, year)

# =============================================================================
# Dash App Layout & Callback
# =============================================================================

app.layout = html.Div(style={'fontFamily': 'Arial, sans-serif', 'padding': '20px'}, children=[
    html.H1("Wirtschaftlichkeitsberechnung für PV-Anlagen (2021-2024)", style={'textAlign': 'center', 'color': '#003366'}),
    html.P("Laden Sie eine Excel-Datei (.xlsx) mit dem Produktionsprofil Ihrer PV-Anlage hoch.", style={'textAlign': 'center'}),
    html.P("Die Datei muss Daten für die Jahre 2021-2024 enthalten. Andere Jahre werden ignoriert.", style={'textAlign': 'center', 'fontSize': '14px'}),

    dcc.Upload(
        id='upload-data',
        children=html.Div(['Ziehen Sie Ihre Datei per Drag & Drop hierher oder ', html.A('wählen Sie eine Datei aus')]),
        style={
            'width': '100%', 'height': '60px', 'lineHeight': '60px',
            'borderWidth': '2px', 'borderStyle': 'dashed', 'borderRadius': '5px',
            'textAlign': 'center', 'margin': '20px 0'
        },
        multiple=False
    ),
    html.Div(id='output-container')
])

@app.callback(
    Output('output-container', 'children'),
    Input('upload-data', 'contents'),
    State('upload-data', 'filename')
)
def update_output(contents, filename):
    if contents is None:
        return html.Div('Bitte laden Sie eine Datei hoch, um die Berechnung zu starten.')

    content_type, content_string = contents.split(',')
    decoded = base64.b64decode(content_string)

    try:
        df_pv_raw = pd.read_excel(io.BytesIO(decoded))
        df_pv = preprocess_pv_data(df_pv_raw)

        if df_pv.empty:
            return html.Div("Fehler: Die hochgeladene Datei enthält keine Produktionsdaten für die Jahre 2021-2024.", style={'color': 'red'})

        yearly_results = []
        
        for year in sorted(df_pv['Time (CET)'].dt.year.unique()):
            if year not in price_dfs:
                continue # Skip if no price data is available for the year

            df_pv_year = df_pv[df_pv['Time (CET)'].dt.year == year]
            df_price_year = price_dfs[year]

            df_merged = pd.merge(df_pv_year, df_price_year, on="Time (CET)", how="inner")
            
            total_production = df_merged['Yield (kwH)'].sum()
            if total_production == 0:
                continue

            # --- Calculations for the year ---
            revenue_spotmarket = (df_merged['Yield (kwH)'] * df_merged['Spotmarktpreis in ct/kWh']).sum() / 100
            specific_revenue_spotmarket = (revenue_spotmarket / total_production * 100).round(2)

            revenue_marketvalue = (df_merged['Yield (kwH)'] * df_merged['Market Value (ct/kWh)']).sum() / 100
            specific_revenue_marketvalue = (revenue_marketvalue / total_production * 100).round(2)
            
            df_merged['Additional Revenue'] = df_merged['Yield (kwH)'] * (df_merged['Anzulegender Wert (ct/kWh)'] - df_merged['Spotmarktpreis in ct/kWh']) / 100
            df_merged.loc[df_merged['Spotmarktpreis in ct/kWh'] < 0, 'Additional Revenue'] = 0.0
            df_merged['Additional Revenue'] = df_merged['Additional Revenue'].apply(lambda x: max(x, 0.0))
            total_revenue_mp = revenue_spotmarket + df_merged['Additional Revenue'].sum()
            specific_revenue_mp = (total_revenue_mp / total_production * 100).round(2)
            
            yearly_results.append({
                'Jahr': year,
                'Produktion (kWh)': f"{total_production:,.0f}",
                'Spez. Erlös Spotmarkt (ct/kWh)': specific_revenue_spotmarket,
                'Spez. Erlös Marktwert (ct/kWh)': specific_revenue_marketvalue,
                'Spez. Erlös Marktprämie (ct/kWh)': specific_revenue_mp,
            })

        if not yearly_results:
             return html.Div("Fehler: Konnte keine Ergebnisse für die Jahre 2021-2024 berechnen.", style={'color': 'red'})

        # --- Create Visualization and Table ---
        results_df = pd.DataFrame(yearly_results)
        
        # Melt the DataFrame for easy plotting with Plotly Express
        plot_df = results_df.melt(
            id_vars=['Jahr'], 
            value_vars=['Spez. Erlös Spotmarkt (ct/kWh)', 'Spez. Erlös Marktwert (ct/kWh)', 'Spez. Erlös Marktprämie (ct/kWh)'],
            var_name='Modell', 
            value_name='Spezifischer Erlös (ct/kWh)'
        )
        
        fig = px.bar(
            plot_df, 
            x='Jahr', 
            y='Spezifischer Erlös (ct/kWh)', 
            color='Modell', 
            barmode='group',
            title='Vergleich der spezifischen PV-Erlöse pro Jahr',
            text='Spezifischer Erlös (ct/kWh)'
        )
        fig.update_traces(textposition='outside')
        fig.update_layout(yaxis_title="ct/kWh", xaxis_title="Jahr")
        
        return html.Div([
            html.H3(f'Ergebnisse für: {filename}', style={'textAlign': 'center'}),
            html.Hr(),
            html.H4("Jahresübersicht der spezifischen Erlöse:", style={'marginTop': '20px'}),
            dash_table.DataTable(
                data=results_df.to_dict('records'),
                columns=[{'name': i, 'id': i} for i in results_df.columns],
                style_cell={'textAlign': 'left'},
                style_header={'fontWeight': 'bold'},
            ),
            dcc.Graph(id='revenue-graph', figure=fig, style={'marginTop': '30px'}),
        ])

    except Exception as e:
        print(e)
        return html.Div(f'Beim Verarbeiten der Datei ist ein Fehler aufgetreten. Stellen Sie sicher, dass es sich um eine .xlsx-Datei mit dem richtigen Format handelt. Fehler: {e}', style={'color': 'red'})

if __name__ == '__main__':
    app.run(debug=True)