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
server = app.server

# Load historical spot market price data
try:
    price_dfs = {
        year: pd.read_csv(f'data/Spotmarktpreis{year}.csv', sep=';')
        for year in [2021, 2022, 2023, 2024, 2025]
    }
except FileNotFoundError as e:
    print(f"Error: {e}. Make sure all market price CSV files are in a 'data' subfolder.")
    price_dfs = {}

# Market value data
market_values_monthly = {
    2021: [5.543, 4.499, 4.105, 4.551, 4.187, 6.864, 7.409, 7.681, 11.715, 12.804, 18.307, 27.075],
    2022: [17.838, 11.871, 20.712, 14.566, 15.132, 18.940, 26.093, 39.910, 31.673, 12.904, 15.374, 24.661],
    2023: [12.291, 12.343, 8.883, 8.002, 5.356, 7.124, 5.173, 7.533, 7.447, 6.763, 8.525, 6.592],
    2024: [7.535, 5.875, 4.965, 3.795, 3.161, 4.635, 3.554, 4.263, 4.512, 6.752, 10.076, 11.171],
}
market_values_monthly[2025] = market_values_monthly[2024] # Placeholder

market_values_yearly = {2021: 7.552, 2022: 22.306, 2023: 7.2, 2024: 4.624}
market_values_yearly[2025] = market_values_yearly[2024] # Placeholder
ANZULEGENDER_WERT = 6.72

# ⚙️ NEW: Load BOTH sets of example production data
example_production_dfs_sb = {} # Südbayern
example_production_dfs_nb = {} # Nordbayern
try:
    example_production_dfs_sb = {
        "SB: 2P Tracker O-W (1.08 MWp)": pd.read_excel('data/2p-Tracker-OW-1,08MW.xlsx'),
        "SB: 2P Tracker (1.05 MWp)": pd.read_excel('data/2p-Tracker-1,05MW.xlsx'),
        "SB: Cow-PV (1.05 MWp)": pd.read_excel('data/Cow-PV-1,05MW.xlsx'),
        "SB: Cow-PV O-W (1.0 MWp)": pd.read_excel('data/Cow-PV-OW-1MW.xlsx'),
        "SB: Vertical PV (0.67 MWp)": pd.read_excel('data/Vertikale-670kW.xlsx'),
    }
    example_production_dfs_nb = {
        "NB: 2P Tracker O-W (1.15 MWp)": pd.read_excel('data/2p-Tracker-OW-1,15MW-NB.xlsx'),
        "NB: 2P Tracker (1 MWp)": pd.read_excel('data/2p-Tracker-1MW-NB.xlsx'),
        "NB: Cow-PV (1,08 MWp)": pd.read_excel('data/Cow-PV-1,08MW-NB.xlsx'),
        "NB: Cow-PV O-W (1,04 MWp)": pd.read_excel('data/Cow-PV-OW-1,04MW-NB.xlsx'),
        "NB: Vertical PV (0.609 MWp)": pd.read_excel('data/Vertikale-609kW-NB.xlsx'),
    }
except FileNotFoundError as e:
    print(f"Error loading example production files: {e}")

# =============================================================================
# Helper Functions
# =============================================================================

def preprocess_price_data(df, year):
    df_processed = df.copy()
    df_processed.columns = ['Datum', 'von', 'Zeitzone von', 'bis', 'Zeitzone bis', 'Spotmarktpreis in ct/kWh']
    df_processed['Spotmarktpreis in ct/kWh'] = df_processed['Spotmarktpreis in ct/kWh'].str.replace(',', '.').astype(float)
    df_processed["Time (CET)"] = pd.to_datetime(df_processed["Datum"] + ' ' + df_processed["von"], format='%d.%m.%Y %H:%M')
    if year in market_values_monthly:
        df_processed['Market Value Monthly (ct/kWh)'] = df_processed['Time (CET)'].dt.month.map(lambda x: market_values_monthly[year][x-1])
        df_processed['Market Value Yearly (ct/kWh)'] = market_values_yearly[year]
        df_processed['Anzulegender Wert (ct/kWh)'] = ANZULEGENDER_WERT
    return df_processed

def preprocess_pv_data(df):
    df_processed = df.copy()
    df_processed.rename(columns={df_processed.columns[0]: 'Time (UTC)', df_processed.columns[1]: 'Yield (kwH)'}, inplace=True)
    df_processed["Yield (kwH)"] = df_processed["Yield (kwH)"].apply(lambda x: max(x, 0.0))
    df_processed["Time (UTC)"] = pd.to_datetime(df_processed["Time (UTC)"])
    df_processed = df_processed.set_index("Time (UTC)").resample('h').sum().reset_index()
    df_processed["Time (CET)"] = df_processed["Time (UTC)"] + pd.Timedelta(hours=1)
    return df_processed

# Pre-process all dataframes
for year, df in price_dfs.items():
    price_dfs[year] = preprocess_price_data(df, year)
# ⚙️ NEW: Preprocess both sets
for name, df in example_production_dfs_sb.items():
    example_production_dfs_sb[name] = preprocess_pv_data(df)
for name, df in example_production_dfs_nb.items():
    example_production_dfs_nb[name] = preprocess_pv_data(df)

# =============================================================================
# Dash App Layout
# =============================================================================

app.layout = html.Div(style={'fontFamily': 'Arial, sans-serif', 'padding': '20px'}, children=[
    html.H1("Wirtschaftlichkeitsberechnung für PV-Anlagen (2021-2025 Forecast)", style={'textAlign': 'center', 'color': '#003366'}),

    # ⚙️ NEW: RadioItems to select dataset
    html.Div([
        html.P("Wählen Sie den Datensatz:", style={'fontWeight': 'bold'}),
        dcc.RadioItems(
            id='dataset-selector',
            options=[
                {'label': 'Südbayern (Standard)', 'value': 'SB'},
                {'label': 'Nordbayern (NB)', 'value': 'NB'},
            ],
            value='SB', # Default selection
            labelStyle={'display': 'inline-block', 'marginRight': '20px'}
        )
    ], style={'textAlign': 'center', 'marginBottom': '20px'}),

    html.Div([
        html.Div([
            html.P("Laden Sie eine eigene Excel-Datei (.xlsx) hoch:"),
            dcc.Upload(id='upload-data', children=html.Div(['Drag & Drop oder ', html.A('Datei auswählen')]),
                       style={'height': '60px', 'lineHeight': '60px', 'borderWidth': '2px', 'borderStyle': 'dashed', 'borderRadius': '5px', 'textAlign': 'center', 'marginBottom': '10px'}),
        ], style={'width': '48%', 'display': 'inline-block'}),

        html.Div([
            html.P("Wählen Sie ein Projekt für die Monatsansicht:"),
            # Dropdown options are now set dynamically in the callback
            dcc.Dropdown(id='project-dropdown'),
        ], style={'width': '48%', 'display': 'inline-block', 'float': 'right'})
    ], style={'marginBottom': '20px'}),

    dcc.Loading(id="loading-spinner", type="circle", children=html.Div(id='output-container'))
])

# =============================================================================
# Dash Callback
# =============================================================================

@app.callback(
    Output('output-container', 'children'),
    Output('project-dropdown', 'options'),
    Output('project-dropdown', 'value'),
    [Input('upload-data', 'contents'),
     Input('project-dropdown', 'value'),
     Input('dataset-selector', 'value')], # ⚙️ NEW: Add dataset selector input
    [State('upload-data', 'filename')]
)
def update_output(contents, selected_project, selected_dataset, filename):

    # ⚙️ NEW: Select the appropriate example dataset
    if selected_dataset == 'SB':
        active_example_dfs = example_production_dfs_sb.copy()
    else: # 'NB'
        active_example_dfs = example_production_dfs_nb.copy()

    all_production_dfs = active_example_dfs.copy()
    dropdown_options = [{'label': name, 'value': name} for name in active_example_dfs.keys()]
    
    # Check if the currently selected project belongs to the newly selected dataset
    # If not, reset the dropdown value to the first project of the new dataset
    current_selected_project = selected_project
    if selected_project not in all_production_dfs.keys():
        current_selected_project = list(all_production_dfs.keys())[0] if all_production_dfs else None


    # --- Process uploaded file if it exists ---
    uploaded_project_name = None
    if contents:
        content_type, content_string = contents.split(',')
        decoded = base64.b64decode(content_string)
        try:
            df_pv_raw = pd.read_excel(io.BytesIO(decoded))
            project_name = filename if filename else "Hochgeladene Datei"
            uploaded_project_name = project_name # Store the name
            all_production_dfs[project_name] = preprocess_pv_data(df_pv_raw)
            current_selected_project = project_name # Automatically select uploaded file
            dropdown_options = [{'label': name, 'value': name} for name in all_production_dfs.keys()] # Update options including uploaded
        except Exception as e:
            error_div = html.Div(f'Fehler beim Verarbeiten der hochgeladenen Datei: {e}', style={'color': 'red'})
            # Return error message AND original dropdown state for the active dataset
            return error_div, dropdown_options, current_selected_project

    yearly_results = []

    # --- Perform calculations for all available projects (examples + uploaded if any) ---
    for project_name, df_pv in all_production_dfs.items():
        # Historical Calculations
        for year in [2021, 2022, 2023, 2024]:
            if year not in df_pv['Time (CET)'].dt.year.unique() or year not in price_dfs: continue

            df_pv_year = df_pv[df_pv['Time (CET)'].dt.year == year]
            df_merged = pd.merge(df_pv_year, price_dfs[year], on="Time (CET)", how="inner")
            total_production = df_merged['Yield (kwH)'].sum()
            if total_production == 0: continue

            revenue_spotmarket = (df_merged[df_merged['Spotmarktpreis in ct/kWh'] >= 0]['Yield (kwH)'] * df_merged[df_merged['Spotmarktpreis in ct/kWh'] >= 0]['Spotmarktpreis in ct/kWh']).sum() / 100

            additional_revenue_yearly = 0
            if year in market_values_yearly and market_values_yearly[year] < ANZULEGENDER_WERT:
                premium = ANZULEGENDER_WERT - market_values_yearly[year]
                eligible_prod = df_merged[df_merged['Spotmarktpreis in ct/kWh'] >= 0]['Yield (kwH)'].sum()
                additional_revenue_yearly = (eligible_prod * premium) / 100

            total_revenue_mp_yearly = revenue_spotmarket + additional_revenue_yearly
            specific_revenue = (total_revenue_mp_yearly / total_production * 100).round(2)

            yearly_results.append({'Projekt': project_name, 'Jahr': str(year), 'Spez. Erlös Marktprämie (Jahr) (ct/kWh)': specific_revenue})

        # Forecast Calculation for 2025
        if 2024 in df_pv['Time (CET)'].dt.year.unique() and 2025 in price_dfs:
            df_price_2025_partial = price_dfs[2025]
            df_price_2024_end = price_dfs[2024][price_dfs[2024]['Time (CET)'].dt.month >= 10].copy()
            df_price_2024_end['Time (CET)'] += pd.DateOffset(years=1)
            df_price_2025_full = pd.concat([df_price_2025_partial, df_price_2024_end])

            df_pv_2024 = df_pv[df_pv['Time (CET)'].dt.year == 2024].copy()
            df_pv_2025 = df_pv_2024
            df_pv_2025['Time (CET)'] += pd.DateOffset(years=1)

            df_merged_2025 = pd.merge(df_pv_2025, df_price_2025_full, on="Time (CET)", how="inner")
            total_production = df_merged_2025['Yield (kwH)'].sum()
            if total_production > 0:
                revenue_spotmarket = (df_merged_2025[df_merged_2025['Spotmarktpreis in ct/kWh'] >= 0]['Yield (kwH)'] * df_merged_2025[df_merged_2025['Spotmarktpreis in ct/kWh'] >= 0]['Spotmarktpreis in ct/kWh']).sum() / 100
                additional_revenue_yearly = 0
                if market_values_yearly[2025] < ANZULEGENDER_WERT:
                    premium = ANZULEGENDER_WERT - market_values_yearly[2025]
                    eligible_prod = df_merged_2025[df_merged_2025['Spotmarktpreis in ct/kWh'] >= 0]['Yield (kwH)'].sum()
                    additional_revenue_yearly = (eligible_prod * premium) / 100
                total_revenue_mp_yearly = revenue_spotmarket + additional_revenue_yearly
                specific_revenue = (total_revenue_mp_yearly / total_production * 100).round(2)

                yearly_results.append({'Projekt': project_name, 'Jahr': '2025 (Forecast)', 'Spez. Erlös Marktprämie (Jahr) (ct/kWh)': specific_revenue})

    # --- Prepare yearly summary table ---
    summary_df = pd.DataFrame(yearly_results)
    pivot_df = pd.DataFrame() # Initialize empty
    if not summary_df.empty:
        pivot_df = summary_df.pivot(index='Projekt', columns='Jahr', values='Spez. Erlös Marktprämie (Jahr) (ct/kWh)').reset_index()
    summary_table_cols = [{"name": i, "id": i} for i in pivot_df.columns]

    # --- Prepare monthly detail table for the SELECTED project ---
    monthly_results = []
    df_pv_selected = all_production_dfs.get(current_selected_project) # Use updated selection
    if df_pv_selected is not None and 2024 in df_pv_selected['Time (CET)'].dt.year.unique() and 2025 in price_dfs:
        # Re-run forecast logic for the selected project to get monthly data
        df_price_2025_partial = price_dfs[2025]
        df_price_2024_end = price_dfs[2024][price_dfs[2024]['Time (CET)'].dt.month >= 10].copy()
        df_price_2024_end['Time (CET)'] += pd.DateOffset(years=1)
        df_price_2025_full = pd.concat([df_price_2025_partial, df_price_2024_end])

        df_pv_2024 = df_pv_selected[df_pv_selected['Time (CET)'].dt.year == 2024].copy()
        df_pv_2025 = df_pv_2024
        df_pv_2025['Time (CET)'] += pd.DateOffset(years=1)
        df_merged_2025 = pd.merge(df_pv_2025, df_price_2025_full, on="Time (CET)", how="inner")

        for month in range(1, 13):
            month_data = df_merged_2025[df_merged_2025['Time (CET)'].dt.month == month]
            total_prod = month_data['Yield (kwH)'].sum()
            prod_neg_price = month_data[month_data['Spotmarktpreis in ct/kWh'] < 0]['Yield (kwH)'].sum()

            rev_spot = 0; add_rev_yearly = 0
            if total_prod > 0:
                rev_spot = (month_data[month_data['Spotmarktpreis in ct/kWh'] >= 0]['Yield (kwH)'] * month_data[month_data['Spotmarktpreis in ct/kWh'] >= 0]['Spotmarktpreis in ct/kWh']).sum() / 100
                if market_values_yearly[2025] < ANZULEGENDER_WERT:
                    premium = ANZULEGENDER_WERT - market_values_yearly[2025]
                    eligible_prod = month_data[month_data['Spotmarktpreis in ct/kWh'] >= 0]['Yield (kwH)'].sum()
                    add_rev_yearly = (eligible_prod * premium) / 100

            total_rev_mp = rev_spot + add_rev_yearly
            spec_rev_spot = (rev_spot / total_prod * 100).round(2) if total_prod > 0 else 0
            spec_rev_mp = (total_rev_mp / total_prod * 100).round(2) if total_prod > 0 else 0

            monthly_results.append({
                'Monat': month,
                'Produktion (kWh)': f"{total_prod:.0f}",
                'Produktion bei Negativpreis (kWh)': f"{prod_neg_price:.0f}",
                'Umsatz Spotmarkt (€)': f"{rev_spot:.2f}",
                'Umsatz Marktprämie (Jahr) (€)': f"{total_rev_mp:.2f}",
                'Spez. Erlös Spotmarkt (ct/kWh)': spec_rev_spot,
                'Spez. Erlös Marktprämie (Jahr) (ct/kWh)': spec_rev_mp
            })

    monthly_df = pd.DataFrame(monthly_results)
    month_names = {1: 'Jan', 2: 'Feb', 3: 'Mar', 4: 'Apr', 5: 'May', 6: 'Jun', 7: 'Jul', 8: 'Aug', 9: 'Sep', 10: 'Oct', 11: 'Nov', 12: 'Dec'}
    monthly_df['Monat'] = monthly_df['Monat'].map(month_names)

    monthly_transposed_df = pd.DataFrame() # Initialize empty
    if not monthly_df.empty:
        monthly_transposed_df = monthly_df.set_index('Monat').T.reset_index().rename(columns={'index': 'Metrik'})
    monthly_table_cols = [{"name": i, "id": i} for i in monthly_transposed_df.columns]

    # --- Prepare Bar Chart ---
    fig = px.bar() # Initialize empty fig
    if not summary_df.empty:
        fig = px.bar(summary_df, x='Jahr', y='Spez. Erlös Marktprämie (Jahr) (ct/kWh)', color='Projekt',
                     barmode='group', title='Jahresvergleich der spezifischen Erlöse (Marktprämienmodell)',
                     text='Spez. Erlös Marktprämie (Jahr) (ct/kWh)')
        fig.update_traces(textposition='outside')
        fig.update_layout(yaxis_title="ct/kWh", xaxis_title="Jahr")

    # --- Define Output Children ---
    output_children = html.Div([
        html.H4("Jahresvergleich: Spezifischer Erlös (Marktprämie Jährlich, in ct/kWh)", style={'marginTop': '20px'}),
        dash_table.DataTable(data=pivot_df.to_dict('records'), columns=summary_table_cols, style_cell={'textAlign': 'left', 'padding': '5px'}, style_header={'fontWeight': 'bold'}, style_table={'overflowX': 'auto'}),

        dcc.Graph(id='revenue-graph', figure=fig, style={'marginTop': '30px'}),

        html.H4(f"Monatsübersicht für 2025 (Forecast): {current_selected_project}", style={'marginTop': '30px'}),
        dash_table.DataTable(data=monthly_transposed_df.to_dict('records'), columns=monthly_table_cols, style_cell={'textAlign': 'left', 'padding': '5px'}, style_header={'fontWeight': 'bold'}, style_table={'overflowX': 'auto'}),
    ])

    return output_children, dropdown_options, current_selected_project

if __name__ == '__main__':
    app.run(debug=True)