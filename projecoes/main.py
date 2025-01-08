import pandas as pd
from statsmodels.tsa.statespace.sarimax import SARIMAX

def projecao(caminho, campos_a_serem_projetados):
    # Caminho
    data = pd.read_excel(caminho)

    # Converter ano e mês para um índice datetime
    data['date'] = pd.to_datetime(data['ano'].astype(str) + '-' + data['mes'].astype(str) + '-01')
    data.set_index('date', inplace=True)
    
    # Dicionário para armazenar os dataframes
    sheets = {}

    for campo in campos_a_serem_projetados:
        data_series = data[campo]

        # Define e ajusta o modelo SARIMA para o campo atual
        model = SARIMAX(data_series, order=(1, 1, 1), seasonal_order=(1, 1, 1, 12), enforce_stationarity=False, enforce_invertibility=False)
        results = model.fit(disp=False)

        # Realiza a projeção para alcançar de outubro de 2024 até dezembro de 2025 (15 meses)
        forecast = results.get_forecast(steps=15)
        forecast_values = forecast.predicted_mean[-15:]  # Seleciona os 15 meses de outubro de 2024 a dezembro de 2025
        forecast_ci = forecast.conf_int().iloc[-15:]

        # Cria o DataFrame com os resultados para o campo atual
        forecast_df = pd.DataFrame({
            'mes': pd.date_range(start='2024-10-01', periods=15, freq='MS').strftime('%m/%Y'),
            f'{campo}_projecao': forecast_values.values,
            f'{campo}_IC_inferior': forecast_ci.iloc[:, 0].values,
            f'{campo}_IC_superior': forecast_ci.iloc[:, 1].values
        })

        # Formatar números no padrão brasileiro
        forecast_df[f'{campo}_projecao'] = forecast_df[f'{campo}_projecao'].map('{:,.2f}'.format).str.replace(',', 'X').str.replace('.', ',').str.replace('X', '.')
        forecast_df[f'{campo}_IC_inferior'] = forecast_df[f'{campo}_IC_inferior'].map('{:,.2f}'.format).str.replace(',', 'X').str.replace('.', ',').str.replace('X', '.')
        forecast_df[f'{campo}_IC_superior'] = forecast_df[f'{campo}_IC_superior'].map('{:,.2f}'.format).str.replace(',', 'X').str.replace('.', ',').str.replace('X', '.')

        # Adiciona o dataframe ao dicionário de planilhas
        sheets[campo] = forecast_df

    # Salva todas as projeções em diferentes abas em um arquivo Excel
    with pd.ExcelWriter("projecao_outubro2024_dezembro2025.xlsx") as writer:
        for campo, df in sheets.items():
            df.to_excel(writer, sheet_name=campo, index=False)
    
    print("Projeções realizadas de outubro de 2024 até dezembro de 2025!")

if __name__ == "__main__":
    caminho = "dados/embrapii_dados_projecao.xlsx"
    campos = ["ctt_projetos", "ctt_valor", "conc_projetos", "ctt_valor_embrapii"]
    projecao(caminho, campos)
