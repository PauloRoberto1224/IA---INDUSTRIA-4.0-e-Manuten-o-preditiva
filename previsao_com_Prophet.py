import pandas as pd
from prophet import Prophet
from sklearn.ensemble import RandomForestRegressor
from datetime import timedelta
import plotly.express as px

# Ajuste o caminho do arquivo para o seu sistema local
arquivo_excel = 'C:/Users/suinfor/Documents/Angular/UNDB/SGI - CORREIAS (ESPESSURA) 1.xlsx'

# Carrega a planilha Excel localizada no caminho fornecido
df = pd.read_excel(arquivo_excel)

# Filtrando os dados para o equipamento EP-313K-02
df_equipamento = df[df['EQUIPAMENTO'] == 'EP-313K-02'].copy()

# Convertendo as datas de inspeção para formato datetime
df_equipamento['DATA INSPEÇÃO'] = pd.to_datetime(df_equipamento['DATA INSPEÇÃO'], errors='coerce')

# Incluindo uma variável de exemplo 'temperatura' (pode ser substituída por dados reais)
df_equipamento['TEMPERATURA'] = 30 + (df_equipamento.index % 5)  # Exemplo de variável auxiliar

# Mantemos apenas as colunas necessárias para a análise
df_equipamento = df_equipamento[['DATA INSPEÇÃO', 'ESPESSURA MÍNIMA', 'TEMPERATURA']]

# Removendo valores nulos
df_equipamento.dropna(subset=['DATA INSPEÇÃO', 'ESPESSURA MÍNIMA'], inplace=True)

# Renomeando as colunas para usar no Prophet
df_equipamento.columns = ['ds', 'y', 'temperatura']

# Exibir os últimos registros e o tamanho do DataFrame após o filtro
print("Últimos dados após a filtragem e remoção de nulos:")
print(df_equipamento.tail())
print(f"Número de linhas válidas: {len(df_equipamento)}")

# Detecção de Outliers - filtrando espessuras negativas ou muito discrepantes
df_equipamento = df_equipamento[(df_equipamento['y'] > 0) & (df_equipamento['y'] <= 16)]

# Obter a última inserção de dados e seu valor
ultima_data_insercao = df_equipamento['ds'].max()
ultima_espessura_insercao = df_equipamento.loc[df_equipamento['ds'] == ultima_data_insercao, 'y'].values[0]
print(f"Última inserção de dados: {ultima_data_insercao.strftime('%d/%m/%Y')} com espessura de {ultima_espessura_insercao} mm")

# Divisão em treino e teste
prediction_size = 5
train_df = df_equipamento[:-prediction_size]

# Verificando se há pelo menos 2 linhas válidas após a limpeza
if len(train_df) < 2:
    raise ValueError("O conjunto de dados possui menos de 2 linhas válidas para o treinamento.")

# Instanciando o modelo Prophet e adicionando o regressor 'temperatura'
b = Prophet()
b.add_regressor('temperatura')

# Treinando o modelo
b.fit(train_df)

# Fazendo previsões para os próximos 10 meses (aproximadamente 300 dias)
future = b.make_future_dataframe(periods=300)
future['temperatura'] = 30  # Você pode ajustar ou coletar dados reais de temperatura

# Realizando a previsão
forecast = b.predict(future)

# Exibindo os primeiros dados da previsão
print(forecast[['ds', 'yhat', 'yhat_lower', 'yhat_upper']].head())

# RandomForest - Usando a previsão do Prophet como input
X_train = train_df[['ds', 'temperatura']].copy()
X_train['ds'] = X_train['ds'].map(pd.Timestamp.timestamp)
y_train = train_df['y']

# Usando previsões do Prophet como entrada
X_future = forecast[['ds', 'temperatura']].copy()
X_future['ds'] = X_future['ds'].map(pd.Timestamp.timestamp)

# Instanciando e treinando o modelo de Random Forest
rf_model = RandomForestRegressor()
rf_model.fit(X_train, y_train)

# Prevendo com o modelo RandomForest
rf_predictions = rf_model.predict(X_future[['ds', 'temperatura']])

# Agora podemos instanciar o cálculo para garantir que a espessura mínima na previsão não seja menor que 2mm
espessura_minima = 2.0

# Garantindo que a previsão da troca seja **após** a última data disponível nos dados
ultima_data_medicao = df_equipamento['ds'].max()

# Filtrando previsões após a última data da medição, garantindo que a espessura prevista seja 2 mm
previsao_troca = forecast[(forecast['yhat'] <= espessura_minima) & (forecast['yhat'] >= 2.0) & (forecast['ds'] > ultima_data_medicao)].head(1)

# Se não houver previsão de troca, exibir ainda assim o desgaste estimado
if not previsao_troca.empty:
    data_proxima_troca = previsao_troca['ds'].values[0]
    espessura_atual = previsao_troca['yhat'].values[0]
else:
    # Se não houver previsão de troca, usa a última previsão disponível
    data_proxima_troca = forecast['ds'].max()
    espessura_atual = forecast.loc[forecast['ds'] == data_proxima_troca, 'yhat'].values[0]

# Calcula o desgaste mensal e diário estimado
dias_entre_medicoes = (pd.to_datetime(data_proxima_troca) - ultima_data_medicao).days
desgaste_total = ultima_espessura_insercao - espessura_minima

# Desgaste diário e mensal estimado
desgaste_diario = desgaste_total / dias_entre_medicoes if dias_entre_medicoes > 0 else 0
desgaste_mensal = desgaste_diario * 30

# Cria uma tabela com todas as informações
dados_analise = {
    "Última Inserção de Dados": [ultima_data_insercao.strftime('%d/%m/%Y')],
    "Espessura na Última Inserção (mm)": [ultima_espessura_insercao],
    "Data da Previsão de Troca": [pd.to_datetime(data_proxima_troca).strftime('%d/%m/%Y')],
    "Espessura Prevista na Troca (mm)": [round(espessura_atual, 2)],
    "Desgaste Estimado Total (mm)": [round(desgaste_total, 2)],
    "Desgaste Diário (mm/dia)": [round(desgaste_diario, 3)],
    "Desgaste Mensal (mm/mês)": [round(desgaste_mensal, 2)]
}

tabela_analise = pd.DataFrame(dados_analise)

# Exibe a tabela de análise
print("Previsão da próxima troca e análise de desgaste:")
print(tabela_analise)

# Salvando a previsão completa e a tabela de análise em arquivos Excel para análise futura
forecast.to_excel('Previsao_Proxima_Troca_EP-313K-02.xlsx', index=False)
tabela_analise.to_excel('Analise_Desgaste_EP-313K-02.xlsx', index=False)

# Visualizando os dados com Plotly
fig = px.line(forecast, x='ds', y='yhat', title="Previsão de Espessura ao Longo do Tempo")
fig.add_scatter(x=forecast['ds'], y=rf_predictions, mode='lines', name='Random Forest Predictions')
fig.show()
