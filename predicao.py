import pandas as pd
from datetime import datetime, timedelta

# Carrega a planilha Excel localizada na mesma pasta que o código
arquivo_excel = 'SGI - CORREIAS (ESPESSURA) 1.xlsx'
df = pd.read_excel(arquivo_excel)

# Define a espessura mínima necessária para a troca da peça
espessura_minima = 2.0  # Espessura mínima em milímetros

# Filtra os dados para o equipamento específico AL-313K-02
df_equipamento = df[df['EQUIPAMENTO'] == 'AL-313K-02'].copy()

# Converte as datas de inspeção para o formato datetime
df_equipamento['DATA INSPEÇÃO'] = pd.to_datetime(df_equipamento['DATA INSPEÇÃO'], errors='coerce')

# Ordena os dados pela data de inspeção, da mais antiga para a mais recente
df_equipamento = df_equipamento.sort_values(by='DATA INSPEÇÃO')

# Considera sempre a maior espessura registrada em cada data de inspeção
df_equipamento = df_equipamento.groupby('DATA INSPEÇÃO').agg({'ESPESSURA MÍNIMA': 'max'}).reset_index()

# Lista para armazenar os resultados das trocas e medições
resultados = []

# Variáveis para controlar as trocas e medições
data_troca = None
espessura_troca = None
ultima_data_medicao = None
ultima_espessura_medicao = None

# Itera sobre as linhas dos dados filtrados
for index, row in df_equipamento.iterrows():
    espessura = row['ESPESSURA MÍNIMA']
    
    # Se a espessura estiver entre 10.00 e 18.00 mm, indica que houve uma troca da peça
    if 10.00 <= espessura <= 18.00:
        if data_troca is not None and espessura != espessura_troca:
            # Calcula o tempo desde a última troca e o desgaste ocorrido
            dias_desde_troca = (row['DATA INSPEÇÃO'] - data_troca).days
            desgaste_ocorrido = espessura_troca - espessura
            desgaste_diario_medio = desgaste_ocorrido / dias_desde_troca if dias_desde_troca > 0 else 0
            
            # Armazena os resultados dessa troca
            resultados.append({
                "Data da Troca": data_troca.strftime('%d/%m/%Y'),
                "Espessura na Troca (mm)": espessura_troca,
                "Data Atual": row['DATA INSPEÇÃO'].strftime('%d/%m/%Y'),
                "Espessura Atual (mm)": espessura,
                "Dias desde a última troca": dias_desde_troca,
                "Desgaste Ocorrido (mm)": round(desgaste_ocorrido, 3),
                "Desgaste Médio Diário (mm/dia)": round(desgaste_diario_medio, 3),
                "Data da Última Medição Antes da Troca": ultima_data_medicao.strftime('%d/%m/%Y') if ultima_data_medicao else None,
                "Espessura da Última Medição Antes da Troca (mm)": ultima_espessura_medicao,
                "Tipo de Registro": "Troca"
            })
        
        # Atualiza as variáveis com a data e espessura da troca
        data_troca = row['DATA INSPEÇÃO']
        espessura_troca = espessura
    
    # Se a espessura for menor que 10.00 mm, é considerada uma medição
    elif espessura < 10.00 and data_troca is not None and espessura != ultima_espessura_medicao:
        dias_desde_troca = (row['DATA INSPEÇÃO'] - data_troca).days
        desgaste_ocorrido = espessura_troca - espessura
        desgaste_diario_medio = desgaste_ocorrido / dias_desde_troca if dias_desde_troca > 0 else 0
        
        # Armazena os resultados dessa medição
        resultados.append({
            "Data da Troca": data_troca.strftime('%d/%m/%Y'),
            "Espessura na Troca (mm)": espessura_troca,
            "Data Atual": row['DATA INSPEÇÃO'].strftime('%d/%m/%Y'),
            "Espessura Atual (mm)": espessura,
            "Dias desde a última troca": dias_desde_troca,
            "Desgaste Ocorrido (mm)": round(desgaste_ocorrido, 3),
            "Desgaste Médio Diário (mm/dia)": round(desgaste_diario_medio, 3),
            "Data da Última Medição Antes da Troca": ultima_data_medicao.strftime('%d/%m/%Y') if ultima_data_medicao else None,
            "Espessura da Última Medição Antes da Troca (mm)": ultima_espessura_medicao,
            "Tipo de Registro": "Medição"
        })

        # Atualiza as variáveis com a data e espessura da última medição
        ultima_data_medicao = row['DATA INSPEÇÃO']
        ultima_espessura_medicao = espessura

# Converte a lista de resultados em um DataFrame
resultados_df = pd.DataFrame(resultados)

# Realiza a previsão a partir da última medição
if not resultados_df.empty:
    ultima_medicao = resultados_df[resultados_df['Tipo de Registro'] == 'Medição'].iloc[-1]
    espessura_atual = ultima_medicao['Espessura Atual (mm)']
    data_ultima_medicao = pd.to_datetime(ultima_medicao['Data Atual'], format='%d/%m/%Y')
    desgaste_medio_diario = ultima_medicao['Desgaste Médio Diário (mm/dia)']
    
    # Prever a data da próxima troca com base no desgaste médio diário
    dias_restantes = (espessura_atual - espessura_minima) / desgaste_medio_diario if desgaste_medio_diario > 0 else float('inf')
    data_proxima_troca = data_ultima_medicao + timedelta(days=dias_restantes)
    
    # Previsão da espessura na próxima troca
    espessura_prevista_na_troca = espessura_minima  # A espessura prevista na troca será a espessura mínima
    
    # Cria o DataFrame com a previsão da próxima troca
    previsao_df = pd.DataFrame([{
        "Data da Troca": data_proxima_troca.strftime('%d/%m/%Y'),
        "Espessura na Troca (mm)": espessura_prevista_na_troca,
        "Data Atual": data_proxima_troca.strftime('%d/%m/%Y'),
        "Espessura Atual (mm)": espessura_prevista_na_troca,
        "Dias desde a última troca": round(dias_restantes, 0),
        "Desgaste Ocorrido (mm)": round(espessura_atual - espessura_minima, 3),
        "Desgaste Médio Diário (mm/dia)": round(desgaste_medio_diario, 3),
        "Data da Última Medição Antes da Troca": ultima_medicao['Data Atual'],
        "Espessura da Última Medição Antes da Troca (mm)": espessura_atual,
        "Tipo de Registro": "Previsão"
    }])
    
    # Adiciona a previsão ao DataFrame original de resultados
    resultados_df = pd.concat([resultados_df, previsao_df], ignore_index=True)

# Assegura que todas as colunas de datas estão no formato datetime para ordenação correta
resultados_df['Data Atual'] = pd.to_datetime(resultados_df['Data Atual'], format='%d/%m/%Y')
resultados_df['Data da Troca'] = pd.to_datetime(resultados_df['Data da Troca'], format='%d/%m/%Y')

# Organiza os resultados por tipo de registro e pela data de inspeção
resultados_df = resultados_df.sort_values(by='Data Atual').reset_index(drop=True)

# Exibe o DataFrame final com a previsão
print(resultados_df)

# Salva os resultados em uma nova planilha Excel
resultados_df.to_excel('Resultados_Desgaste_AL-313K-02.xlsx', index=False)


##LIMPAR OS DADOS
## CRIAR O PLOT
## VARIAÇÃO DE MEDIÇÕES ( ARRUMAR)