import pandas as pd

# Carregar o arquivo Excel
arquivo_excel = r'C:\pythonjr\drpdrc\dados_volts_127_220.xlsx'
df = pd.read_excel(arquivo_excel)

# Definição dos limites
limite1_220 = 191
limite2_220 = 202
limite3_220 = 231
limite4_220 = 233
limite1_127 = 110
limite2_127 = 117
limite3_127 = 133
limite4_127 = 135

# Calcular a média para cada coluna
medias = df.mean()

# Total de registros por coluna
total_registros = len(df)

# Resultados por coluna
resultados = {}

# Analisar e contar valores para cada coluna
for coluna in df.columns:
    media_coluna = medias[coluna]
    contagem = {'n_criticos': 0, 'n_precarios': 0}

    for valor in df[coluna]:
        if media_coluna > 150:
            if valor < limite1_220 or valor > limite4_220:
                contagem['n_criticos'] += 1
            if (limite1_220 <= valor < limite2_220) or (limite3_220 < valor <= limite4_220):
                contagem['n_precarios'] += 1
        else:
            if valor < limite1_127 or valor > limite4_127:
                contagem['n_criticos'] += 1
            if (limite1_127 <= valor < limite2_127) or (limite3_127 < valor <= limite4_127):
                contagem['n_precarios'] += 1

    drp = (contagem['n_precarios'] / total_registros) * 100
    drc = (contagem['n_criticos'] / total_registros) * 100
    resultados[coluna] = {'DRP': drp, 'DRC': drc}

# Exibir os valores de DRP e DRC para cada coluna
for coluna, indicadores in resultados.items():
    print(f"{coluna}: DRP = {indicadores['DRP']:.2f}%, DRC = {indicadores['DRC']:.2f}%")



