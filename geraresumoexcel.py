import pandas as pd

# Carregar o arquivo Excel
arquivo_excel = r'C:\pythonjr\DRP-DRC\dados_volts_127_220.xlsx'
df = pd.read_excel(arquivo_excel)

# Definição dos limites das faixas de tensão
limite1_220 = 191
limite2_220 = 202
limite3_220 = 231
limite4_220 = 233
limite1_127 = 110
limite2_127 = 117
limite3_127 = 133
limite4_127 = 135

# Definição dos limites para DRP e DRC
limDRP = 0.03
limDRC = 0.005

# Calcular os valores médio, máximo e mínimo para cada coluna
medias = df.mean()
maximos = df.max()
minimos = df.min()

# Total de registros por coluna
total_registros = len(df)

# Resultados por coluna
resultados = {}

# Inicialização de contadores para o resumo final
qtdemedicoes = 0
contagem_drp = 0
contagem_drc = 0
contagem_ambos = 0

# Analisar e contar valores para cada coluna
for coluna in df.columns:
    media_coluna = medias[coluna]
    maximo_coluna = maximos[coluna]
    minimo_coluna = minimos[coluna]
    contagem = {'n_criticos': 0, 'n_precarios': 0}

    for valor in df[coluna]:
        if media_coluna > 150:
            if valor < limite1_220 or valor > limite4_220:
                contagem['n_criticos'] += 1
            elif (limite1_220 <= valor < limite2_220) or (limite3_220 < valor <= limite4_220):
                contagem['n_precarios'] += 1
        else:
            if valor < limite1_127 or valor > limite4_127:
                contagem['n_criticos'] += 1
            elif (limite1_127 <= valor < limite2_127) or (limite3_127 < valor <= limite4_127):
                contagem['n_precarios'] += 1

    drp = (contagem['n_precarios'] / total_registros) * 100
    drc = (contagem['n_criticos'] / total_registros) * 100
    resultados[coluna] = {
        'DRP': drp,
        'DRC': drc,
        'Media': media_coluna,
        'Max': maximo_coluna,
        'Min': minimo_coluna
    }

    # Contador da quantidade total de medições
    qtdemedicoes += 1

    # Contagem para o resumo final
    if drp > limDRP and drc > limDRC:
        contagem_ambos += 1
    elif drp > limDRP:
        contagem_drp += 1
    elif drc > limDRC:
        contagem_drc += 1

# Converter os resultados em um DataFrame para exportação
df_resultados = pd.DataFrame.from_dict(resultados, orient='index')
df_resultados.reset_index(inplace=True)
df_resultados.rename(columns={'index': 'Coluna'}, inplace=True)

# Salvando o DataFrame no arquivo Excel
caminho_arquivo_saida = r'C:\pythonjr\DRP-DRC\resultados_analise.xlsx'
df_resultados.to_excel(caminho_arquivo_saida, index=False)

print('Os resultados foram salvos com sucesso no arquivo Excel:', caminho_arquivo_saida)
