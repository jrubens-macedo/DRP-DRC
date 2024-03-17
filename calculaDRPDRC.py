import pandas as pd

# Carregar o arquivo Excel
arquivo_excel = r'C:\pythonjr\drpdrc\dados_volts_127_220.xlsx'
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
    valores_coluna = df[coluna]
    maximo_coluna = maximos[coluna]
    minimo_coluna = minimos[coluna]
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
    resultados[coluna] = {'DRP': drp, 'DRC': drc, 'Max': maximo_coluna, 'Min': minimo_coluna}

    # Contador da quantidade total de medições
    qtdemedicoes += 1

    # Contagem para o resumo final
    if drp > limDRP and drc > limDRC:
        contagem_ambos += 1
    elif drp > limDRP:
        contagem_drp += 1
    elif drc > limDRC:
        contagem_drc += 1

# Exibir os valores de DRP e DRC para cada coluna
for coluna, indicadores in resultados.items():
    print(f"{coluna}: DRP = {indicadores['DRP']:.2f}%, DRC = {indicadores['DRC']:.2f}%, Max = {indicadores['Max']}, Min = {indicadores['Min']}")

percdrp = (contagem_drp/qtdemedicoes)*100
percdrc = (contagem_drc/qtdemedicoes)*100
percdrpdrc = (contagem_ambos/qtdemedicoes)*100

# Exibir o resumo final
print('-----------------------------------------------------------------')
print(f' RESUMO FINAL')
print('-----------------------------------------------------------------')
print(f" Total de medições analisadas: {qtdemedicoes}")
print(f" Quantidade de medições que violaram apenas DRP: {contagem_drp} ({percdrp:.2f}%)")
print(f" Quantidade de medições que violaram apenas DRC: {contagem_drc} ({percdrc:.2f}%)")
print(f" Quantidade de medições que violaram DRP e DRC: {contagem_ambos} ({percdrpdrc:.2f}%)")
print('-----------------------------------------------------------------')


