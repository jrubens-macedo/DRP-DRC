# Cálculo dos indicadores de qualidade de tensão em regime permanente
# DRP, DRC, P99% e P1%
# By Prof. Dr. José Rubens Macedo Junior
# 28/03/2024

import pandas as pd
import time

# Capturando o tempo de início
start_time = time.time()

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

# Definição dos limites para P99% e P1%
limp99_220 = 233
limp1_220 = 202
limp99_127 = 133
limp1_127 = 117

# Calcular os valores médio, máximo, mínimo, percentil 1% e percentil 99% para cada coluna
medias = df.mean()
maximos = df.max()
minimos = df.min()
percentil_1 = df.quantile(0.01)
percentil_99 = df.quantile(0.99)

# Total de registros por coluna
total_registros = len(df)

# Resultados por coluna
resultados = {}

# Inicialização de contadores para o resumo final
qtdemedicoes = 0
contagem_drp = 0
contagem_drc = 0
contagem_drp_drc = 0
contagem_drp_zero = 0
contagem_drc_zero = 0
contagem_drpdrc_zero = 0
contagem_p99 = 0
contagem_p1 = 0
contagem_p1_p99 = 0

# Analisar e contar valores para cada coluna
for coluna in df.columns:
    media_coluna = medias[coluna]
    valores_coluna = df[coluna]
    maximo_coluna = maximos[coluna]
    minimo_coluna = minimos[coluna]
    perc_1_coluna = percentil_1[coluna]
    perc_99_coluna = percentil_99[coluna]
    contagem = {'n_criticos': 0, 'n_precarios': 0}

# Contagem de violações de percentil
    if media_coluna >= 150:
        if perc_99_coluna > limp99_220:
            contagem_p99 += 1
        if perc_1_coluna < limp1_220:
            contagem_p1 += 1
        if perc_99_coluna > limp99_220 and perc_1_coluna < limp1_220:
            contagem_p1_p99 += 1
    else:
        if perc_99_coluna > limp99_127:
            contagem_p99 += 1
        if perc_1_coluna < limp1_127:
            contagem_p1 += 1
        if perc_99_coluna > limp99_127 and perc_1_coluna < limp1_127:
            contagem_p1_p99 += 1
# Fim da contagem de violações percentil

    for valor in df[coluna]:
        if media_coluna >= 150:
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
    resultados[coluna] = {'DRP': drp, 'DRC': drc, 'Max': maximo_coluna, 'Min': minimo_coluna, 'Media': media_coluna,
                          'Percentil1': perc_1_coluna, 'Percentil99': perc_99_coluna}

    # Contador da quantidade total de medições
    qtdemedicoes += 1

    # Contagem de violações DRP e DRC
    if drp > 100*limDRP and drc > 100*limDRC:
        contagem_drp_drc += 1
    if drp > 100*limDRP:
        contagem_drp += 1
    if drc > 100*limDRC:
        contagem_drc += 1
    if drp == 0:
        contagem_drp_zero += 1
    if drc == 0:
        contagem_drc_zero += 1
    if drp == 0 and drc == 0:
        contagem_drpdrc_zero += 1

# Exibir os valores de DRP, DRC, percentil 1% e percentil 99% para cada coluna
for coluna, indicadores in resultados.items():
    print(f"{coluna}: DRP = {indicadores['DRP']:.2f}%, DRC = {indicadores['DRC']:.2f}%, Max = {indicadores['Max']:.2f}V, "
          f"Min = {indicadores['Min']:.2f}V, Media = {indicadores['Media']:.2f}V, Percentil 1% = {indicadores['Percentil1']:.2f}V, "
          f"Percentil 99% = {indicadores['Percentil99']:.2f}V")

percdrp = (contagem_drp / qtdemedicoes) * 100
percdrc = (contagem_drc / qtdemedicoes) * 100
percdrpdrc = (contagem_drp_drc / qtdemedicoes) * 100

percp99 = (contagem_p99 / qtdemedicoes) * 100
percp1 = (contagem_p1 / qtdemedicoes) * 100
percp1p99 = (contagem_p1_p99 / qtdemedicoes) * 100

totaldrpdrc = (contagem_drp + contagem_drc - contagem_drp_drc)
totalp1p99 = (contagem_p1 + contagem_p99 - contagem_p1_p99)
perctotaldrpdrc = (totaldrpdrc / qtdemedicoes) * 100
perctotalp1p99 = (totalp1p99 / qtdemedicoes) * 100
percdrpzero = (contagem_drp_zero / qtdemedicoes) * 100
percdrczero = (contagem_drc_zero / qtdemedicoes) * 100
percdrpdrczero = (contagem_drpdrc_zero / qtdemedicoes) * 100

# Exibir o resumo final
print('--------------------------------------------------------------------------------')
print(f'\033[1m   R E S U M O    F I N A L\033[0m')
print('--------------------------------------------------------------------------------')
print(f"   Total de medições analisadas ............................: \033[95m{qtdemedicoes}\033[0m")
print('--------------------------------------------------------------------------------')
print(f"   Quantidade total de violações de DRP e/ou DRC ...........: \033[93m{totaldrpdrc}\033[0m ( \033[94m{perctotaldrpdrc:.2f}%\033[0m )")
print(f"   Quantidade de medições que violaram DRP .................: \033[93m{contagem_drp}\033[0m ( \033[94m{percdrp:.2f}%\033[0m )")
print(f"   Quantidade de medições que violaram DRC .................: \033[93m{contagem_drc}\033[0m ( \033[94m{percdrc:.2f}%\033[0m )")
print(f"   Quantidade de medições que violaram DRP e DRC ...........: \033[93m{contagem_drp_drc}\033[0m ( \033[94m{percdrpdrc:.2f}%\033[0m )")
print(f"   Quantidade de medições com DRP = 0 ......................: \033[93m{contagem_drp_zero}\033[0m ( \033[94m{percdrpzero:.2f}%\033[0m )")
print(f"   Quantidade de medições com DRC = 0 ......................: \033[93m{contagem_drc_zero}\033[0m ( \033[94m{percdrczero:.2f}%\033[0m )")
print(f"   Quantidade de medições com DRP e DRC = 0 ................: \033[93m{contagem_drpdrc_zero}\033[0m ( \033[94m{percdrpdrczero:.2f}%\033[0m )")
print('--------------------------------------------------------------------------------')
print(f"   Quantidade total de violações de P1% e/ou P99% ..........: \033[93m{totalp1p99}\033[0m ( \033[94m{perctotalp1p99:.2f}%\033[0m )")
print(f"   Quantidade de medições que violaram P99% ................: \033[93m{contagem_p99}\033[0m ( \033[94m{percp99:.2f}%\033[0m )")
print(f"   Quantidade de medições que violaram P1% .................: \033[93m{contagem_p1}\033[0m ( \033[94m{percp1:.2f}%\033[0m )")
print(f"   Quantidade de medições que violaram P99% e P1% ..........: \033[93m{contagem_p1_p99}\033[0m ( \033[94m{percp1p99:.2f}%\033[0m )")
print('--------------------------------------------------------------------------------')

# Converter os resultados em um DataFrame para exportação
df_resultados = pd.DataFrame.from_dict(resultados, orient='index')
df_resultados.reset_index(inplace=True)
df_resultados.rename(columns={'index': 'Coluna'}, inplace=True)

# Salvando o DataFrame em arquivo Excel
caminho_arquivo_saida = r'C:\pythonjr\drpdrc\resultados.xlsx'
df_resultados.to_excel(caminho_arquivo_saida, index=False)

print('Os resultados foram salvos com sucesso em arquivo Excel!')
print('Localização:', caminho_arquivo_saida)

# Capturando o tempo de término
end_time = time.time()
# Calculando o tempo de execução do código
execution_time = end_time - start_time
print(f"Tempo de execução do código: {execution_time:.2f} segundos")