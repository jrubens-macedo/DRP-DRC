# Gerador de gráficos dos indicadores de qualidade de tensão em regime permanente
# DRP, DRC, P99% e P1%
# By Prof. Dr. José Rubens Macedo Junior
# 28/03/2024

import pandas as pd
import matplotlib.pyplot as plt

# Carregar os arquivos Excel
arquivo_excel_1 = r'C:\pythonjr\drpdrc\dados_volts_127_220.xlsx'
arquivo_excel_2 = r'C:\pythonjr\drpdrc\resultados.xlsx'
df1 = pd.read_excel(arquivo_excel_1)
df2 = pd.read_excel(arquivo_excel_2)

# Plotagem das tensões
plt.figure(figsize=(10, 6))
for coluna in df1.columns:
    plt.plot(df1[coluna], label=coluna, linewidth=0.6)
plt.xlabel('Amostras de tensão de 10 minutos', fontsize=16)
plt.ylabel('Tensão (V)', fontsize=16)
plt.xticks(fontsize=14)
plt.yticks(fontsize=14)
plt.grid(True, linestyle='--')
plt.show()

# Plotar os valores máximos e mínimos de tensão
maximos = df2.iloc[:, 4].values
minimos = df2.iloc[:, 5].values
plt.figure(figsize=(10, 6))
plt.scatter(range(len(maximos)), maximos, label='Máximos', color='red', s=10)
plt.scatter(range(len(minimos)), minimos, label='Mínimos', color='blue', s=10)
plt.xlabel('Medição de tensão', fontsize=16)
plt.ylabel('Tensão (V)', fontsize=16)
plt.legend()
plt.grid(True, linestyle='--')
plt.xticks(fontsize=14)
plt.yticks(fontsize=14)
plt.show()

# Plotagem do histograma de tensão
histv = df1.values.flatten()
plt.figure(figsize=(10, 6))
plt.hist(histv, bins=100, color='green', edgecolor='black', alpha=0.7)
plt.xlabel('Tensão eficaz de 10 minutos', fontsize=16)
plt.ylabel('Qtde. de ocorrências', fontsize=16)
plt.xticks(fontsize=14)
plt.yticks(fontsize=14)
plt.grid(True, linestyle='--')
plt.show()

# Plotagem dos histogramas de DRP e DRC
fig, axs = plt.subplots(1, 2, figsize=(12, 5))
# Gerar histograma dos DRPs
axs[0].hist(df2.iloc[:, 1], bins=100, color='red', edgecolor='black', alpha=0.7, linewidth=0)
axs[0].set_xlabel('DRP (%)', fontsize=16)
axs[0].set_ylabel('Qtde. de ocorrências', fontsize=16)
axs[0].set_title('(a)', fontsize=18)
axs[0].grid(True, linestyle='--')
axs[0].set_ylim(0, 150)
axs[0].set_xlim(0, 100)
axs[0].tick_params(axis='x', labelsize=14)
axs[0].tick_params(axis='y', labelsize=14)
# Gerar histograma dos DRCs
axs[1].hist(df2.iloc[:, 2], bins=100, color='blue', edgecolor='black', alpha=0.7, linewidth=0)
axs[1].set_xlabel('DRC (%)', fontsize=16)
axs[1].set_ylabel('Qtde. de ocorrências', fontsize=16)
axs[1].set_title('(b)', fontsize=18)
axs[1].grid(True, linestyle='--')
axs[1].set_ylim(0, 150)
axs[1].set_xlim(0, 100)
axs[1].tick_params(axis='x', labelsize=14)
axs[1].tick_params(axis='y', labelsize=14)
plt.tight_layout()
plt.show()

# Plotagem dos histogramas de P1% e P99%
fig, axs = plt.subplots(1, 2, figsize=(12, 5))
# Gerar histograma dos P99%
axs[0].hist(df2.iloc[:, 6], bins=100, color='red', edgecolor='black', alpha=0.7, linewidth=0)
axs[0].set_xlabel('Percentil 1% (V)', fontsize=16)
axs[0].set_ylabel('Qtde. de ocorrências', fontsize=16)
axs[0].set_title('(a)', fontsize=18)
axs[0].grid(True, linestyle='--')
axs[0].set_ylim(0, 250)
axs[0].set_xlim(50, 300)
axs[0].tick_params(axis='x', labelsize=14)
axs[0].tick_params(axis='y', labelsize=14)
# Gerar histograma dos P1%
axs[1].hist(df2.iloc[:, 7], bins=100, color='blue', edgecolor='black', alpha=0.7, linewidth=0)
axs[1].set_xlabel('Percentil 99% (V)', fontsize=16)
axs[1].set_ylabel('Qtde. de ocorrências', fontsize=16)
axs[1].set_title('(b)', fontsize=18)
axs[1].grid(True, linestyle='--')
axs[1].set_ylim(0, 250)
axs[1].set_xlim(50, 300)
axs[1].tick_params(axis='x', labelsize=14)
axs[1].tick_params(axis='y', labelsize=14)
plt.tight_layout()
plt.show()






