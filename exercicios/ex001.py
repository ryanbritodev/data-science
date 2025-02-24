import pandas as pd
import numpy as np

dataset = pd.read_csv('/content/dataset_estudo.csv')
datasetOrdenado = dataset.transform(np.sort)

# Selecionando todos os valores númericos das colunas ordenadas
moda = datasetOrdenado.select_dtypes(include=["number"]).mode()
media = datasetOrdenado.select_dtypes(include=["number"]).mean()
mediana = datasetOrdenado.select_dtypes(include=["number"]).median()

print("--- Moda, Média e Mediana ---")
print(datasetOrdenado)

print() # Quebra de linha
print(f"Moda das Idades: {moda.iloc[0, 0]}")
print(f"Moda das Horas de Estudo: {moda.iloc[0, 1]}")
print(f"Moda das Notas: {moda.iloc[0, 2]}")

print() # Quebra de linha
print(f"Média das Idades: {media.iloc[0]}")
print(f"Média das Horas de Estudo: {media.iloc[1]}")
print(f"Média das Notas: {media.iloc[2]}")

print() # Quebra de linha
print(f"Mediana das Idades: {mediana.iloc[0]}")
print(f"Mediana das Horas de Estudo: {mediana.iloc[1]}")
print(f"Mediana das Notas: {mediana.iloc[2]}")
