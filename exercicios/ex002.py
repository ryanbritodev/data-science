import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

# Gerar 100 números aleatórios (1 a 99)
np.random.seed(42)
dados = np.random.randint(1, 100, 100)

# Dataframe
df = pd.DataFrame(dados, columns=['Valores'])

# Criar uma tabela 10x10
df_tabela = df.values.reshape(10, 10)
df_tabela = pd.DataFrame(df_tabela)

# Exibindo o dataset
print("\nDataset - 100 Números Aleatórios (Tabela 10x10):")
print(df_tabela)

# Criar m histograma
plt.figure(figsize=(8, 5))
plt.hist(df['Valores'], bins=10, edgecolor='black', alpha=0.7)
plt.xlabel('Intervalo de Valores')
plt.ylabel('Frequência')
plt.title("Histograma do Dataset - 100 Números Aleatórios")
plt.grid(axis="y", linestyle='--', alpha=0.7)
plt.show()
