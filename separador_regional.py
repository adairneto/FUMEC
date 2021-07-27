'''
Still being implemented. The idea is simply to take a list and generate n new sheets based on that list.
'''

import pandas as pd

filename = input("File address (.csv): ")
df = pd.read_csv('filename', encoding='utf-8', sep=";")

# Listar regionais
regional = []
for i in df.index:
    if df['Search'][i] not in regional:
        regional.append(df['Search'][i])

# Cria nova planilha por regional
norte = pd.DataFrame(columns = df.columns)
j = 0 
for i in df.index:
    if df['Search'][i] == "Regional Norte":
        norte.append(df.loc(i), ignoreIndex=False)
        j += 1

sul = pd.DataFrame(columns = df.columns)
leste = pd.DataFrame(columns = df.columns)
noroeste = pd.DataFrame(columns = df.columns)
sudoeste = pd.DataFrame(columns = df.columns)