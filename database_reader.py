'''
Goal: this code was written for a specific application at FUMEC, but can be generalized.
It takes a database (.csv), treats the dates and drop ID duplicates keeping the last entries.
Then it adds the age and some other useful info, simplifies and reorganizes the columns generating an .xlsx file.
'''

import pandas as pd
import datetime

filename = input("Insira o nome do arquivo: ")
df = pd.read_csv(filename, encoding='utf-8', sep=";")

# Removes the data from the actual month (not implemented)

# Adds the date 31/12/2021 to the empty cells in DATA SAÍDA 
for i in df.index:
    if df['DATA SAÍDA'][i] != df['DATA SAÍDA'][i]:
        df.at[i,'DATA SAÍDA'] = datetime.date(2021, 12, 31).strftime("%d/%m/%Y")

df['DATA ENTRADA'] = pd.to_datetime(df['DATA ENTRADA'], format="%d/%m/%Y")
df['DATA SAÍDA'] = pd.to_datetime(df['DATA SAÍDA'], format="%d/%m/%Y")
df.sort_values(['ID', 'DATA SAÍDA'], ascending=['True','False'], inplace=True)

# Delete duplicates
df = df.drop_duplicates(subset = "ID", keep="last")

# Remove the added date
for i in df.index:
    if df['DATA SAÍDA'][i] == pd.to_datetime('31/12/2021', format="%d/%m/%Y"):
        df.at[i,'DATA SAÍDA'] = float('nan')
        
# Adds age
def calculate_age(born):
    today = datetime.date.today()
    return today.year - born.year - ((today.month, today.day) < (born.month, born.day))

df['DATA DE NASCIMENTO.'] = pd.to_datetime(df['DATA DE NASCIMENTO.'], format="%d/%m/%Y")
for i in df.index:
    df.at[i,'IDADE'] = calculate_age(df['DATA DE NASCIMENTO.'][i])
    if df['IDADE'][i] < 10:
        df.at[i,'FAIXA ETÁRIA'] = "0-9"
    elif df['IDADE'][i] < 20:
        df.at[i,'FAIXA ETÁRIA'] = "10-19"
    elif df['IDADE'][i] < 30:
        df.at[i,'FAIXA ETÁRIA'] = "20-29"
    elif df['IDADE'][i] < 40:
        df.at[i,'FAIXA ETÁRIA'] = "30-39"        
    elif df['IDADE'][i] < 50:
        df.at[i,'FAIXA ETÁRIA'] = "40-49"        
    elif df['IDADE'][i] < 60:
        df.at[i,'FAIXA ETÁRIA'] = "50-59"        
    elif df['IDADE'][i] < 70:
        df.at[i,'FAIXA ETÁRIA'] = "60-69"
    elif df['IDADE'][i] < 80:
        df.at[i,'FAIXA ETÁRIA'] = "70-79"
    elif df['IDADE'][i] < 90:
        df.at[i,'FAIXA ETÁRIA'] = "80-89"
    elif df['IDADE'][i] < 100:
        df.at[i,'FAIXA ETÁRIA'] = "90-99"

# Adds the program
for i in df.index:
    if df["SÉRIE"][i] == "CE1 - CONSOL. ESCOLARIDADE - SEM 1":
        df.at[i,"PROGRAMA"] = "PCE"
    elif df["SÉRIE"][i] == "CE2 - CONSOL. ESCOLARIDADE - SEM 2":
        df.at[i,"PROGRAMA"] = "PCE"
    elif df["SÉRIE"][i] == "CCA - PROGRAMA EDUCAÇÃO AMPLIADA AO":
        df.at[i,"PROGRAMA"] = "PEALV"
    elif df["SÉRIE"][i] == "PAA - PROGRAMA APOIO A ALFABETIZAÇÃO":
        df.at[i,"PROGRAMA"] = "PAA"
    else:
        df.at[i,"PROGRAMA"] = "EJA"
        
# Simplifies column Etnia
for i in df.index:
    if df["ETNIA"][i] == "-":
        df.at[i,"ETNIA"] = "NAO INFORMADA"
    elif df["ETNIA"][i] == "NAO DECLARADA":
        df.at[i,"ETNIA"] = "NAO INFORMADA"
        
# Reorder columns
cols = ['ID',
 'RA',
 'RG',
 'CPF',
 'NOME',
 'SEXO',
 'ETNIA',
 'DATA DE NASCIMENTO.',
 'IDADE',
 'FAIXA ETÁRIA',
 'DDD',
 'RESIDENCIAL',
 'RECADO',
 'DDD CELULAR',
 'CELULAR',
 'NACIONALIDADE',
 'DATA ENTRADA NO BRASIL',
 'RNE',
 'DATA EMISSÃO RNE',
 'MUNICÍPIO',
 'UF',
 'NECESSIDADES ESPECIAIS',
 'ENDEREÇO',
 'Nº',
 'COMPLEMENTO',
 'BAIRRO',
 'CIDADE DO ALUNO',
 'UF DO ALUNO',
 'CEP',
 'SITUAÇÃO ANTERIOR',
 'SITUAÇÃO',
 'SÉRIE',
 'PROGRAMA',
 'DATA ENTRADA',
 'DATA SAÍDA',
 'CENTRO DE CUSTO',
 'REGIONAL',
 'ENDEREÇO DA ESCOLA',
 'Nº DA ESCOLA',
 'BAIRRO.1',
 'CIDADE',
 'UF.1',
 'CEP.1',
 'TURNO',
 'TURMA',
 'NOME PROFESSOR',
 'TEMPO DE PERMANÊNCIA NA FUMEC',
 'QUANT. DE DESISTÊNCIAS DO ALUNO',
 'LAUDADO']
df = df[cols]
        
# Generate Pivot Table (not implemented)
conta_programa = pd.pivot_table(df, values = 'ID', index = 'PROGRAMA', columns='SITUAÇÃO', aggfunc='count')
conta_regional = pd.pivot_table(df, values = 'ID', index = 'REGIONAL', columns='SITUAÇÃO', aggfunc='count')
conta_idade = pd.pivot_table(df, values = 'ID', index = 'FAIXA ETÁRIA', columns='SITUAÇÃO', aggfunc='count')
conta_sexo = pd.pivot_table(df, values = 'ID', index = 'SEXO', columns='SITUAÇÃO', aggfunc='count')
conta_etnia = pd.pivot_table(df, values = 'ID', index = 'ETNIA', columns='SITUAÇÃO', aggfunc='count')
conta_especiais = pd.pivot_table(df, values = 'ID', index = 'NECESSIDADES ESPECIAIS', columns='SITUAÇÃO', aggfunc='count')

# Write final CSV file
output = input("Insira o nome desejado:")+".xlsx"
# df.to_excel(output, index=False)

writer = pd.ExcelWriter(output,engine='xlsxwriter')   
df.to_excel(writer,sheet_name='Base',startrow=0 , startcol=0, index=False)
dim_1 = len(conta_programa)
dim_2 = dim_1 + len(conta_regional)
dim_3 = dim_2 + len(conta_idade)
dim_4 = dim_3 + len(conta_sexo)
dim_5 = dim_4 + len(conta_etnia)
conta_programa.to_excel(writer,sheet_name='TDs',startrow=0, startcol=0)
conta_regional.to_excel(writer,sheet_name='TDs',startrow=dim_1, startcol=0) 
conta_idade.to_excel(writer,sheet_name='TDs',startrow=dim_2, startcol=0) 
conta_sexo.to_excel(writer,sheet_name='TDs',startrow=dim_3, startcol=0) 
conta_etnia.to_excel(writer,sheet_name='TDs',startrow=dim_4, startcol=0) 
conta_especiais.to_excel(writer,sheet_name='TDs',startrow=dim_5, startcol=0) 
writer.save()