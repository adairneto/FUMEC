# Objetivo: juntar várias planilhas num único arquivo (caso 1) ou juntar abas de um mesmo arquivo (caso 2)

import pandas as pd

def multiple_files(): # Case 1: Multiple .xlsx files
    from os import listdir
    from os.path import isfile, join
    mypath = input("Escreva o endereço no qual se encontram os arquivos: ")
    planilhas = [f for f in listdir(mypath) if isfile(join(mypath, f))]
    print(planilhas)
    rem = int(input("Até qual linha quer remover? Digite -1 para nenhuma: "))
    if rem != -1:
        to_drop = [i for i in range(rem)]
    
    for i in planilhas:
        filename = mypath+i
        df = pd.read_excel(filename)
        if rem != -1:
            df = df.drop(index=to_drop)
        # df = df[:-1] # Comment if you want the last line
        df_list.append(df)
    
    dfl = pd.concat(df_list)
    dfl = dfl[dfl[1].notna()] # If you need to remove empty rows, uncomment this line
    # dfl.columns = ['ID', 'NOME', 'PLATAFORMA', 'WHATSAPP', 'APOSTILA', 'E-MAIL', 'TELEFONE', 'OUTROS']
    
    output = input("Insira o nome desejado:")+".xlsx"
    dfl.to_excel(output, index=False)


def multiple_tabs(): # Case 2: Multiple tabs at the same .xlsx file
    # Alternative: can use df = pd.read_excel(filename, sheet_name = None) # read all sheets
    
    filename = input("File address: ")
    rem = int(input("Até qual linha quer remover? Digite -1 para nenhuma: "))
    if rem != -1:
        to_drop = [i for i in range(rem)]
    
    xlsx = pd.ExcelFile(filename)
    for sheet_name in xlsx.sheet_names:
        df = pd.read_excel(xlsx, sheet_name, header=None)
        if rem != -1:
            df = df.drop(index=to_drop)
        #df = df[:-1]
        df_list.append(df)
    
    dfl = pd.concat(df_list)
    dfl = dfl[dfl[1].notna()] # If you need to remove empty rows, uncomment this line
    # dfl.columns = ['ID', 'NOME', 'PLATAFORMA', 'WHATSAPP', 'APOSTILA', 'E-MAIL', 'TELEFONE', 'OUTROS']
    
    output = input("Insira o nome desejado:")+".xlsx"
    dfl.to_excel(output, index=False)
    
df_list = []
choice = int(input("Type\n'1' for multiple .xlsx files\n'2' for multiple tabs at the same .xlsx file\n"))
if choice == 1:
    multiple_files()
elif choice == 2:
    multiple_tabs()