# Objetivo:
# Caso 1: juntar várias planilhas num único arquivo
# Caso 2: juntar abas de um mesmo arquivo
# Caso 3: juntar vários dataframes num mesmo .xlsx

import pandas as pd

def multiple_files(df_list): # Case 1: Multiple .xlsx files
    from os import listdir
    from os.path import isfile, join
    mypath = input("Escreva o endereço no qual se encontram os arquivos: ")
    planilhas = [f for f in listdir(mypath) if isfile(join(mypath, f))]
    
    for i in planilhas:
        if i[-4:] != 'xlsx':
            planilhas.remove(i)
    
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
    dfl.to_excel(output, index=False, header=False)


def multiple_tabs(df_list): # Case 2: Multiple tabs at the same .xlsx file
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
    dfl.to_excel(output, index=False, header=False)

def same_tab(): # Case 3: Join multiple dataframes at the same .xlsx tab

    filename = input("File address: ")
    df = pd.read_excel(filename, header = None)
    
    n = df.shape[0]
    
    i = 0
    while i < n:
        if type(df[0][i]) == str and 'PLANILHA  INTERAÇÃO ALUNO X PROFESSOR' in df[0][i]:
            to_drop = [j for j in range(i, i+7)]
            df.drop(index=to_drop, inplace = True)
            i += 7
        else:
            i += 1
            
    df = df[df[1].notna()]
    
    output = input("Insira o nome desejado:")+".xlsx"
    df.to_excel(output, index=False, header=False)
            
def main():
    df_list = []
    choice = int(input("'1' for multiple .xlsx files\n'2' for multiple tabs at the same .xlsx file\n'3' for multiple sheets at the same tab\nChoice: "))
    if choice == 1:
        multiple_files(df_list)
    elif choice == 2:
        multiple_tabs(df_list)
    elif choice == 3:
        same_tab()
        
if __name__ == "__main__":
    main()