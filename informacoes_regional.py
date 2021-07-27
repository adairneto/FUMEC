import pandas as pd
   
filename = input("Insira o nome do arquivo: ")
df = pd.read_excel(filename)

# Parte 1: Regional

# Situações em que consta como aluno ativo
situacoes = ['SEMESTRE ESPECIAL', 'TÃRMINO - C.E.', 'TÃRMINO - PAA', 'MATRICULADO','PROGRESSÃO NO CICLO - FASE 2', 'PROMOVIDO P/ CICLO 2', 'PROGRESSÃO NO CICLO - FASE 4', 'PROGRESSÃO NO CICLO - FASE 3', 'RETIDO NO CICLO 1', 'RECLASSIFICADO CICLO 2 FASE 2', 'CONCLUINTE', 'RETIDO NO CICLO 2', 'PROGRESSÃO - C.E.', 'TÉRMINO - C.E.', 'CONTINUIDADE - PEALV', 'CONCLUINTE', 'CONTINUIDADE - PAA', 'TÉRMINO - PAA', 'TÉRMINO - PEALV']

for i in ['REGIONAL NORTE', 'REGIONAL SUL', 'REGIONAL LESTE', 'REGIONAL NOROESTE', 'REGIONAL SUDOESTE', 'PAA', 'GPEJA']:
    
    # Dados gerais
    df1 = df[df['REGIONAL'] == i]
    df2 = df1.set_index(["NOME PROFESSOR", "CENTRO DE CUSTO"]).count(level="NOME PROFESSOR")["ID"]
    professores = df2.count()
    df3 = df1.set_index(["CENTRO DE CUSTO", "PROGRAMA"]).count(level="CENTRO DE CUSTO")["ID"]
    classes = df3.count()
    alunos = len(df1.count(axis = 'columns'))
    df_ativos_reg = df1[df1['SITUAÇÃO'].isin(situacoes)]
    mat_ativas = len(df_ativos_reg.count(axis = 'columns'))
    
    # Resumo da regional
    print("Resumo total regional {}:\n".format(i))
    print("Alunos: {}; Matrículas Ativas: {}; Professores: {}; CDs: {}\n\n".format(alunos, mat_ativas, professores, classes))
  
    # Recorte por programa
    
    # programas = df_uef['PROGRAMA'].unique()
    programas = ['EJA', 'PCE', 'PEALV', 'PAA']
    
    for j in programas:
        print("Resumo de regional {} - Programa {} \n".format(i, j))
        df_temp = df1[df1['PROGRAMA'] == j] # DF com programa atual
        df_temp2 = df_temp.set_index(["NOME PROFESSOR", "CENTRO DE CUSTO"]).count(level="NOME PROFESSOR")["ID"]
        professores_programa = df_temp2.count()
        df_temp3 = df_temp.set_index(["CENTRO DE CUSTO", "PROGRAMA"]).count(level="CENTRO DE CUSTO")["ID"]
        classes_programa = df_temp3.count()
        alunos_programa = len(df_temp.count(axis = 'columns'))
        df_ativos_programa = df_temp[df_temp['SITUAÇÃO'].isin(situacoes)]
        mat_ativas_programa = len(df_ativos_programa.count(axis = 'columns'))
        print("Alunos: {}; Matrículas Ativas: {}; Professores: {}; CDs: {}\n\n".format(alunos_programa, mat_ativas_programa, professores_programa, classes_programa))

        # Gera planilha com informações
        # titulo = "{} - {}.xlsx".format(i, j)
        # planilha_programa.append([alunos_programa, mat_ativas_programa, professores_programa, classes_programa])
        # df_programa = pd.DataFrame(planilha_programa, columns = ['Alunos', 'Matrículas Ativas', 'Professores', 'CDs'])
        # df.to_excel(titulo, index = True)


# Tabela dinâmica com alunos por programa e regional
# print("TABELA DINÂMICA")
# table_reg = pd.pivot_table(df, values='ID', index=['PROGRAMA'], columns='REGIONAL', aggfunc='count', fill_value=0, margins=False, dropna=True, margins_name='All', observed=False)
# print(table_reg)

# Parte 2: UEF

final_uef = []

# uefs = df['UEF'].unique()
uefs = ['5045 - UEF PE EMEF PE. JOSE NARCISO V. EHRENBERG',
       '5047 - UEF EMEF JOAO ALVES DOS SANTOS',
       '5055 - UEF PROFA ANALIA FERRAZ DA COSTA COUTO - EMEF',
       '5194 - UEF FUMEC DESCENTRALIZADA CASI',
       '5298 - UEF PFTO ANTONIO DA COSTA SANTOS',   
       '5420 - UEF CPAT - CENTRO PÚBLICO DE APOIO AO TRABALHADOR',
       '5070 - UEF PE FRANCISCO SILVA EMEF',
       '5394 - UEF CEPROCAMP JOSÉ ALVES',    
       '5080 - UEF MARIA PAVANATTI FÁVARO - EMEF',
       '5218 - UEF FUMEC DESCENTRALIZADA CAMBARÁ',
       'PAA', 'CPEJA']

for i in uefs:
    df_uef = df[df['UEF'] == i]
    
    final_uef.append([i])
    
    # Dados gerais
    df2 = df_uef.set_index(["NOME PROFESSOR", "CENTRO DE CUSTO"]).count(level="NOME PROFESSOR")["ID"]
    professores = df2.count()
    df3 = df_uef.set_index(["CENTRO DE CUSTO", "PROGRAMA"]).count(level="CENTRO DE CUSTO")["ID"]
    classes = df3.count()
    alunos = len(df_uef.count(axis = 'columns'))
    df_ativos_reg = df_uef[df_uef['SITUAÇÃO'].isin(situacoes)]
    mat_ativas = len(df_ativos_reg.count(axis = 'columns'))
    
    # Resumo da UEF
    print("Resumo total regional {}:\n".format(i))
    print("Alunos: {}; Matrículas Ativas: {}; Professores: {}; CDs: {}\n\n".format(alunos, mat_ativas, professores, classes))
    
    # Recorte por programa
       
    # programas = df_uef['PROGRAMA'].unique()
    programas = ['EJA', 'PCE', 'PEALV', 'PAA']
    
    for j in programas:
        print("Resumo de UEF {} = Programa {} \n".format(i, j))
        df_temp = df_uef[df_uef['PROGRAMA'] == j]
        df_temp2 = df_temp.set_index(["NOME PROFESSOR", "CENTRO DE CUSTO"]).count(level="NOME PROFESSOR")["ID"]
        professores_uef = df_temp2.count()
        df_temp3 = df_temp.set_index(["CENTRO DE CUSTO", "PROGRAMA"]).count(level="CENTRO DE CUSTO")["ID"]
        classes_uef = df_temp3.count()
        alunos_uef = len(df_temp.count(axis = 'columns'))
        df_ativos_uef = df_temp[df_temp['SITUAÇÃO'].isin(situacoes)]
        mat_ativas_uef = len(df_ativos_uef.count(axis = 'columns'))
        print("Alunos: {}; Matrículas Ativas: {};Professores: {}; CDs: {}\n\n".format(alunos_uef, mat_ativas_uef, professores_uef, classes_uef))
        
        final_uef.append(["Resumo do programa {}".format(j)])
        final_uef.append(['Alunos', 'Matrículas Ativas', 'Professores', 'CD'])
        final_uef.append([alunos_uef, mat_ativas_uef, professores_uef, classes_uef])

    final_uef.append(["Resumo da {}".format(i)])
    final_uef.append([alunos, mat_ativas, professores, classes])

# Tabela dinâmica com alunos por programa e uef
# print("TABELA DINÂMICA")
# table_uef = pd.pivot_table(df, values='ID', index=['PROGRAMA'], columns='UEF', aggfunc='count', fill_value=0, margins=False, dropna=True, margins_name='All', observed=False)
# print(table_uef)

# Impressão
# table_reg.to_excel('2015-REG.xlsx', index = True)
# table_uef.to_excel('2015-UEF.xlsx', index = True)