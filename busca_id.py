# Code written to extract ID information from the user e-mail. Some optional code is commented.

import pandas as pd

filename = input("Insira o nome do arquivo: ")
df = pd.read_csv(filename, encoding='utf-8', sep=";")

# There's no need for splitting the data, all that can be done in the for loop. Just here for reference material.
# df = df['First Name [Required],Last Name [Required],Email Address [Required],Status [READ ONLY],Last Sign In [READ ONLY],Email Usage [READ ONLY]'].str.split(',', expand=True)
# Splits the e-mail
# df["A"] = ""
# df["B"] = ""
# df[['A', 'B']] = df['Email Address [Required]'].str.split('@', 1, expand=True)

# Splits the letters and digits from the column A and then creates a new column with the ID
for i in df.index:
    digits = []
    name = []
    # If the csv is already splitted, use 'Email Address [Required]' instead of 2
    for j in df['Email Address [Required]'][i]: # If using the split method above, change 'Email Address [Required' to A
        if j.isdigit():
            digits.append(j)
        else:
            name.append(j)
    # df.at[i,'NOME'] = ''.join(name)
    
    # Not always necessary
    if len(digits) > 9:
        digits = digits[:-1]
        
    df.at[i,'ID'] = ''.join(digits)

# Deletes unnecessary data
# del df["A"], df["B"]

output = input("Insira o nome desejado:")+".xlsx"
df.to_excel(output, index=False)