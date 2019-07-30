# -*- coding: utf-8 -*-
"""
Created on Wed Jul 31 00:01:04 2019

@author: eduardo

1 - Colete a informação desejada, nesse caso o número de cada item que está dentro da ficha.
2 - A váriavel nom_files é proveniente do script de extração dos museus, com ela é possível armazenar 
a sequência que o script utilizou para processar os arquivos, lembrando que essa informação contém o caminho da pasta, que deve
ser retirado.
3 - A váriavel flattened diz respeito a informação retirada no item 1, lembre-se que essa informação é retirada como uma list of lists,
por isso temos que tornar ela em uma lista única ou flat, feito isso adicione ao fim dos dados coletados nessa lista o ".docx".
4 - Feito isso o script faz o resto do tralho, comparando a ordem entre as duas listas e renomeando de acordo com essa ordem.
"""

folder_name = 'caminho'
files_list = os.listdir(folder_name)
for f in files_list: # iterate over all files in files_list
    if f.endswith(".docx"): # check if DOCX file (optional)
        if f in nom_files:
            print (f + ' se torna ' + flattened[nom_files.index(f)]) # DEBUG
            os.rename(folder_name+f, folder_name+flattened[nom_files.index(f)])
