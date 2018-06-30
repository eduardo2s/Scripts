# -*- coding: utf-8 -*-
"""
Created on Fri Jun 29 22:04:49 2018

@author: eduar
"""

import csv 
from tqdm import tqdm
file_name = "C:/Users/eduar/Desktop/autores_tratados.csv" #filename is argument 1
            
with open(file_name, encoding='utf-8') as f:
    reader = csv.reader(f, delimiter =";")
    next(reader) # skip header
    data = [r for r in reader]
        
source = []
target = []
for item in tqdm(data, desc="Criando Relações"):
    for i in range(len(item)-1):
        for j in range(1,len(item)):
            if item[i] != item[j]:
                source.append(item[i])
                target.append(item[j])
                #print(' -- '.join([item[i], item[j]]))

