# -*- coding: utf-8 -*-
"""
Created on Fri Mar 16 21:41:54 2018

@author: eduar
"""

#from time import sleep
#import pythoncom
from time import sleep
import win32com.client as win32
import glob
import os
from tqdm import tqdm

lines_desired = [1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27]
lines_desired_2 = [1,2,3,4,5,6,7,8,9,10,11]
word = win32.gencache.EnsureDispatch('Word.Application')
#word = win32.Dispatch('Word.Application')
word.Visible = False
#sleep(1)

final_list=[]
n = 0
for infile in tqdm(glob.glob( os.path.join('', '*.doc') ), desc="processando dados", unit="files"):
#My checks here
    doc = word.Documents.Open(os.getcwd()+"\\"+infile)

    sleep(1)
    if doc.Tables.Count == 2:
        table = doc.Tables(1)
        table_2 = doc.Tables(2)
        linhas = table.Rows.Count + table_2.Rows.Count
        #print("{} Processado... {}".format(infile, n))
        n += 1
        for i in range(1 , linhas):
            listateste = []
            listateste.append(table.Cell(Row = lines_desired[1], Column = 1).Range.Text.rstrip('\r\x07'))
            listateste.append(table.Cell(Row = lines_desired[1], Column = 2).Range.Text.rstrip('\r\x07'))
            listateste.append(table.Cell(Row = lines_desired[1], Column = 3).Range.Text.rstrip('\r\x07'))
            listateste.append(table.Cell(Row = lines_desired[2], Column = 1).Range.Text.rstrip('\r\x07'))
            listateste.append(table.Cell(Row = lines_desired[2], Column = 2).Range.Text.rstrip('\r\x07'))
            listateste.append(table.Cell(Row = lines_desired[2], Column = 3).Range.Text.rstrip('\r\x07'))
            listateste.append(table.Cell(Row = lines_desired[3], Column = 1).Range.Text.rstrip('\r\x07'))
            listateste.append(table.Cell(Row = lines_desired[3], Column = 2).Range.Text.rstrip('\r\x07'))
            listateste.append(table.Cell(Row = lines_desired[3], Column = 3).Range.Text.rstrip('\r\x07'))
            listateste.append(table.Cell(Row = lines_desired[4], Column = 1).Range.Text.rstrip('\r\x07'))
            listateste.append(table.Cell(Row = lines_desired[4], Column = 2).Range.Text.rstrip('\r\x07'))
            listateste.append(table.Cell(Row = lines_desired[4], Column = 3).Range.Text.rstrip('\r\x07'))
            listateste.append(table.Cell(Row = lines_desired[5], Column = 1).Range.Text.rstrip('\r\x07'))
            listateste.append(table.Cell(Row = lines_desired[5], Column = 2).Range.Text.rstrip('\r\x07'))
            listateste.append(table.Cell(Row = lines_desired[5], Column = 3).Range.Text.rstrip('\r\x07'))
            listateste.append(table.Cell(Row = lines_desired[6], Column = 1).Range.Text.rstrip('\r\x07'))
            listateste.append(table.Cell(Row = lines_desired[6], Column = 2).Range.Text.rstrip('\r\x07'))
            listateste.append(table.Cell(Row = lines_desired[6], Column = 3).Range.Text.rstrip('\r\x07'))
            listateste.append(table.Cell(Row = lines_desired[7], Column = 1).Range.Text.rstrip('\r\x07'))
            listateste.append(table.Cell(Row = lines_desired[7], Column = 2).Range.Text.rstrip('\r\x07'))
            listateste.append(table.Cell(Row = lines_desired[7], Column = 3).Range.Text.rstrip('\r\x07'))
            listateste.append(table.Cell(Row = lines_desired[11], Column = 1).Range.Text.rstrip('\r\x07'))
            listateste.append(table.Cell(Row = lines_desired[12], Column = 1).Range.Text.rstrip('\r\x07'))
            listateste.append(table.Cell(Row = lines_desired[13], Column = 1).Range.Text.rstrip('\r\x07'))
            listateste.append(table.Cell(Row = lines_desired[14], Column = 1).Range.Text.rstrip('\r\x07'))
            listateste.append(table.Cell(Row = lines_desired[15], Column = 1).Range.Text.rstrip('\r\x07'))
            listateste.append(table.Cell(Row = lines_desired[8], Column = 2).Range.Text.rstrip('\r\x07'))
            listateste.append(table.Cell(Row = lines_desired[9], Column = 2).Range.Text.rstrip('\r\x07'))
            listateste.append(table.Cell(Row = lines_desired[10], Column = 2).Range.Text.rstrip('\r\x07'))
            listateste.append(table_2.Cell(Row = lines_desired_2[0], Column = 1).Range.Text.rstrip('\r\x07'))
            listateste.append(table_2.Cell(Row = lines_desired_2[1], Column = 1).Range.Text.rstrip('\r\x07'))
            listateste.append(table_2.Cell(Row = lines_desired_2[1], Column = 2).Range.Text.rstrip('\r\x07'))
            listateste.append(table_2.Cell(Row = lines_desired_2[2], Column = 1).Range.Text.rstrip('\r\x07'))
            listateste.append(table_2.Cell(Row = lines_desired_2[3], Column = 1).Range.Text.rstrip('\r\x07'))
            listateste.append(table_2.Cell(Row = lines_desired_2[4], Column = 1).Range.Text.rstrip('\r\x07'))
            listateste.append(table_2.Cell(Row = lines_desired_2[5], Column = 1).Range.Text.rstrip('\r\x07'))
            listateste.append(table_2.Cell(Row = lines_desired_2[6], Column = 1).Range.Text.rstrip('\r\x07'))
            listateste.append(table_2.Cell(Row = lines_desired_2[7], Column = 1).Range.Text.rstrip('\r\x07'))
            listateste.append(table_2.Cell(Row = lines_desired_2[8], Column = 1).Range.Text.rstrip('\r\x07'))
            listateste.append(table_2.Cell(Row = lines_desired_2[8], Column = 2).Range.Text.rstrip('\r\x07'))
            listateste.append(table_2.Cell(Row = lines_desired_2[9], Column = 1).Range.Text.rstrip('\r\x07'))
            sleep(1)
        
        final_list.append(listateste)
    doc.Close(False)
word.Quit()
    
list2 = [x for x in final_list if x != []]

import csv
with open("Lote_2_part1.csv", "w", encoding="utf-8") as f:
    writer = csv.writer(f, quoting=csv.QUOTE_MINIMAL)
    writer.writerow(['UF-Município','Objeto','Número', 'Distrito/Bairro','Título','Nº Anterior', 'Endereço','Subclasse','Origem','Acervo','Classe','Procedência', 'Local no Prédio','Época','Modo de Aquisição/Data', 'Proprietário','Autoria', 'Conjunto com Nº','Responsável Imediato','Material/Técnica','Termos de Indexação', 'Documentação Fotográfica', 'Proteção', 'Proteção Legal', 'Condições de Segurança', 'Estado de Conservação', 'Marcas/Inscrições/Legendas', 'Dimensões(cm)', 'Descrição', 'Especificação do Estado de Conservação', 'Restaurações', 'Restauradores', 'Características Técnicas', 'Características Iconográficas/Ornamentais', 'Dados Históricos', 'Referências Bibliográficas/Arquivísticas', 'Observações', 'Preenchimento Técnico', 'Revisão Técnica', 'Dados Complementares'])
    writer.writerows(list2)
print ('Arquivo CSV gerado no diretorio! Pronto para ajustes!')


