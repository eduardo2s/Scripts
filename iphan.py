# -*- coding: utf-8 -*-
"""
Created on Tue Aug 20 15:10:31 2019

@author: eduardo
"""

from time import sleep
import win32com.client as win32
import glob
import os
from docx.api import Document

lines_desired = [1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17]
word = win32.gencache.EnsureDispatch('Word.Application')
word.Visible = 0

path = 'pasta com os documentos'

final_list=[]
nom_files = []
n = 0
for infile in glob.glob(os.path.join(path, '*.docx') ):
    document = Document(infile)
    doc = word.Documents.Open(infile)
    print (doc.Tables.Count)
    if doc.Tables.Count == 2:
        table = doc.Tables(1)
        table2 = doc.Tables(2)
        linhas = table.Rows.Count + table2.Rows.Count
        #print("{} Processado... {}".format(infile, n))
        n += 1
        for i in range(1 , linhas):
            listateste = []
            listateste.append(table.Cell(Row = lines_desired[1], Column = 3).Range.Text.rstrip('\r\x07'))
            listateste.append(table.Cell(Row = lines_desired[2], Column = 3).Range.Text.rstrip('\r\x07'))
            listateste.append(table.Cell(Row = lines_desired[3], Column = 3).Range.Text.rstrip('\r\x07'))
            listateste.append(table.Cell(Row = lines_desired[4], Column = 3).Range.Text.rstrip('\r\x07'))
            listateste.append(table.Cell(Row = lines_desired[5], Column = 3).Range.Text.rstrip('\r\x07'))
            listateste.append(table.Cell(Row = lines_desired[6], Column = 3).Range.Text.rstrip('\r\x07'))
            listateste.append(table.Cell(Row = lines_desired[7], Column = 3).Range.Text.rstrip('\r\x07'))
            listateste.append(table.Cell(Row = lines_desired[1], Column = 2).Range.Text.rstrip('\r\x07'))
            listateste.append(table.Cell(Row = lines_desired[2], Column = 2).Range.Text.rstrip('\r\x07'))
            listateste.append(table.Cell(Row = lines_desired[3], Column = 2).Range.Text.rstrip('\r\x07'))
            listateste.append(table.Cell(Row = lines_desired[4], Column = 2).Range.Text.rstrip('\r\x07'))
            listateste.append(table.Cell(Row = lines_desired[5], Column = 2).Range.Text.rstrip('\r\x07'))
            listateste.append(table.Cell(Row = lines_desired[6], Column = 2).Range.Text.rstrip('\r\x07'))
            listateste.append(table.Cell(Row = lines_desired[7], Column = 2).Range.Text.rstrip('\r\x07'))
            listateste.append(table.Cell(Row = lines_desired[1], Column = 1).Range.Text.rstrip('\r\x07'))
            listateste.append(table.Cell(Row = lines_desired[2], Column = 1).Range.Text.rstrip('\r\x07'))
            listateste.append(table.Cell(Row = lines_desired[3], Column = 1).Range.Text.rstrip('\r\x07'))
            listateste.append(table.Cell(Row = lines_desired[4], Column = 1).Range.Text.rstrip('\r\x07'))
            listateste.append(table.Cell(Row = lines_desired[5], Column = 1).Range.Text.rstrip('\r\x07'))
            listateste.append(table.Cell(Row = lines_desired[6], Column = 1).Range.Text.rstrip('\r\x07'))
            listateste.append(table.Cell(Row = lines_desired[7], Column = 1).Range.Text.rstrip('\r\x07'))
            listateste.append(table.Cell(Row = lines_desired[8], Column = 2).Range.Text.rstrip('\r\x07'))
            listateste.append(table.Cell(Row = lines_desired[9], Column = 2).Range.Text.rstrip('\r\x07'))
            listateste.append(table.Cell(Row = lines_desired[10], Column = 2).Range.Text.rstrip('\r\x07'))
            listateste.append(table.Cell(Row = lines_desired[11], Column = 1).Range.Text.rstrip('\r\x07'))
            listateste.append(table.Cell(Row = lines_desired[13], Column = 1).Range.Text.rstrip('\r\x07'))
            listateste.append(table.Cell(Row = lines_desired[14], Column = 1).Range.Text.rstrip('\r\x07'))
            listateste.append(table.Cell(Row = lines_desired[15], Column = 1).Range.Text.rstrip('\r\x07'))
            listateste.append(table.Cell(Row = lines_desired[16], Column = 1).Range.Text.rstrip('\r\x07'))
            listateste.append(table2.Cell(Row = lines_desired[0], Column = 1).Range.Text.rstrip('\r\x07'))
            listateste.append(table2.Cell(Row = lines_desired[1], Column = 1).Range.Text.rstrip('\r\x07'))
            listateste.append(table2.Cell(Row = lines_desired[1], Column = 2).Range.Text.rstrip('\r\x07'))
            listateste.append(table2.Cell(Row = lines_desired[2], Column = 1).Range.Text.rstrip('\r\x07'))
            listateste.append(table2.Cell(Row = lines_desired[3], Column = 1).Range.Text.rstrip('\r\x07'))
            listateste.append(table2.Cell(Row = lines_desired[4], Column = 1).Range.Text.rstrip('\r\x07'))
            listateste.append(table2.Cell(Row = lines_desired[5], Column = 1).Range.Text.rstrip('\r\x07'))
            listateste.append(table2.Cell(Row = lines_desired[6], Column = 1).Range.Text.rstrip('\r\x07'))
            listateste.append(table2.Cell(Row = lines_desired[7], Column = 1).Range.Text.rstrip('\r\x07'))
            listateste.append(table2.Cell(Row = lines_desired[8], Column = 1).Range.Text.rstrip('\r\x07'))
            listateste.append(table2.Cell(Row = lines_desired[8], Column = 2).Range.Text.rstrip('\r\x07'))
            listateste.append(table2.Cell(Row = lines_desired[9], Column = 1).Range.Text.rstrip('\r\x07'))
            sleep(1)  
        for p in document.paragraphs:
            listateste.append(p.text)
        final_list.append(listateste)
        
    doc.Close(False)
word.Quit()

# faz ajustes na lista final, retira itens em branco da lista
list2 = [list(filter(None, lst)) for lst in final_list] 

import pandas as pd
df = pd.DataFrame(list2)
df.to_csv('caminho\\iphanv1.csv',index=False)
