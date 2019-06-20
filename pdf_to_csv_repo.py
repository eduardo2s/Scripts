# -*- coding: utf-8 -*-
"""
Created on Thu Jun 20 10:43:17 2019

@author: eduar
"""
from PyPDF2 import PdfFileWriter, PdfFileReader
import re
import slate3k as slate
import pandas as pd

# primeira etapa, remova páginas desnecessárias e salva num novo arquivo
pages_to_delete = [0, 1, 754, 755, 756,757,758,759,760,761] # começa do 0
infile = PdfFileReader('C:\\Users\\eduar\\Downloads\\catalogoMNtapeçaria.pdf', 'rb')
output = PdfFileWriter()

for i in range(infile.getNumPages()):
    if i not in pages_to_delete:
        p = infile.getPage(i)
        output.addPage(p)

with open('C:\\Users\\eduar\\Desktop\\catalogoMNapenasitens.pdf', 'wb') as f:
    output.write(f)
    

# segunda etapa abre o arquivo já como texto
with open('C:\\Users\\eduar\\Desktop\\catalogoMNapenasitens.pdf','rb') as f:
    extract = slate.PDF(f)
    extracted_text = [extract[i] for i in range(len(extract))]

# remove quebra de linhas
extracted_text = [x.replace('\n', ' ') for x in extracted_text]
# trata espaços duplicados. Verificar se não é necessário rodar mais de uma vez 
extracted_text = [x.replace('  ', ' ') for x in extracted_text]

titulo = [re.findall(r'(.*)(?= Medidas: )', text) if re.findall(r'(.*)(?= Procedencia: )', text) == '' else re.findall(r'(.*)(?=Procedencia:)', text) for text in extracted_text]
titulo = [''.join(e) for e in titulo]

procedencia = [re.findall(r'(?<= Procedencia: )(.*?)(?= Medidas: )', text) for text in extracted_text]
procedencia = [''.join(e) for e in procedencia]

medidas = [re.findall(r'(?<= Medidas: )(.*?)(?= Hilos: )', text) for text in extracted_text]
medidas = [''.join(e) for e in medidas]

hilos = [re.findall(r'(?<= Hilos: )(.*?)(?= Técnica: )', text) for text in extracted_text]
hilos = [''.join(e) for e in hilos]

tecnicas = [re.findall(r'(?<= Técnica: )(.*?)(?= Estructuras: )', text) for text in extracted_text]
tecnicas = [''.join(e) for e in tecnicas]

estruturas = [re.findall(r'(?<= Estructuras: )(.*?)(?= Descripción: )', text) for text in extracted_text]
estruturas = [''.join(e) for e in estruturas]

descricao = [re.findall(r'(?<= Descripción: )(.*?)(?= Comentarios: )', text) for text in extracted_text]
descricao = [''.join(e) for e in descricao]

comentarios = [re.findall(r'(?<= Comentarios: )(.*?)(?= Pieza N°: )', text) for text in extracted_text]
comentarios = [''.join(e) for e in comentarios]

referencias = [re.findall(r'(?<= Referencias: )(.*?)(?= Pieza N°: )', text) for text in extracted_text]
referencias = [''.join(e) for e in referencias]

n_peca = [re.findall(r'(?<= Pieza N°: )(.*?)(?= Foto: )', text) for text in extracted_text]
n_peca = [''.join(e) for e in n_peca]

foto = [re.findall(r'(?<= Foto: )(.*?)(?=   )', text) for text in extracted_text]
foto = [''.join(e) for e in foto]

# cria um dataframe para exportar como csv
df = pd.DataFrame({
    'Título':titulo, 
    'Procedencia':procedencia,
    'Medidas':medidas,
    'hilos':hilos,
    'Técnica':tecnicas,
    'Estructuras':estruturas,
    'Descripción':descricao,
    'Comentários':comentarios,
    'Referencias':referencias,
    'Pieza N°':n_peca,
    'Foto':foto,
}) 

df.to_csv(r'C:\\Users\\eduar\\Desktop\\catalogoMNtapeçaria.csv', index=False)

    
