# -*- coding: utf-8 -*-
"""
Created on Thu Mar 21 22:30:50 2019

@author: eduardo
"""

import zipfile
from PIL import Image
import io
import os
from tqdm import tqdm

# Define o Caminho
path = "caminho"


def redact_images(filename,FilePath):
    #outfile = filename.replace(".docx", "_redacted.docx")
    with zipfile.ZipFile(filename) as inzip:
        #with zipfile.ZipFile(outfile, "w") as outzip:
            i = 0
            for info in inzip.infolist():
                name = info.filename
                content = inzip.read(info)
                if name.endswith((".png", ".jpeg", ".gif")):
                        fmt = name.split(".")[-1]
                        Name = name.split("/")[-1]
                        img = Image.open(io.BytesIO(content))
                        img.save(FilePath + str(Name))
                        outb = io.BytesIO()
                        img.save(outb, fmt)
                        content = outb.getvalue()
                        info.file_size = len(content)
                        info.CRC = zipfile.crc32(content)
                        i += 1
                #outzip.writestr(info, content)

               
for filename in tqdm(os.listdir(path), desc="processando dados", unit="files"):
    if filename.endswith('.docx'):
        #print(filename)
        # separa o nome do arquivo e dua .extensão
        name, ext = os.path.splitext(filename)
        # caminho para o arquivo
        file = path + filename
        # armazena o caminho onde o dados serão salvos
        filepath = (path + name + '\\')
        # cria um diretorio com o nome do arquivo (no caso das fichas o nº do item)
        os.makedirs(path + name)
        # executa a função cirada antes, onde é definido o caminho do arquivo e onde as imagens serão salvas
        redact_images(file, filepath)
