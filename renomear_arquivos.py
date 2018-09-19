# -*- coding: utf-8 -*-
"""
Created on Mon Sep 17 21:10:15 2018

@author: eduarado
"""

import os
import ntpath

# Caminho para a pasta com subpastas/imagens
path = "C:\\Users\\eduar\\Desktop\\ficahs\\novas fichas\\Imagens\\"

#acessa o diretorio de forma recursiva para acessar as subpastas
for root, dir, files in os.walk(path):
  #acessa as subpastas para acessas os arquivos dentro delas
  for file in files:
      dirname = ntpath.basename(root)
      #caminho original
      ori = root + '\\' + file
      #adiciona o nome da pasta + underscore + a posição do arquivo + a extensão
      name, ext = os.path.splitext(file) 
      dest = dirname + '_' + str(files.index(file)) + ext
      os.rename(ori, dest)
   



        
