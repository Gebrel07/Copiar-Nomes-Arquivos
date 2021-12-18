#!/usr/bin/env python
# coding: utf-8

# In[4]:


print('Iniciando...')

import pathlib
import os
import pandas as pd

# requirements: pathlib, pandas, openpyxl, pyinstaller


# In[5]:


#pegar current working directory
pasta_atual = os.getcwd()
caminho = pathlib.Path(r'{}'.format(pasta_atual))

lista_arquivos = []

print('Copiando Nomes...')

#iterar sobre todos os arquivos da pasta
for arquivo in caminho.iterdir():
    nome_arquivo = arquivo.name #capturar nome do arquivo
    
    # TODO: criar procedimentos para casos de FC, EC e arquivos aleatorios (FEITO)
    # TODO: criar filedialog para selecionar pasta desejada
    # TODO: criar filedialog para passar lista de arquivos desejados
    # TODO: criar procedimentos para evitar erros por causa de arquivos duplicados (com (2) ou outros numeros no final) (FEITO)
    # TODO: criar procedimentos para evitar datas digitadas errado (com mais de 8 digitos) (FEITO)
    # TODO: adicionar coluna de extensão do arquivo (pdf, xlsx, word, etc) (FEITO)
    
    if 'ASO ' in nome_arquivo or 'FC 'in nome_arquivo or 'EC ' in nome_arquivo:
        # tratamentos de texto
        nome_arquivo = nome_arquivo.strip()
        nome_arquivo = nome_arquivo.upper()
        nome_arquivo = nome_arquivo.replace('ASO ', 'ASO_')
        nome_arquivo = nome_arquivo.replace('FC ', 'FC_')
        nome_arquivo = nome_arquivo.replace('EC ', 'EC_')
        
        # transformar o string em lista
        aux = list(nome_arquivo)
        for indice, char in enumerate(aux):
            # encontrar o primeiro digito
            if char.isdigit():
                aux[indice-1] = '_' # _ antes do dia (2 digitos)
                aux.insert(indice+2, '/') # mes (2 digitos)
                aux.insert(indice+5, '/') # ano (4 digitos)
                aux.insert(indice+10, '_') # apos o ano
                aux = ''.join(aux)
                nome_arquivo = aux.split('_', 3) # retornar lista com 4 itens
                break
                
        lista_arquivos.append(nome_arquivo) #incluir nome na lista
        
# checar formato da data (coluna Obs)
for indice, item in enumerate(lista_arquivos):
    if len(lista_arquivos[indice][2]) != 10:
        lista_arquivos[indice].append('Erro na Data')

print(f'{len(lista_arquivos)} nomes copiados')


# In[8]:


#criar tabela de nomes
df = pd.DataFrame(lista_arquivos, columns = ('Tipo', 'Nome do Arquivo', 'Data', 'Extensao', 'Obs'))
df.sort_values('Nome do Arquivo', inplace = True) #ordem alfabetica
df.to_excel('Nomes dos Arquivos.xlsx', index = False) #para excel sem indice

print('Arquivo Excel criado')
print('Concluído!')
sair = input('Pressione Enter para Sair\n')


# In[ ]:




