print('Iniciando...')

import os

import pathlib
import pandas as pd

# requirements: pathlib, pandas, openpyxl,  xlsxwriter, pyinstaller


# TODO: criar filedialog para selecionar pasta desejada
# TODO: criar filedialog para passar lista de arquivos desejados

# ------------FUNCOES--------------------


def tratar_nome(nome_arquivo:str) -> str:
    '''Faz tratamentos no nome do arquivo, removendo caracteres
    e espaços desnecessários. Substitui o primeiro e ultimo espaços por underline.

    Parametros:
        nome_arquivo (str): nome do arquivo a ser tratado

    Retorna:
        nome_formatado (str): nome do arquivo após tratamentos de texto.
    '''
    nome_arquivo = nome_arquivo.strip()
    nome_arquivo = nome_arquivo.upper()
    nome_arquivo = nome_arquivo.replace('ASO ', 'ASO_')
    nome_arquivo = nome_arquivo.replace('FC ', 'FC_')
    nome_arquivo = nome_arquivo.replace('EC ', 'EC_')
    nome_formatado = nome_arquivo.replace('.PDF', '_PDF')
    return nome_formatado


def criar_linha_tabela(aux: list) -> list:
    '''Separa o nome do arquivo em uma lista com 4 itens:
    Tipo, Nome, Data e Extensao.
    Se uma data nao for encontrada, insere (falta data)
    no terceiro item.

    Parametros:
        aux (list): nome do arquivo em forma de lista

    Retorna:
        nome_arquivo (list): nome do arquivo separado em 4 partes
    '''
    for indice, char in enumerate(aux):
        # encontrar o primeiro digito
        if char.isdigit():
            aux[indice - 1] = '_'  # _ antes do dia (2 digitos)
            aux.insert(indice + 2, '-')  # mes (2 digitos)
            aux.insert(indice + 5, '-')  # ano (4 digitos)
            # aux.insert(indice + 10, '_')  # apos o ano
            aux = ''.join(aux)
            nome_arquivo = aux.split('_', 3)  # retornar lista com 4 itens
            break
        
        # se nao houver digitos: preencher coluna da data com '(falta data)'
        elif indice == (len(aux)-1):
            nome_arquivo = ''.join(aux)
            nome_arquivo = nome_arquivo.replace('_PDF', '_(falta data)_PDF')
            nome_arquivo = nome_arquivo.split('_', 3)
    return nome_arquivo


def checar_data(lista_arquivos: list) -> list:
    '''Confere o numero de caracteres no terceiro item do nome
    do arquivo. Se houver erro no tamanho da data inclui observacao no 
    final da lista. Se nao, inclui espaco vazio.

    Parametros:
        lista_arquivos (list): matriz com listas de 4 itens

    Retorna:
        lista_arquivos (list): matriz com listas de 5 itens
    '''
    for indice, item in enumerate(lista_arquivos):
        if len(lista_arquivos[indice][2]) != 10:
            lista_arquivos[indice].append('Erro na Data')
        else:
            lista_arquivos[indice].append('')
    return lista_arquivos


# ------------PROGRAMA PRINCIPAL--------------------


# pegar current working directory
pasta_atual = os.getcwd()
caminho = pathlib.Path(r'{}'.format(pasta_atual))

lista_arquivos = []
lista_outros = []

print('Copiando Nomes...')

# iterar sobre todos os arquivos da pasta
for arquivo in caminho.iterdir():
    nome_arquivo = arquivo.name  # capturar nome do arquivo
    if 'ASO ' in nome_arquivo or 'FC ' in nome_arquivo or 'EC ' in nome_arquivo:
        # tratar nome do arquivo e transformar string em lista
        aux = list(tratar_nome(nome_arquivo))
        # separar nome do arquivo em lista de 4 itens
        linha_tabela = criar_linha_tabela(aux)
        # incluir nome na lista
        lista_arquivos.append(linha_tabela) 
    elif '.pdf' in nome_arquivo or '.jpg' in nome_arquivo or '.jpeg' in nome_arquivo:
        # pegar nome do arquivo separar extensao
        nome_arquivo = nome_arquivo.replace('.', '_')
        nome_arquivo = nome_arquivo.split('_', 1)
        # append em lista_outros
        lista_outros.append(nome_arquivo)
       
# checar formato da data e criar coluna Obs
lista_tratada = checar_data(lista_arquivos) # 5 colunas

print(f'{len(lista_tratada)} nomes copiados')
print(f'{len(lista_outros)} arquivos não renomeados encontrados')

# criar df's
df_arquivos = pd.DataFrame(data=lista_tratada, columns=('Tipo', 'Nome do Arquivo', 'Data', 'Extensao', 'Obs'))
df_arquivos.sort_values(by='Nome do Arquivo', inplace=True)  # ordem alfabetica

df_outros = pd.DataFrame(data=lista_outros, columns=('Nome do Arquivo', 'Extensao'))
df_outros.sort_values(by='Nome do Arquivo', inplace=True)

# criar planilha
writer = pd.ExcelWriter(path='Nomes dos Arquivos.xlsx', engine='xlsxwriter')

df_arquivos.to_excel(writer, sheet_name='Sheet1', index=False)
df_outros.to_excel(writer, sheet_name='Outros', index=False)

writer.save()

print('Arquivo Excel criado')
print('Concluído!')
sair = input('Pressione Enter para Sair\n')
