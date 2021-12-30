print('Iniciando...')

import pathlib
import os
import pandas as pd

# requirements: pathlib, pandas, openpyxl, pyinstaller


# TODO: criar procedimentos para casos de FC, EC e arquivos aleatorios (FEITO)
# TODO: criar filedialog para selecionar pasta desejada
# TODO: criar filedialog para passar lista de arquivos desejados
# TODO: criar procedimentos para evitar erros por causa de arquivos duplicados...
# ...(com (2) ou outros numeros no final) (FEITO)
# TODO: criar procedimentos para evitar datas digitadas errado (com mais de 8 digitos) (FEITO)
# TODO: adicionar coluna de extensão do arquivo (pdf, xlsx, word, etc) (FEITO)


# pegar current working directory
# pasta_atual = os.getcwd()
pasta_atual = '/home/gabriel/Desktop/Hashtag-Python/GRS integracao SOC APIs/Teste PDF'
caminho = pathlib.Path(r'{}'.format(pasta_atual))

lista_arquivos = []

print('Copiando Nomes...')


# iterar sobre todos os arquivos da pasta
for arquivo in caminho.iterdir():
    nome_arquivo = arquivo.name  # capturar nome do arquivo

    if 'ASO ' in nome_arquivo or 'FC ' in nome_arquivo or 'EC ' in nome_arquivo:
        # tratamentos de texto
        nome_arquivo = nome_arquivo.strip()
        nome_arquivo = nome_arquivo.upper()
        nome_arquivo = nome_arquivo.replace('ASO ', 'ASO_')
        nome_arquivo = nome_arquivo.replace('FC ', 'FC_')
        nome_arquivo = nome_arquivo.replace('EC ', 'EC_')
        nome_arquivo = nome_arquivo.replace('.PDF', '_PDF')

        # transformar o string em lista
        aux = list(nome_arquivo)
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

        lista_arquivos.append(nome_arquivo)  # incluir nome na lista


# checar formato da data (coluna Obs)
for indice, item in enumerate(lista_arquivos):
    if len(lista_arquivos[indice][2]) != 10:
        lista_arquivos[indice].append('Erro na Data')
    else:
        lista_arquivos[indice].append('')


print(f'{len(lista_arquivos)} nomes copiados')

# criar tabela de nomes
df = pd.DataFrame(lista_arquivos, columns=('Tipo', 'Nome do Arquivo', 'Data', 'Extensao', 'Obs'))
df.sort_values('Nome do Arquivo', inplace=True)  # ordem alfabetica
df.to_excel('Nomes dos Arquivos.xlsx', index=False)  # para excel sem indice

print('Arquivo Excel criado')
print('Concluído!')
sair = input('Pressione Enter para Sair\n')
