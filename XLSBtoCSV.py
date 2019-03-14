# -*- coding: utf-8 -*-
import csv
import glob
import shutil
import os
from pyxlsb import open_workbook

# Inicia Leitura dos arquivos contidos no diretorio
files = []
nomesheet = []
files = (glob.glob(caminhoxlsb+'*.xlsb'))
filesSize = len((glob.glob('*.xlsb')))
i = 0
print('Processo Iniciado')
for filesSize in files:
    with open_workbook(files[i]) as arquivo:
        print('Convertendo: ' + files[i])
        for name in arquivo.sheets:
            with arquivo.get_sheet(name) as sheet, open(name + '.csv', 'w', newline='', encoding='UTF-8') as f:
                writer = csv.writer(f, delimiter=';')
                for row in sheet.rows():
                    writer.writerow([c.v for c in row])
            # pega o nomesheet do csv
            nomesheet = (name + '.csv')
            print(nomesheet + ' Convertido com Sucesso!')    
    i+=1

print(f'Processo Finalizado! {i} Arquivos Convertidos ')