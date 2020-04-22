# -*- coding: utf-8 -*-
"""
Extrai os dados relevantes dos documentos DL.
Os documentos devem estar no formato DWG e separados em uma pasta por DL, no mesmo diretório do software.
O software converte os arquivos DWG em DXF, para então obter os dados do documento e os salvar no formato CSV

23 de maio 2017
@author: Fellipe Garcia Marques
"""
import os
import dxfgrabber
import glob
import dl2csv_funcs
import pandas as pd
import sys
import shutil

# Filtra as pastas que correspondem à documentos DL
current_dir = os.getcwd()
dirs = glob.glob('*DL*/')
if len(dirs) == 0:
    print("Não foram encontrados documentos para conversão de dados\n")
    os.raw_input("Pressione Enter para sair...")
    sys.exit(1)

for directory in dirs:
    if (os.path.isdir(os.path.join(current_dir, directory + "DXF"))):
        shutil.rmtree(os.path.join(current_dir, directory + "DXF"))
        os.mkdir(os.path.join(current_dir, directory + "DXF"))
        

# Abre os dados de correspondência entre TAG - Sistema
faixa_csv = 'faixa_tag.csv'
faixas_sistema = dl2csv_funcs.faixa(faixa_csv)

# Converte os arquivos dwg para o formato dxf, caso necessário
for directory in dirs:
    os.chdir(os.path.join(current_dir, directory))
    dwgs = glob.glob('*.dwg')

    print("Convertendo o documento {0} para DXF\n".format(directory[0:-1]))
    try:
        os.chdir("C:\Program Files\ODA\ODAFileConverter_title 21.2.0")
    except OSError:
        print("O conversor de DWG para DXF não foi encontrado no diretório:\n" \
              "C:\Program Files (x86)\ODA\Teigha File Converter 4.02.2\n")
        sys.exit(1)
    config = '\"ACAD2010\" \"DXF\" \"0\" \"1\"'
    doc = '\"' + os.path.normpath(current_dir) + '\\' + os.path.normpath(directory) + '\"'
    dest = doc[0:-1] + '\DXF\"'
    os.system("ODAFileConverter.exe " + doc + ' ' + dest + ' ' + config)
    os.chdir(current_dir)

# Obtém os dados de cada página e escreve em um arquivo csv
for directory in dirs:
    os.chdir(os.path.join(current_dir, directory + '/DXF'))
    dxfs = glob.glob('*.dxf*')

    print("\r\n\nObtendo os dados do documento: {0}\r\n".format(directory[0:-1]))
    dados = []
    total_folhas = len(dxfs)
    for folha, dxf in enumerate(dxfs):
        dxf_reader = dxfgrabber.readfile(dxf)
        # obtém os dados da página e retorna como um objeto, gerando uma lista
        dados.append(dl2csv_funcs.find_data(dxf_reader, faixas_sistema))
        # indica barra de progresso para a obtenção dos dados do documento
        dl2csv_funcs.imprime_progresso(folha, total_folhas)

    # corrige a dimensão das observações, para escrever no arquivo csv
    max_obs = dl2csv_funcs.fix_dim_obs(dados)

    os.chdir(current_dir)
    # Salva os dados
    print("\r\nEscrevendo dados no arquivo {0}\r\n".format(directory[0:-1] + '.xlsx'))
    columns = ['Sistema', 'Sistema Faixa TAG', 'TAG', 'Tipo']
    columns.extend(['Observação ' + str(i) for i in range(max_obs)])
    columns.extend(['Painel', 'Controlador', 'Título', 'Folha'])
    df_builder = []
    for dado in dados:
        if len(dado.tags) > 0:
            for index, tag in enumerate(dado.tags):
                row = []
                row = [dado.sistemas[index], dado.sistemas_faixa[index], dado.tags[index], dado.tipo[index]]
                row.extend(dado.observacoes[index])
                row.extend([dado.paineis[index], dado.controlador[index],
                            dado.titulo[index], dado.folha[index]])
                df_builder.append(row)
    df = pd.DataFrame(df_builder, columns=columns)
    df.to_excel(directory[0:-1] + '.xlsx')
    
