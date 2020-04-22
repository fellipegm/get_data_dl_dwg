# -*- coding: utf-8 -*-
"""
Funções do dl2csv
Extrai os dados relevantes dos documentos DL.
Os documentos devem estar no formato DWG e separados em uma pasta por DL, no mesmo diretório do software.
O software converte os arquivos DWG em DXF, para então obter os dados do documento e os salvar no formato CSV

23 de maio 2017
@author: Fellipe Garcia Marques
"""

from collections import Counter
import re
import csv
import itertools
import math
import sys


# Define a classe DadosPagina, que contém todos os dados extraídos de uma determinada página
class DadosPagina(object):
    def __init__(self, tags, sistemas, tipo, observacoes, paineis, controlador,
                 titulo, folha, faixas_sistema):
        # Concatena os TAGs dos instrumentos em uma única string e encontra o sistema equivalente à faixa de TAGs
        tag_lista = []
        if len(tags) > 0:
            self.sistemas_faixa = []
            for tag in tags:
                if re.match('^[0-9].*', tag[0]):
                    tag_lista.append(tag[1] + '-' + tag[0])
                    tag_filt = re.sub('([A-Z](.|\\Z)+|\.(.|\\Z)+)', '', tag[0])
                    tag_filt = re.sub('/.*', '', tag_filt)
                    for faixa_sistema in faixas_sistema:
                        try:
                            if int(faixa_sistema[2]) >= int(tag_filt) >= int(faixa_sistema[1]):
                                self.sistemas_faixa.append(faixa_sistema[0])
                                break
                            elif faixa_sistema == faixas_sistema[-1]:
                                self.sistemas_faixa.append('')
                        except ValueError:
                            self.sistemas_faixa.append('')
                else:
                    tag_lista.append(tag[0] + '-' + tag[1])
                    tag_filt = re.sub('([A-Z](.|\\Z)+|\.(.|\\Z)+)', '', tag[1])
                    tag_filt = re.sub('/.*', '', tag_filt)
                    for faixa_sistema in faixas_sistema:
                        try:
                            if int(faixa_sistema[2]) >= int(tag_filt) >= int(faixa_sistema[1]):
                                self.sistemas_faixa.append(faixa_sistema[0])
                                break
                            elif faixa_sistema == faixas_sistema[-1]:
                                self.sistemas_faixa.append('')
                        except ValueError:
                            self.sistemas_faixa.append('')
        self.tags = tag_lista

        # garante que a lista sistemas tem a mesma dimensão da lista de TAGs
        if len(sistemas) < len(tag_lista):
            for i in range(len(tag_lista) - len(sistemas)):
                sistemas.append('')
        self.sistemas = sistemas

        self.tipo = tipo
        self.observacoes = limpa(observacoes)

        # garante que a lista paineis tem a mesma dimensão da lista de TAGs
        if len(paineis) < len(tag_lista):
            for i in range(len(tag_lista) - len(paineis)):
                paineis.append('')
        self.paineis = limpa(paineis)

        self.controlador = limpa(controlador)
        self.titulo = limpa(titulo)
        self.folha = folha


# obtém os dados de equivalência entre faixa de TAG - sistema do arquivo csv
def faixa(csv_file):
    with open(csv_file, 'r') as csvfile:
        row_reader = csv.reader(csvfile, delimiter=';')
        faixas_sistema = []
        for row in row_reader:
            faixas_sistema.append([row[0], row[1], row[2]])
    return faixas_sistema


# função utilizada para que todas as observações dos TAGs tenham a mesma dimensão
def fix_dim_obs(dados):
    obs_column_max = 0
    for dado in dados:
        for obs in dado.observacoes:
            obs_column_max = max(obs_column_max, len(obs))
    for index1, dado in enumerate(dados):
        for index2, observacao in enumerate(dado.observacoes):
            if len(dado.observacoes[index2]) < obs_column_max:
                for k in range(obs_column_max - len(dados[index1].observacoes[index2])):
                    dados[index1].observacoes[index2].append('')
    return obs_column_max


# constante que define os escapes do string, que devem ser filtrados
escapes = ''.join([chr(char) for char in range(1, 32)])


# função para limpar caracteres indesejados
def limpa(txts):
    for index1, txt in enumerate(txts):
        if type(txt) == str:
            txts[index1] = re.sub('\(NOTA.*?\)', '', txts[index1])
            txts[index1] = re.sub('\\s+\\Z', '', txts[index1])
        else:
            for index2, sub_txt in enumerate(txt):
                txts[index1][index2] = re.sub('\(NOTA.*?\)', '', txts[index1][index2])
                txts[index1][index2] = re.sub('\\s+\\Z', '', txts[index1][index2])
    return txts


# divide a string em substrings para os separadores sep
def split(txt, seps):
    default_sep = seps[0]
    for sep in seps[1:]:
        txt = txt.replace(sep, default_sep)
    return [i.strip() for i in txt.split(default_sep)]


# obtém os pontos médios de uma lista com 4 entradas
def sort_middle(pontos):
    lista = []
    for ponto in pontos:
        lista.append(ponto[0])
        lista.append(ponto[1])
    lista_sorted = sorted(lista)
    return lista_sorted[1], lista_sorted[2]


# lista que define os textos do cabeçalho de entrada e saída do documento DWG
txt_entrada_saida = [u'ENTRADA', u'SAÍDA']


def find_data(dxf,  # arquivo dxf
              faixas_sistema,  # list contendo os dados de faixas de tags para determinado sistema
              ):
    """ Retorna as entradas e saídas do documento """

    # se a folha for cancelada, não retorna nada
    cancelada = [entity for entity in dxf.entities if (entity.dxftype == 'TEXT' or entity.dxftype == 'MTEXT') and
                 entity.plain_text() == u'CANCELADA']
    if len(cancelada) != 0:
        return DadosPagina('', '', '', '', '', '', '', '', '')

    # obtém as entidades de texto com o cabeçalho de entrada e saída
    entrada_saida_dxfs_class = [entity for entity in dxf.entities if
                               (entity.dxftype == u'TEXT' or entity.dxftype == u'MTEXT') and
                               (entity.plain_text() == txt_entrada_saida[0] or
                                entity.plain_text() == txt_entrada_saida[1])]

    # se não encontrou entrada e saída, não retorna nada, visto que há um problema na página
    if len(entrada_saida_dxfs_class) < 2:
        return DadosPagina('', '', '', '', '', '', '', '', '')

    # encontra as linhas do documento e separa em linhas verticais e horizontais
    lines = [entity for entity in dxf.entities if (entity.dxftype == u'LINE' or entity.dxftype == u'LWPOLYLINE')]

    vlines = []
    hlines = []
    for line in lines:
        if line.dxftype == u'LINE':
            if abs(line.start[0]-line.end[0]) < 0.1:
                vlines.append([(float(line.start[0]), float(line.start[1])), (float(line.end[0]), float(line.end[1]))])
            if abs(line.start[1] - line.end[1]) < 0.1:
                hlines.append([(float(line.start[0]), float(line.start[1])), (float(line.end[0]), float(line.end[1]))])
        else:
            for points in itertools.combinations(line.points, 2):
                if abs(points[0][0] - points[1][0]) < 0.1:
                    vlines.append(
                        [(float(points[0][0]), float(points[0][1])), (float(points[1][0]), float(points[1][1]))])
                if abs(points[0][1] - points[1][1]) < 0.1:
                    hlines.append(
                        [(float(points[0][0]), float(points[0][1])), (float(points[1][0]), float(points[1][1]))])

    # determina os limites dos cabeçalho de entrada e saída
    xlim_headers = {}
    for entrada_saida in entrada_saida_dxfs_class:
        xmax = 1e20
        xmin = 0
        for vline in vlines:
            if entrada_saida.insert[0] < vline[0][0] and \
                    (entrada_saida.insert[1] < vline[0][1] or entrada_saida.insert[1] < vline[1][1]):
                xmax = min(xmax, vline[0][0])
                if xmax == vline[0][0]:
                    vline_y_io = vline
            elif entrada_saida.insert[0] > vline[0][0] and \
                    (entrada_saida.insert[1] < vline[0][1] or entrada_saida.insert[1] < vline[1][1]):
                xmin = max(xmin, vline[0][0])
                if xmin == vline[0][0]:
                    vline_y_io = vline
        xlim_headers[entrada_saida.plain_text()] = (xmax, xmin)

    ylim_headers = {}
    for entrada_saida in entrada_saida_dxfs_class:
        ymax = 1e20
        ymin = 0
        for hline in hlines:
            if entrada_saida.insert[1] < hline[0][1] and \
                    (entrada_saida.insert[0] < hline[0][0] or entrada_saida.insert[0] < hline[1][0]):
                ymax = min(ymax, hline[0][1])
                if ymax == hline[0][1]:
                    hline_x_io = hline
            elif entrada_saida.insert[1] > hline[0][1] and \
                    (entrada_saida.insert[0] < hline[0][0] or entrada_saida.insert[0] < hline[1][0]):
                ymin = max(ymin, hline[0][1])
                if ymin == hline[0][1]:
                    hline_x_io = hline
        ylim_headers[entrada_saida.plain_text()] = (ymax, ymin)

    try:
        hline_x_io
    except NameError:
        return DadosPagina('', '', '', '', '', '', '', '', '')
    # define os limites dos IOs de entrada e saída, baseado nas linhas verticais e horizontais
    io_xlim_headers = {}
    io_xlim_headers['ENTRADA'] = (xlim_headers['ENTRADA'][0], min(hline_x_io[0][0], hline_x_io[1][0]))
    io_xlim_headers['SAÍDA'] = (max(hline_x_io[0][0], hline_x_io[1][0]), xlim_headers['SAÍDA'][1])
    io_ylim_headers = {}
    io_ylim_headers['ENTRADA'] = (ylim_headers['ENTRADA'][1], min(vline_y_io[0][1], vline_y_io[1][1]))
    io_ylim_headers['SAÍDA'] = (ylim_headers['SAÍDA'][1], min(vline_y_io[0][1], vline_y_io[1][1]))

    # encontra todos os textos da página
    textos_dxf_class = [entity for entity in dxf.entities if
                        (entity.dxftype == 'TEXT' or entity.dxftype == 'MTEXT')]

    # encontra o título da página
    px_tit_min, px_tit_max = sort_middle(xlim_headers.values())
    py_tit_min, py_tit_max = sort_middle(ylim_headers.values())
    for texto in textos_dxf_class:
        if px_tit_min < texto.insert[0] < px_tit_max and py_tit_min < texto.insert[1] < py_tit_max and \
                not re.match('(\\A\\s{2,}\\Z|[0-9]\\Z|\.\\Z)', texto.plain_text()):
            titulo = texto.plain_text()

    # encontra o controlador da página
    cont_line = 0
    for hline in hlines:
        if hline[0][1] < py_tit_min and (hline[0][0] < px_tit_min or hline[1][0] < px_tit_min):
            cont_line = max(cont_line, hline[0][1])
    for texto in textos_dxf_class:
        if py_tit_min > texto.insert[1] > cont_line and px_tit_max > texto.insert[0] > px_tit_min and \
                not re.match('(\\A\\s{2,}\\Z|[0-9]\\Z|\.\\Z)', texto.plain_text()):
            controlador = texto.plain_text()

    # encontra a folha
    for entity in dxf.entities:
        if re.match(u'INSERT', entity.dxftype):
            if re.match(u'CARIMBO.*?', entity.name):
                carimbo = entity
    try:
        folha = carimbo.attribs[6].text
    except IndexError:
        folha = ''

    # encontra os IOs de entrada ou saída, os painéis e as observações
    tags_ios = []
    txt_painel_io = []
    obs_tag_io = []
    tipo = []
    for entrada_saida in entrada_saida_dxfs_class:
        # filtra os I/Os, como objeto
        ios = [entity for entity in dxf.entities if entity.dxftype == 'INSERT' and
               io_xlim_headers[entrada_saida.plain_text()][1] < entity.insert[0] <
               io_xlim_headers[entrada_saida.plain_text()][0] and
               io_ylim_headers[entrada_saida.plain_text()][1] < entity.insert[1] <
               io_ylim_headers[entrada_saida.plain_text()][0] and
               not re.match('TESTE.+|CARIM.+', entity.name) and
               len(entity.attribs) > 1]
        # se não encontrar IO, vai para o próximo cabeçalho
        if len(ios) == 0:
            continue

        # obtém os tags em formato string
        for io in ios:
            try:
                tags_ios.append([io.attribs[0].text,
                                 io.attribs[1].text])
            except IndexError:
                'do nothing'

        # procura os textos da coluna de IO
        textos_io = [texto for texto in textos_dxf_class if
                     io_xlim_headers[entrada_saida.plain_text()][1]-10 < texto.insert[0] <
                     io_xlim_headers[entrada_saida.plain_text()][0]+10 and
                     io_ylim_headers[entrada_saida.plain_text()][1] < texto.insert[1] <
                     io_ylim_headers[entrada_saida.plain_text()][0] and
                     not re.match('(\\A\\s{3,}|[0-9]\\Z|\.\\Z|NOTA.*?)', texto.plain_text()) and
                     texto.layer != u'CARIMBO MATERIAIS']


        # determina painel de cada entrada
        paineis_dxf_class = [texto for texto in textos_dxf_class if
                     io_xlim_headers[entrada_saida.plain_text()][1] < texto.insert[0] <
                     io_xlim_headers[entrada_saida.plain_text()][0] and
                     io_ylim_headers[entrada_saida.plain_text()][1] < texto.insert[1] <
                     io_ylim_headers[entrada_saida.plain_text()][0] and
                     not re.match('(\\A\\s{3,}|[0-9]\\Z|\.\\Z|NOTA.*?)', texto.plain_text()) and
                     texto.layer == u'CARIMBO MATERIAIS']

        # encontra posições máxima e mínima dos IOs
        pos_min_io = 1e20
        pos_max_io = 0
        for index, io in enumerate(ios):
            pos_min_io = min(pos_min_io, io.insert[1])
            if pos_min_io == io.insert[1]:
                min_io = io
                index_io_min = index
            pos_max_io = max(pos_max_io, io.insert[1])
            if pos_max_io == io.insert[1]:
                max_io = io
                index_io_max = index

        # determina o painel de cada IO
        dados_pos_io = paineis_dxf_class
        dados_pos_io.append(min_io)
        for io in ios:
            for index, dado_pos_io in enumerate(dados_pos_io):
                if dado_pos_io == dados_pos_io[-1]:
                    txt_painel_io.append('')
                    break
                # foi necessário adicionar uma constante para obter textos com referência fora do campo de IO
                if dado_pos_io.insert[1] > io.insert[1] > dados_pos_io[index + 1].insert[1] - 10:
                    txt_painel_io.append(dado_pos_io.plain_text())
                    break

        # determina observações de cada entrada

        # cada linha da matriz de distâncias corresponde a uma entidade texto
        # cada coluna contém a distância do texto para um IO e o ângulo do IO para o texto
        matriz_distancias = []  # [texto_i][(dist_io_j, angulo_io_j)]
        for texto in textos_io:
            mat_ang_modulo_io_obs = []
            for io in ios:
                y_io = (io.attribs[0].insert[1] + io.attribs[1].insert[1])/2
                dist_io = math.sqrt((io.attribs[0].insert[0]-texto.insert[0])**2 + (y_io-texto.insert[1])**2)
                if texto.insert[0] > y_io and texto.insert[1] > io.attribs[0].insert[1]:
                    ang_io = 180.0/math.pi*math.acos((texto.insert[0]-io.attribs[0].insert[0])/dist_io)
                if texto.insert[0] < y_io and texto.insert[1] > io.attribs[0].insert[1]:
                    ang_io = 180-180.0/math.pi*math.acos((io.attribs[0].insert[0]-texto.insert[0])/dist_io)
                if texto.insert[0] < y_io and texto.insert[1] < io.attribs[0].insert[1]:
                    ang_io = 180+180.0/math.pi*math.acos((io.attribs[0].insert[0]-texto.insert[0])/dist_io)
                if texto.insert[0] > y_io and texto.insert[1] < io.attribs[0].insert[1]:
                    ang_io = 360.0-180.0/math.pi*math.acos((texto.insert[0]-io.attribs[0].insert[0])/dist_io)
                mat_ang_modulo_io_obs.append([dist_io, ang_io])
            matriz_distancias.append(mat_ang_modulo_io_obs)

        # determina se há observações acima, abaixo ou acima e abaixo dos IOs
        flag_uptext = False
        flag_downtext = False
        angulo_busca = 20
        for dist_text in matriz_distancias:
            if angulo_busca < dist_text[index_io_max][1] < 180-angulo_busca:
                flag_uptext = True
            if 180+angulo_busca < dist_text[index_io_min][1] < 360-angulo_busca:
                flag_downtext = True

        # determina qual texto corresponde a cada IO
        observacao_ios = []
        for i_txt, texto in enumerate(textos_io):
            dist_txt_min = 1e20
            for i_io, io in enumerate(ios):
                if flag_uptext == flag_downtext:
                    dist_txt_min = min(dist_txt_min, matriz_distancias[i_txt][i_io][0])
                elif flag_uptext:
                    if not 180+angulo_busca < matriz_distancias[i_txt][i_io][1] < 360-angulo_busca:
                        dist_txt_min = min(dist_txt_min, matriz_distancias[i_txt][i_io][0])
                elif flag_downtext:
                    if not angulo_busca < matriz_distancias[i_txt][i_io][1] < 180-angulo_busca:
                        dist_txt_min = min(dist_txt_min, matriz_distancias[i_txt][i_io][0])
                if dist_txt_min == matriz_distancias[i_txt][i_io][0]:
                    i_txt_min = i_txt
                    i_io_min = i_io
            observacao_ios.append([i_io_min, i_txt_min])

        # cria uma lista com todas as observações de cada IO
        lista_observacao_io = []
        for index, io in enumerate(ios):
            lista_observacao_io.append([])
            for k, index_obs in enumerate(observacao_ios):
                if index == index_obs[0]:
                    lista_observacao_io[index].append(textos_io[index_obs[1]].plain_text())
                elif k == len(observacao_ios)-1 and len(lista_observacao_io[index]) < 1:
                    lista_observacao_io[index].append('')
        obs_tag_io.extend(lista_observacao_io)

        # define tipo do I/O
        for io in ios:
            tipo.append(entrada_saida.plain_text())

    # se não foram encontrados tags, retorna vetor nulo
    if len(tags_ios) == 0:
        return DadosPagina('', '', '', '', '', '', '', '', '')

    # escreve uma lista com a folha, titulo e controlador de cada IO
    tag_folha = []
    tag_titulo = []
    tag_controlador = []
    for tag_ios in tags_ios:
        try:
            tag_folha.append(folha)
        except NameError:
            tag_folha.append('')
        try:
            tag_titulo.append(titulo)
        except NameError:
            tag_titulo.append('')
        try:
            tag_controlador.append(controlador)
        except NameError:
            tag_controlador.append('')

    # encontra o sistema de cada IO
    # se há referência ao sistema no título, então os TAGs são deste sistema
    # se não há referência no título, então cada IO tem o sistema referente a seu painel e
    # os que não possuem sistema, herdam o sistema mais frequente da folha
    try:
        titulo = split(titulo, ('-', ' '))
    except NameError:
        titulo = ''

    sistema_titulo = []
    for string in titulo:
        if re.match('\\b(2|5|3)[0-9][0-9][0-9]\\b', string):
            sistema_titulo = string
            break

    sistema_painel = []
    for txt_painel in txt_painel_io:
        txt_painel = split(txt_painel, ('-', ' '))
        for string in txt_painel:
            if re.match('\\b(2|5|3)[0-9][0-9][0-9]\\b', string):
                sistema_painel.append(string)
            elif string == txt_painel[-1]:
                sistema_painel.append('')

    tag_sistema = []
    if len(sistema_titulo) != 0:
        for tag in tags_ios:
            tag_sistema.append(sistema_titulo)
    else:
        freq_sistemas = Counter(sistema_painel)
        sist_mais_frequente = freq_sistemas.most_common()
        for index, tag in enumerate(tags_ios):
            if sistema_painel[index] == '':
                if sist_mais_frequente[0][0] == '':
                    try:
                        tag_sistema.append(sist_mais_frequente[1][0])
                    except IndexError:
                        tag_sistema.append(sist_mais_frequente[0][0])
            else:
                tag_sistema.append(sistema_painel[index])

    # retorna os dados encontrados na folha
    return DadosPagina(tags_ios, tag_sistema, tipo, obs_tag_io, txt_painel_io,
                       tag_controlador, tag_titulo, tag_folha, faixas_sistema)


# =============================================================================
#     Escreve a barra de progresso
# =============================================================================
def imprime_progresso(processo, total_processos):
    bar_length = 20
    freq_atualizacao = int(total_processos/100)
    if freq_atualizacao == 0:
        freq_atualizacao = 1
    if total_processos < 1:
        text = "\r Finalizado"
        sys.stdout.write(text)
        sys.stdout.flush()
    elif processo % freq_atualizacao == 0:
        progresso = float(processo) / float(total_processos)
        block = int(round(bar_length*progresso))
        text = "\rProgresso: [{0}{1}] {2:.0f}%".format("#"*block, "-"*(bar_length-block), progresso*100)
        sys.stdout.write(text)
        sys.stdout.flush()
    else:
        pass
