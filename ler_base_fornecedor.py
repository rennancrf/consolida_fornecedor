import pandas as pd
import os
import datetime as dt

import openpyxl
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Color, PatternFill, Font, Border, Side, Color, Alignment
from openpyxl.formula.translate import Translator
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.formatting.rule import ColorScaleRule, CellIsRule, FormulaRule

import docx
from docx import Document

# import pyautogui
# import PySimpleGUI as sg

# Atribuindo caminho dos diretórios
diretorio = os.getcwd()

path_bases_auxiliares = diretorio + '/bases_auxiliares/'
path_fornecedores = diretorio + '/fornecedores/'
path_catalogo = diretorio + '/catalogo/'
path_template = diretorio + '/templates/'
path_pedidos_flores = diretorio + '/pedidos/'

# Carregando template
wb = load_workbook(filename = path_template + 'Catalogo_flores.xlsx')
ws = wb.active

# Padronizando estilos da planilha
fonte = Font(name='Calibri', size=10, color="000000", bold = False)
borda = Side(border_style="thin", color="000000")
cor_fundo = PatternFill("solid", fgColor="00FFFFFF")
cor_rodape = PatternFill("solid",fgColor="F4B082")

# Inserindo linhas no template
ws[f"B3"].value = dt.datetime.now().strftime('%d-%m-%Y')
       

#Separando tipos de arquivos
arqs_fornecedores_xlsx = pd.DataFrame()
arqs_fornecedores_jpg = pd.DataFrame()

# Carregando bases auxiliares
base_de_cores = pd.DataFrame()
base_de_cores = pd.read_excel(path_bases_auxiliares + 'base_de_cores.xlsx', header=0)

# Lendo pastas de fornecedores
diretorio_fornecedores = os.listdir(path_fornecedores)

# ----- Iniciando processo de importação das bases de fornecedores -----

# Lendo bases de fornecedor 1
arqs_fornecedores = os.listdir(path_fornecedores + 'fornecedor 1/')
 
for arqs in range(len(arqs_fornecedores)):
    if len(arqs_fornecedores) > 0:

        #Separando arquivos .xlsx
        if arqs_fornecedores[arqs].find(".xlsx") != -1:
            
            # Lendo sheets da planilha
            xlsx_sheets_temp = pd.ExcelFile(path_fornecedores + 'fornecedor 1/' + arqs_fornecedores[arqs]).sheet_names
                        
            for sheets in range(len(xlsx_sheets_temp)):
                
                xlsx_temp = pd.DataFrame()
                xlsx_temp = pd.read_excel(path_fornecedores + 'fornecedor 1/' + arqs_fornecedores[arqs], sheet_name = xlsx_sheets_temp[sheets], header=None)
                xlsx_temp.columns = ['Flores','Valor']
                xlsx_temp = xlsx_temp.dropna()
                xlsx_temp = xlsx_temp.join(base_de_cores.set_index('Flores'), how='left', on='Flores')

                if sheets == 0:

                    # Verificando total de linhas da base
                    total_linhas = len(xlsx_temp)
                    print(f"sheet 0 = {total_linhas}")

                    for linha in range(total_linhas):
                    
                        ws[f"A{linha + 6}"].value = xlsx_temp.iat[linha,0]

                        ws[f"B{linha + 6}"].value = xlsx_temp.iat[linha,3]

                        ws[f"C{linha + 6}"].value = xlsx_temp.iat[linha,4]
                        
                        # ws[f"D{linha + 6}"].value = xlsx_sheets_temp[sheets]
                        ws[f"D{linha + 6}"].value = xlsx_temp.iat[linha,2]
                        
                        ws[f"E{linha + 6}"].value = "Fornecedor 1"

                        ws[f"G{linha + 6}"].value = xlsx_temp.iat[linha,1]
                        ws[f"G{linha + 6}"].number_format = 'R$ #,##0.00'

                        ws[f"H{linha + 6}"].value = f"=F{linha + 6}*G{linha + 6}"
                        ws[f"H{linha + 6}"].number_format = 'R$ #,##0.00' 

                        ws[f"A{linha + 6}"].font = fonte
                        ws[f"B{linha + 6}"].font = fonte
                        ws[f"C{linha + 6}"].font = fonte
                        ws[f"D{linha + 6}"].font = fonte
                        ws[f"E{linha + 6}"].font = fonte
                        ws[f"F{linha + 6}"].font = fonte
                        ws[f"G{linha + 6}"].font = fonte
                        ws[f"H{linha + 6}"].font = fonte
                        
                        ws[f"A{linha + 6}"].border = Border(top=borda, left=borda, right=borda, bottom=borda)
                        ws[f"B{linha + 6}"].border = Border(top=borda, left=borda, right=borda, bottom=borda)
                        ws[f"C{linha + 6}"].border = Border(top=borda, left=borda, right=borda, bottom=borda)
                        ws[f"D{linha + 6}"].border = Border(top=borda, left=borda, right=borda, bottom=borda)
                        ws[f"E{linha + 6}"].border = Border(top=borda, left=borda, right=borda, bottom=borda)
                        ws[f"F{linha + 6}"].border = Border(top=borda, left=borda, right=borda, bottom=borda)
                        ws[f"G{linha + 6}"].border = Border(top=borda, left=borda, right=borda, bottom=borda)
                        ws[f"H{linha + 6}"].border = Border(top=borda, left=borda, right=borda, bottom=borda)

                        ws[f"A{linha + 6}"].fill = cor_fundo
                        ws[f"B{linha + 6}"].fill = cor_fundo
                        ws[f"C{linha + 6}"].fill = cor_fundo
                        ws[f"D{linha + 6}"].fill = cor_fundo
                        ws[f"E{linha + 6}"].fill = cor_fundo
                        ws[f"F{linha + 6}"].fill = cor_fundo
                        ws[f"G{linha + 6}"].fill = cor_fundo
                        ws[f"H{linha + 6}"].fill = cor_fundo

                    total_linhas_sheet_anterior = total_linhas
                    print(f"total_linhas_sheet_anterior = {total_linhas_sheet_anterior}")
                
                else:

                    total_linhas = len(xlsx_temp)
                    print(f"sheet {sheets} = {total_linhas}")
                    print(f"total_linhas_sheet_anterior = {total_linhas_sheet_anterior}")

                    for linha in range(total_linhas):
        
                        ws[f"A{linha + total_linhas_sheet_anterior + 6}"].value = xlsx_temp.iat[linha,0]
                        
                        ws[f"B{linha + total_linhas_sheet_anterior + 6}"].value = xlsx_temp.iat[linha,3]

                        ws[f"C{linha + total_linhas_sheet_anterior + 6}"].value = xlsx_temp.iat[linha,4]
                        
                        # ws[f"D{linha + total_linhas_sheet_anterior + 6}"].value = xlsx_sheets_temp[sheets]
                        ws[f"D{linha + total_linhas_sheet_anterior + 6}"].value = xlsx_temp.iat[linha,2]

                        ws[f"E{linha + total_linhas_sheet_anterior + 6}"].value = "Fornecedor 1"
                        
                        ws[f"G{linha + total_linhas_sheet_anterior + 6}"].value = xlsx_temp.iat[linha,1]
                        ws[f"G{linha + total_linhas_sheet_anterior + 6}"].number_format = 'R$ #,##0.00'

                        ws[f"H{linha + total_linhas_sheet_anterior + 6}"].value = f"=F{linha + total_linhas_sheet_anterior + 6}*G{linha + total_linhas_sheet_anterior + 6}"
                        ws[f"H{linha + total_linhas_sheet_anterior + 6}"].number_format = 'R$ #,##0.00' 

                        ws[f"A{linha + total_linhas_sheet_anterior + 6}"].font = fonte
                        ws[f"B{linha + total_linhas_sheet_anterior + 6}"].font = fonte
                        ws[f"C{linha + total_linhas_sheet_anterior + 6}"].font = fonte
                        ws[f"D{linha + total_linhas_sheet_anterior + 6}"].font = fonte
                        ws[f"E{linha + total_linhas_sheet_anterior + 6}"].font = fonte
                        ws[f"F{linha + total_linhas_sheet_anterior + 6}"].font = fonte
                        ws[f"G{linha + total_linhas_sheet_anterior + 6}"].font = fonte
                        ws[f"H{linha + total_linhas_sheet_anterior + 6}"].font = fonte

                        ws[f"A{linha + total_linhas_sheet_anterior + 6}"].border = Border(top=borda, left=borda, right=borda, bottom=borda)
                        ws[f"B{linha + total_linhas_sheet_anterior + 6}"].border = Border(top=borda, left=borda, right=borda, bottom=borda)
                        ws[f"C{linha + total_linhas_sheet_anterior + 6}"].border = Border(top=borda, left=borda, right=borda, bottom=borda)
                        ws[f"D{linha + total_linhas_sheet_anterior + 6}"].border = Border(top=borda, left=borda, right=borda, bottom=borda)
                        ws[f"E{linha + total_linhas_sheet_anterior + 6}"].border = Border(top=borda, left=borda, right=borda, bottom=borda)
                        ws[f"F{linha + total_linhas_sheet_anterior + 6}"].border = Border(top=borda, left=borda, right=borda, bottom=borda)
                        ws[f"G{linha + total_linhas_sheet_anterior + 6}"].border = Border(top=borda, left=borda, right=borda, bottom=borda)
                        ws[f"H{linha + total_linhas_sheet_anterior + 6}"].border = Border(top=borda, left=borda, right=borda, bottom=borda)

                        ws[f"A{linha + total_linhas_sheet_anterior + 6}"].fill = cor_fundo
                        ws[f"B{linha + total_linhas_sheet_anterior + 6}"].fill = cor_fundo
                        ws[f"C{linha + total_linhas_sheet_anterior + 6}"].fill = cor_fundo
                        ws[f"D{linha + total_linhas_sheet_anterior + 6}"].fill = cor_fundo
                        ws[f"E{linha + total_linhas_sheet_anterior + 6}"].fill = cor_fundo
                        ws[f"F{linha + total_linhas_sheet_anterior + 6}"].fill = cor_fundo
                        ws[f"G{linha + total_linhas_sheet_anterior + 6}"].fill = cor_fundo
                        ws[f"H{linha + total_linhas_sheet_anterior + 6}"].fill = cor_fundo

                    total_linhas_sheet_anterior = total_linhas_sheet_anterior + total_linhas

# Lendo bases de fornecedor 2
arqs_fornecedores = os.listdir(path_fornecedores + 'fornecedor 2/')
       
for arqs in range(len(arqs_fornecedores)):
    if len(arqs_fornecedores) > 0:

        #Separando arquivos .xlsx
        if arqs_fornecedores[arqs].find(".xlsx") != -1:
            # Lendo sheets da planilha
            xlsx_sheets_temp = pd.ExcelFile(path_fornecedores + 'fornecedor 2/' + arqs_fornecedores[arqs]).sheet_names
                        
            for sheets in range(len(xlsx_sheets_temp)):
                
                xlsx_temp = pd.DataFrame()
                xlsx_temp = pd.read_excel(path_fornecedores + 'fornecedor 2/' + arqs_fornecedores[arqs], sheet_name = xlsx_sheets_temp[sheets], header=None, skiprows=3)
                xlsx_temp.columns = ['Flores','Valor']
                xlsx_temp = xlsx_temp.dropna()
                xlsx_temp = xlsx_temp.join(base_de_cores.set_index('Flores'), how='left', on='Flores')
                
                # Verificando total de linhas da base
                total_linhas = len(xlsx_temp)
                print(f"sheet {sheets} = {total_linhas}")
                print(f"total_linhas_sheet_anterior = {total_linhas_sheet_anterior}")
                
                for linha_fornecedor_2 in range(total_linhas):
                    
                    ws[f"A{linha_fornecedor_2 + total_linhas_sheet_anterior + 6}"].value = xlsx_temp.iat[linha_fornecedor_2,0]
                    
                    ws[f"B{linha_fornecedor_2 + total_linhas_sheet_anterior + 6}"].value = xlsx_temp.iat[linha_fornecedor_2,3]

                    ws[f"C{linha_fornecedor_2 + total_linhas_sheet_anterior + 6}"].value = xlsx_temp.iat[linha_fornecedor_2,4]
                    
                    # ws[f"D{linha_fornecedor_2 + total_linhas_sheet_anterior + 6}"].value = "N/A"
                    ws[f"D{linha_fornecedor_2 + total_linhas_sheet_anterior + 6}"].value = xlsx_temp.iat[linha_fornecedor_2,2]

                    ws[f"E{linha_fornecedor_2 + total_linhas_sheet_anterior + 6}"].value = "Fornecedor 2"
                    
                    ws[f"G{linha_fornecedor_2 + total_linhas_sheet_anterior + 6}"].value = xlsx_temp.iat[linha_fornecedor_2,1]
                    ws[f"G{linha_fornecedor_2 + total_linhas_sheet_anterior + 6}"].number_format = 'R$ #,##0.00'

                    ws[f"H{linha_fornecedor_2 + total_linhas_sheet_anterior + 6}"].value = f"=F{linha_fornecedor_2 + total_linhas_sheet_anterior + 6}*G{linha_fornecedor_2 + total_linhas_sheet_anterior + 6}"
                    ws[f"H{linha_fornecedor_2 + total_linhas_sheet_anterior + 6}"].number_format = 'R$ #,##0.00' 

                    ws[f"A{linha_fornecedor_2 + total_linhas_sheet_anterior + 6}"].font = fonte
                    ws[f"B{linha_fornecedor_2 + total_linhas_sheet_anterior + 6}"].font = fonte
                    ws[f"C{linha_fornecedor_2 + total_linhas_sheet_anterior + 6}"].font = fonte
                    ws[f"D{linha_fornecedor_2 + total_linhas_sheet_anterior + 6}"].font = fonte
                    ws[f"E{linha_fornecedor_2 + total_linhas_sheet_anterior + 6}"].font = fonte
                    ws[f"F{linha_fornecedor_2 + total_linhas_sheet_anterior + 6}"].font = fonte
                    ws[f"G{linha_fornecedor_2 + total_linhas_sheet_anterior + 6}"].font = fonte
                    ws[f"H{linha_fornecedor_2 + total_linhas_sheet_anterior + 6}"].font = fonte

                    ws[f"A{linha_fornecedor_2 + total_linhas_sheet_anterior + 6}"].border = Border(top=borda, left=borda, right=borda, bottom=borda)
                    ws[f"B{linha_fornecedor_2 + total_linhas_sheet_anterior + 6}"].border = Border(top=borda, left=borda, right=borda, bottom=borda)
                    ws[f"C{linha_fornecedor_2 + total_linhas_sheet_anterior + 6}"].border = Border(top=borda, left=borda, right=borda, bottom=borda)
                    ws[f"D{linha_fornecedor_2 + total_linhas_sheet_anterior + 6}"].border = Border(top=borda, left=borda, right=borda, bottom=borda)
                    ws[f"E{linha_fornecedor_2 + total_linhas_sheet_anterior + 6}"].border = Border(top=borda, left=borda, right=borda, bottom=borda)
                    ws[f"F{linha_fornecedor_2 + total_linhas_sheet_anterior + 6}"].border = Border(top=borda, left=borda, right=borda, bottom=borda)
                    ws[f"G{linha_fornecedor_2 + total_linhas_sheet_anterior + 6}"].border = Border(top=borda, left=borda, right=borda, bottom=borda)
                    ws[f"H{linha_fornecedor_2 + total_linhas_sheet_anterior + 6}"].border = Border(top=borda, left=borda, right=borda, bottom=borda)

                    ws[f"A{linha_fornecedor_2 + total_linhas_sheet_anterior + 6}"].fill = cor_fundo
                    ws[f"B{linha_fornecedor_2 + total_linhas_sheet_anterior + 6}"].fill = cor_fundo
                    ws[f"C{linha_fornecedor_2 + total_linhas_sheet_anterior + 6}"].fill = cor_fundo
                    ws[f"D{linha_fornecedor_2 + total_linhas_sheet_anterior + 6}"].fill = cor_fundo
                    ws[f"E{linha_fornecedor_2 + total_linhas_sheet_anterior + 6}"].fill = cor_fundo
                    ws[f"F{linha_fornecedor_2 + total_linhas_sheet_anterior + 6}"].fill = cor_fundo
                    ws[f"G{linha_fornecedor_2 + total_linhas_sheet_anterior + 6}"].fill = cor_fundo
                    ws[f"H{linha_fornecedor_2 + total_linhas_sheet_anterior + 6}"].fill = cor_fundo

                total_linhas_sheet_anterior = total_linhas_sheet_anterior + total_linhas
                    
# Editando cor das células do rodapé da tabela
ws[f"A{total_linhas_sheet_anterior + 6}"].fill = cor_rodape
ws[f"B{total_linhas_sheet_anterior + 6}"].fill = cor_rodape
ws[f"C{total_linhas_sheet_anterior + 6}"].fill = cor_rodape
ws[f"D{total_linhas_sheet_anterior + 6}"].fill = cor_rodape
ws[f"E{total_linhas_sheet_anterior + 6}"].fill = cor_rodape
ws[f"F{total_linhas_sheet_anterior + 6}"].fill = cor_rodape
ws[f"G{total_linhas_sheet_anterior + 6}"].fill = cor_rodape
ws[f"H{total_linhas_sheet_anterior + 6}"].fill = cor_rodape

# Inserindo bordas nas células de rodapé da tabela
ws[f"A{total_linhas_sheet_anterior + 6}"].border = Border(top=borda, left=borda, right=borda, bottom=borda)
ws[f"B{total_linhas_sheet_anterior + 6}"].border = Border(top=borda, left=borda, right=borda, bottom=borda)
ws[f"C{total_linhas_sheet_anterior + 6}"].border = Border(top=borda, left=borda, right=borda, bottom=borda)
ws[f"D{total_linhas_sheet_anterior + 6}"].border = Border(top=borda, left=borda, right=borda, bottom=borda)
ws[f"E{total_linhas_sheet_anterior + 6}"].border = Border(top=borda, left=borda, right=borda, bottom=borda)
ws[f"F{total_linhas_sheet_anterior + 6}"].border = Border(top=borda, left=borda, right=borda, bottom=borda)
ws[f"G{total_linhas_sheet_anterior + 6}"].border = Border(top=borda, left=borda, right=borda, bottom=borda)
ws[f"H{total_linhas_sheet_anterior + 6}"].border = Border(top=borda, left=borda, right=borda, bottom=borda)



ws.sheet_view.showGridLines = False
wb.save(path_catalogo + 'Catalogo Flores - ' + dt.datetime.now().strftime('%d-%m-%Y') + '.xlsx')
wb.close()

# ----- Transferindo versões de catálogos antigos para o backlog -----

# Criando diretorio de backlog caso não exista
if not (os.path.isdir(path_catalogo + '_old')):
    os.mkdir(path_catalogo + '_old')

arquivos_catalogos_old = os.listdir(path_catalogo)

for arqs_old in range(len(arquivos_catalogos_old)):
    
    if (arquivos_catalogos_old[arqs_old] != 'Catalogo Flores - ' + dt.datetime.now().strftime('%d-%m-%Y') + '.xlsx') and (arquivos_catalogos_old[arqs_old].find('.xlsx') != -1):
        os.replace(path_catalogo + arquivos_catalogos_old[arqs_old], path_catalogo + '_old/' + arquivos_catalogos_old[arqs_old])



print('Catalogo Atualizado!')

