import openpyxl
import pyperclip
import pyautogui

workbook = openpyxl.load_workbook("produtos_ficticios.xlsx")

sheet_produtos = workbook['Produtos']

for linha in sheet_produtos.iter_rows(min_row=2):
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(136,205, duration=1)
    pyautogui.hotkey('ctrl','v')


    # descricao = linha[1].value
    # categoria = linha[2].value
    # codigo_produto = linha[3].value
    # peso_kg = linha[4].value
    # dimensoes = linha[5].value
    # preco = linha[6].value
    # quantd_estoque = linha[7].value
    # data_validade = linha[8].value
    # cor = linha[9].value
    # tamanho = linha[10].value
    # material = linha[11].value
    # fabricante = linha[12].value
    # pais_origem = linha[13].value
    # observacoes = linha[14].value
    # codigo_barras = linha[15].value
    # localizacao_armazem = linha[16].value

