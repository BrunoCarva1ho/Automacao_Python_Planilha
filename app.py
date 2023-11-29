import openpyxl
import pyperclip
import pyautogui
from time import sleep

workbook = openpyxl.load_workbook("produtos_ficticios.xlsx")

sheet_produtos = workbook['Produtos']

#Abra a página de cadastro de produtos após iniciar o código
#5 segundos até começar, clique na página e aguarde
sleep(5)


for linha in sheet_produtos.iter_rows(min_row=2):
    
    
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.hotkey('tab')
    pyautogui.hotkey('ctrl','v')

    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.hotkey('tab')
    pyautogui.hotkey('ctrl','v')

    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.hotkey('tab')
    pyautogui.hotkey('ctrl','v')

    codigo_produto = linha[3].value
    pyperclip.copy(codigo_produto)
    pyautogui.hotkey('tab')
    pyautogui.hotkey('ctrl','v')

    peso_kg = linha[4].value
    pyperclip.copy(peso_kg)
    pyautogui.hotkey('tab')
    pyautogui.hotkey('ctrl','v')

    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.hotkey('tab')
    pyautogui.hotkey('ctrl','v')
    

    #PRESSIONAR BOTÃO PARA PRÓXIMA PÁGINA
    pyautogui.hotkey('tab','Enter')
    sleep(2)
    
    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.hotkey('tab')
    pyautogui.hotkey('ctrl','v')

    quantd_estoque = linha[7].value
    pyperclip.copy(quantd_estoque)
    pyautogui.hotkey('tab')
    pyautogui.hotkey('ctrl','v')

    data_validade = linha[8].value
    pyperclip.copy(data_validade)
    pyautogui.hotkey('tab')
    pyautogui.hotkey('ctrl','v')

    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.hotkey('tab')
    pyautogui.hotkey('ctrl','v')

    tamanho = linha[10].value
    pyperclip.copy(tamanho[0])
    pyautogui.hotkey('tab','Enter')
    pyautogui.hotkey(tamanho[0])
    sleep(1)
    pyautogui.hotkey('Enter')
    
    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.hotkey('tab')
    pyautogui.hotkey('ctrl','v')
    

    #PRESSIONAR BOTÃO PARA PRÓXIMA PÁGINA
    pyautogui.hotkey('tab','Enter')
    sleep(2)

    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.hotkey('tab')
    pyautogui.hotkey('ctrl','v')
    
    pais_origem = linha[13].value
    pyperclip.copy(pais_origem)
    pyautogui.hotkey('tab')
    pyautogui.hotkey('ctrl','v')

    observacoes = linha[14].value
    pyperclip.copy(observacoes)
    pyautogui.hotkey('tab')
    pyautogui.hotkey('ctrl','v')

    codigo_barras = linha[15].value
    pyperclip.copy(codigo_barras)
    pyautogui.hotkey('tab')
    pyautogui.hotkey('ctrl','v')

    localizacao_armazem = linha[16].value
    pyperclip.copy(localizacao_armazem)
    pyautogui.hotkey('tab')
    pyautogui.hotkey('ctrl','v')


    #PRESSIONAR BOTÃO PARA CONCLUIR
    pyautogui.hotkey('tab','Enter')
