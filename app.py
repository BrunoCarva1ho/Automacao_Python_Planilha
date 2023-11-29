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

    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.click(118,279, duration=1)
    pyautogui.hotkey('ctrl','v')

    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(115,404, duration=1)
    pyautogui.hotkey('ctrl','v')

    codigo_produto = linha[3].value
    pyperclip.copy(codigo_produto)
    pyautogui.click(120,477, duration=1)
    pyautogui.hotkey('ctrl','v')

    peso_kg = linha[4].value
    pyperclip.copy(peso_kg)
    pyautogui.click(117,558, duration=1)
    pyautogui.hotkey('ctrl','v')

    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.click(115,636, duration=1)
    pyautogui.hotkey('ctrl','v')
    

    #PRESSIONAR BOTÃO PARA PRÓXIMA PÁGINA
    pyautogui.click(137,679, duration=1)
    
    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(121,231, duration=1)
    pyautogui.hotkey('ctrl','v')

    quantd_estoque = linha[7].value
    pyperclip.copy(quantd_estoque)
    pyautogui.click(127,306, duration=1)
    pyautogui.hotkey('ctrl','v')

    data_validade = linha[8].value
    pyperclip.copy(data_validade)
    pyautogui.click(124,390, duration=1)
    pyautogui.hotkey('ctrl','v')

    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(119,462, duration=1)
    pyautogui.hotkey('ctrl','v')

    tamanho = linha[10].value
    pyperclip.copy(tamanho[0])
    pyautogui.click(124,540, duration=1)
    pyautogui.hotkey(tamanho[0])
    pyautogui.hotkey('Enter')
    
    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.click(127,617, duration=1)
    pyautogui.hotkey('ctrl','v')
    

    #PRESSIONAR BOTÃO PARA PRÓXIMA PÁGINA
    pyautogui.click(140,674, duration=1)

    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(121,248, duration=1)
    pyautogui.hotkey('ctrl','v')
    
    pais_origem = linha[13].value
    pyperclip.copy(pais_origem)
    pyautogui.click(118,322, duration=1)
    pyautogui.hotkey('ctrl','v')

    observacoes = linha[14].value
    pyperclip.copy(observacoes)
    pyautogui.click(128,411, duration=1)
    pyautogui.hotkey('ctrl','v')

    codigo_barras = linha[15].value
    pyperclip.copy(codigo_barras)
    pyautogui.click(120,523, duration=1)
    pyautogui.hotkey('ctrl','v')

    localizacao_armazem = linha[16].value
    pyperclip.copy(localizacao_armazem)
    pyautogui.click(120,600, duration=1)
    pyautogui.hotkey('ctrl','v')


    #PRESSIONAR BOTÃO PARA CONCLUIR
    pyautogui.click(143,656, duration=1)
