import openpyxl
import pyperclip
import pyautogui
from time import sleep
# entrar na planilha

workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')

sheet_produtos = workbook['Produtos']

for linha in sheet_produtos.iter_rows(min_row=2):
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(283, 320, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.click(108, 423, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(159, 572, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    codigo_produto = linha[3].value
    pyperclip.copy(codigo_produto)
    pyautogui.click(138, 672, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(154, 782, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.click(119, 874, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    # clicando em próximo
    pyautogui.click(99, 938, duration=1)
    sleep(3)

    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(141, 333, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    quantidade_em_estoque = linha[7].value
    pyperclip.copy(quantidade_em_estoque)
    pyautogui.click(140, 432, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    data_de_validade = linha[8].value
    pyperclip.copy(data_de_validade)
    pyautogui.click(139, 534, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(139, 640, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    # ler info da planilha
    # se 'pequeno', clicar em uma posição
    # se 'médio', clicar em uma posição
    # se 'grande', clicar em uma posição
    tamanho = linha[10].value
    pyautogui.click(173, 742, duration=1)
    if tamanho == 'Pequeno':
        pyautogui.click(137, 775, duration=1)
    elif tamanho == 'Médio':
        pyautogui.click(123, 800, duration=1)
    else:
        pyautogui.click(110, 828, duration=1)

    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.click(125, 849, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    # clicando em próximo novamente
    pyautogui.click(101, 911, duration=1)
    sleep(3)

    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(112, 371, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    pais_origem = linha[13].value
    pyperclip.copy(pais_origem)
    pyautogui.click(112, 481, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    observacoes = linha[14].value
    pyperclip.copy(observacoes)
    pyautogui.click(115, 578, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    codigo_de_barras = linha[15].value
    pyperclip.copy(codigo_de_barras)
    pyautogui.click(116, 746, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    localizacao_fabricante = linha[16].value
    pyperclip.copy(localizacao_fabricante)
    pyautogui.click(140, 840, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    # clicando em Concluir
    pyautogui.click(100, 904, duration=1)
    sleep(3)

    # clicando OK no pop-up
    pyautogui.click(736, 264, duration=1)
    sleep(3)

    # clicando "Adicionar Mais Um" no pop-up
    pyautogui.click(467, 620, duration=1)
    sleep(3)
