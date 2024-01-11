# NÃO ADIANTA PEGAR O CÓDIGO E QUERER RODAR SEM MUDAR AS COORDENAS!
# CADA JANELA TEM POSIÇÕES(COORDENADAS DIFERENTES)
# SENDO ASSIM, ASSISTA O VIDEO INTEIRO PARA ENTENDER COMO MONTAR ESTE PROGRAMA

import openpyxl
import pyperclip
import pyautogui
from time import sleep

# Entrar na planilha
workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
sheet_produtos = workbook['Produtos']

# Copiar informação de um campo e colar no seu campo correspondente
for linha in sheet_produtos.iter_rows(min_row=2):
    # Nome do Produto
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(845,266, duration=1)
    pyautogui.hotkey('ctrl', 'v')

   # Descrição do Produto
    descricao_produto = linha[1].value
    pyperclip.copy(descricao_produto)
    pyautogui.click(840,323, duration=1)
    pyautogui.hotkey ('ctrl','v')

    #Categoria
    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(841,412, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    # codigo_produto
    codigo_produto = linha[3].value
    pyperclip.copy(codigo_produto)
    pyautogui.click(841,471, duration=1)
    pyautogui.hotkey('ctrl', 'v')


    # peso
    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(839,534, duration=1)
    pyautogui.hotkey('ctrl', 'v')


    # dimensao
    dimensao = linha[5].value
    pyperclip.copy(dimensao)
    pyautogui.click(837,595, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    #botao proximo
    pyautogui.click(851,631, duration=1)
    sleep(3)

    # preco
    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(846,284, duration=1)
    pyautogui.hotkey('ctrl','v')

    # quantidade_estoque
    quantidade_estoque = linha[7].value
    pyperclip.copy(quantidade_estoque)
    pyautogui.click(844,349, duration=1)
    pyautogui.hotkey('ctrl','v')

    # dtValidade
    dtValidade = linha[8].value
    pyperclip.copy(dtValidade)
    pyautogui.click(843,397, duration=1)
    pyautogui.hotkey('ctrl','v')

    # cor
    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(838,458, duration=1)
    pyautogui.hotkey('ctrl','v')
    
    # tamanho
    tamanho = linha[10].value
    pyautogui.click(863,516, duration=1)
    if tamanho == 'Pequeno':
        pyautogui.click(851,540, duration=1)
    elif tamanho == 'Médio':
        pyautogui.click(856,552, duration=1)
    else:
        pyautogui.click(864,573, duration=1)

    # material
    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.click(837,572, duration=1)
    pyautogui.hotkey('ctrl','v')

    pyautogui.click(848,612, duration=1)
    sleep(3)


    # fabricante
    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(847,307, duration=1)
    pyautogui.hotkey('ctrl','v')


    # pais_origem
    pais_origem = linha[13].value
    pyperclip.copy(pais_origem)
    pyautogui.click(840,365, duration=1)
    pyautogui.hotkey('ctrl','v')
    

    # observacao
    observacao = linha[14].value
    pyperclip.copy(observacao)
    pyautogui.click(846,423, duration=1)
    pyautogui.hotkey('ctrl','v')


    # cBarras
    cBarras = linha[15].value
    pyperclip.copy(cBarras)
    pyautogui.click(840,519, duration=1)
    pyautogui.hotkey('ctrl','v')

    # localizacao_armazem
    localizacao_armazem = linha[16].value
    pyperclip.copy(localizacao_armazem)
    pyautogui.click(846,574, duration=1)
    pyautogui.hotkey('ctrl','v')

    pyautogui.click(846,610, duration=1)
   

    pyautogui.click(1230,235, duration=1)

    pyautogui.click(1056,459, duration=1)
    
   
    # Nome do Produto
    # Descrição
    # Categoria
    # Codigo Produto
    # Peso
    # Dimensões
    # Botão próximo
    # Preço
    # Quantidade em estoque
    # Data de validade
    # Cor
    # Tamanho
    # material   
    # Botão próximo
    # Botão concluir
    # Botão confirmar inclusão
    # Botão confirmação 2
    # iniciar cadastro novamente
# Repitir esses passos para outros campos até preencher campos daquela página
# Clicar em próxima
# Repetir os mesmos passos e ir para a próxima página(página 2)
# Repetir os mesmos passos e finalizar o cadastro daquele produto e clicar em concluir
# clicar em ok, para finalizar o processo
# Clicar no ok mais um vez na mensagem de confirmação de salvamento no banco de dados
# Clicar em "adicionar mais um e repetir o processo até finalizar a planilha"
