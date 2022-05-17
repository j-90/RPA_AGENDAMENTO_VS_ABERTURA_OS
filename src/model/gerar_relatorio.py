import pyautogui as p
import os
from datetime import date


def gera_relatorio():
    # Acessando o menu para gerar o relatório
    p.hotkey('alt', 'r')
    p.sleep(0.5)
    p.press('down', presses=6)
    p.sleep(0.5)
    p.press('enter')
    p.sleep(0.5)

    # Loop para aguardar o botão 'Preview' aparecer na tela
    preview = p.locateCenterOnScreen(r'caminho_do_arquivo\images\preview.png', confidence=0.95)
    while preview == None:
        p.sleep(0.5)
        preview = p.locateCenterOnScreen(r'caminho_do_arquivo\images\preview.png', confidence=0.95)
    p.sleep(1)

    # Calculando a data inicial e final da semana corrente, inserindo nos respectivos campos e clicando no botão 'Preview'
    data_atual = date.today()

    data_em_texto_atual = data_atual.strftime('%d/%m/%Y')
    dia = data_em_texto_atual[:2]
    mes = data_em_texto_atual[3:5]
    ano = data_em_texto_atual[8:]

    setData = date.fromordinal(data_atual.toordinal() - 4)
    setData_barra = setData.strftime('%d/%m/%Y')
    dia_setData = setData_barra[:2]
    mes_setData = setData_barra[3:5]
    ano_setData = setData_barra[8:]

    p.typewrite(dia_setData + mes_setData + ano_setData)
    p.sleep(0.5)
    p.press('tab')
    p.sleep(0.5)
    p.typewrite(dia + mes + ano)
    p.sleep(1)

    preview = p.locateCenterOnScreen(r'caminho_do_arquivo\images\preview.png', confidence=0.95)
    if preview != None:
        p.click(preview.x, preview.y)
    p.sleep(1)

    # Loop para aguardar o relatório aparecer na tela
    agendamento_servicos = p.locateCenterOnScreen(r'caminho_do_arquivo\images\agendamento_servicos_analitico_todos.png', confidence=0.95)
    while agendamento_servicos == None:
        p.sleep(1)
        agendamento_servicos = p.locateCenterOnScreen(r'caminho_do_arquivo\images\agendamento_servicos_analitico_todos.png', confidence=0.95)
    p.sleep(0.5)

    p.hotkey('alt', 'i')
    p.sleep(0.5)
    p.press('right')
    p.sleep(0.5)
    p.press('tab', presses=2)
    p.sleep(0.5)
    p.press('enter')
    p.sleep(0.5)
    p.typewrite('caminho_do_arquivo\\nome_do_arquivo')
    p.sleep(0.5)
    p.press('enter')

