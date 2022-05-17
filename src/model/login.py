import pyautogui as p
import os


def login_sistema(): 
    # Abrindo o módulo oficina do Dealer e realizando o login
    credenciais = open('caminho_do_arquivo\\login.txt','r')
    chaves = credenciais.readlines()
    credenciais.close()

    login = chaves[0][:-1]
    password = chaves[1]

    p.PAUSE=1
    os.startfile("caminho_do_arquivo\executavel.exe")
    p.sleep(3)

    # Inserindo as credenciais de usuário e senha do Dealer
    usuario = p.locateCenterOnScreen('caminho_do_arquivo/images/usuario.png',confidence = 0.95)
    if usuario != None:
        p.click(usuario.x+203,usuario.y)
        p.press('backspace', presses=10)
        p.sleep(1)
        p.write(login)
    p.sleep(0.5)
    p.press('tab')
    p.sleep(0.5)
    p.write(password)
    p.sleep(2)
    p.press('tab', presses=2)
    p.sleep(0.5)

    # Selecionando a empresa para realizar o login (NOME DA EMPRESA)
    p.typewrite('nome da empresa')
    p.sleep(0.5)
    p.press('tab', presses=3)
    p.sleep(0.5)
    p.press('space')

    # p.sleep(0.5) 
    # p.press('tab')
    # p.press('enter')

    # Loop para aguardar o carregamento do banco de dados do Dealer
    cancelar = p.locateCenterOnScreen('caminho_do_arquivo/images/cancelar.png', confidence = 0.95)
    while cancelar != None:
        p.sleep(0.5)
        cancelar = p.locateCenterOnScreen('caminho_do_arquivo/images/cancelar.png', confidence = 0.95)
    p.sleep(1)

