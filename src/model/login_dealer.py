import pyautogui as p
import os


def login_dealer_oficina(): 
    # Abrindo o módulo oficina do Dealer e realizando o login
    credenciais = open('C:\\RPA\\credenciais\\login_dealer.txt','r')
    chaves = credenciais.readlines()
    credenciais.close()

    login = chaves[0][:-1]
    password = chaves[1]

    p.PAUSE=1
    os.startfile("C:\Conces\sof\sof.exe")
    p.sleep(3)

    # Inserindo as credenciais de usuário e senha do Dealer
    usuario = p.locateCenterOnScreen('C:/RPA/arquivos/images/usuario.png',confidence = 0.95)
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

    # # Selecionando a empresa para realizar o login
    # p.click(860,380) #Coordenadas da seta do menu dropdown com os nomes das empresas
    # p.sleep(0.5)
    # p.scroll(2000)
    # p.sleep(0.5)
    # empresa = p.locateCenterOnScreen(f'C:/RPA/arquivos/images/muda_empresa/{name}.png', confidence = 0.95)
    # while empresa == None:
    #     p.click(860,638)
    #     empresa = p.locateCenterOnScreen(f'C:/RPA/arquivos/images/muda_empresa/{name}.png', confidence = 0.95)
    # p.click(empresa.x,empresa.y)

    # Selecionando a empresa para realizar o login (JANGADA VEICULOS)
    p.typewrite('jangada veiculos')
    p.sleep(0.5)
    p.press('tab', presses=3)
    p.sleep(0.5)
    p.press('space')

    # p.sleep(0.5) 
    # p.press('tab')
    # p.press('enter')

    # Loop para aguardar o carregamento do banco de dados do Dealer
    cancelar = p.locateCenterOnScreen('C:/RPA/arquivos/images/cancelar.png', confidence = 0.95)
    while cancelar != None:
        p.sleep(0.5)
        cancelar = p.locateCenterOnScreen('C:/RPA/arquivos/images/cancelar.png', confidence = 0.95)
    p.sleep(1)

