import pyautogui as p
import os
from src.model.login import login_sistema
from src.model.gerar_relatorio import gera_relatorio
from src.model.gerar_nova_planilha import gera_nova_planilha
from src.model.enviar_email import envia_email


try:
    print('Fazendo login no sistema...\n')
    login_sistema()
    print('Login efetuado!\n')

    print('Gerando relatorio...\n')
    gera_relatorio()
    print('Planilha gerada com sucesso!\n')

    gera_nova_planilha()

    envia_email()

finally:
    p.sleep(2)
    os.system("taskkill /im executavel.exe")

