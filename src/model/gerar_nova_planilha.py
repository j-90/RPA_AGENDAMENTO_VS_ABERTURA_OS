import pandas as pd
import xlrd
import xlwt
import pyautogui as p
import os
import logging
import matplotlib
import datetime
import math
from datetime import datetime


def gera_nova_planilha():
    # Transformando o arquivo .xls em .xlsx
    print('Gerando o arquivo XLSX a partir do arquivo XLS...\n')
    wb = xlrd.open_workbook('C:\\RPA\\arquivos\\Agendamento_de_Servicos_Jangada_Renault.XLS', encoding_override='latin1')
    df = pd.read_excel(wb, dtype=str)
    df.to_excel('C:\\RPA\\arquivos\\Agendamento_de_Servicos_Jangada_Renault.xlsx', index = False)
    print('Arquivo XLSX gerado!\n')
    p.sleep(1)

    # Removendo as colunas desnecessárias do arquivo .xlsx
    print('Removendo as colunas desnecessarias...\n')
    df1 = pd.read_excel('C:\\RPA\\arquivos\\Agendamento_de_Servicos_Jangada_Renault.xlsx', dtype=str)
    # df1 = pd.DataFrame(df1)
    df1 = df1.drop(["ag_hr", "ag_dtcad", "obs_ag", "tm_cd", "fun_cd", "fun_nmguerra",
                    "ag_telefone1", "ag_telefone2", "fun_cad", "fun_cad.1", "sof_os_os_hraber", "name_1", "cont_ag",
                    "cont_os", "name_2", "sof_ag_ag_hrcad", "sof_ag_ag_telefone3", "sof_ag_ag_dtprev", "sof_ag_ag_hrprev",
                    "sof_ag_ag_preordem", "agtm_ccc", "pm_fabricante", "ag_tp", "ag_status", "count_noshow", "count_reagendado",
                    "count_cancelado"], axis=1)
    print('Colunas removidas!\n')
    p.sleep(1)

    # # Deletando as linhas que estão vazias na coluna 'os_dtaber'
    print('Removendo as linhas vazias...\n')
    df1.dropna(axis=0, subset=['os_nr'], inplace=True)
    print('Linhas vazias removidas!\n')
    p.sleep(1)

    # Removendo a hora zerada contida nas colunas 'ag_dt' e 'os_dtaber' e formatando a data
    print('Selecionando apenas a data dos campos de data de agendamento e data de abertura da OS...\n')

    lista_agendamento = [str(x)[:10] for x in df1['ag_dt']]
    lista_abertura_os = [str(x)[:10] for x in df1['os_dtaber']]

    print('Salvando apenas as datas nas respectivas colunas...\n')

    df1['ag_dt'] = lista_agendamento
    df1['os_dtaber'] = lista_abertura_os
    for data_agendamento in lista_agendamento:
        data_agendamento.replace('%y-%m-%d', '%d/%m/%y')
    for data_abertura_os in lista_abertura_os:
        data_abertura_os.replace('%y-%m-%d', '%d/%m/%y')

    print('Colunas de data de agendamento e data de abertura da OS atualizadas!\n')
    p.sleep(1) 

    # Criando uma nova coluna para armazenar o status de diferença entre a data de agendamento e a data de abertura
    print('Comparando as datas de agendamento e de abertura da OS...\n')
    lista_pontualidade = []
    lista_auxiliar = zip(lista_agendamento, lista_abertura_os)
    for elemento in lista_auxiliar:
        if datetime.strptime(elemento[0], '%d/%m/%Y').date().toordinal()  >= datetime.strptime(elemento[1], '%d/%m/%Y').date().toordinal():
            lista_pontualidade.append('Ok')
        else:
            lista_pontualidade.append('Retardo')
    p.sleep(1)
    
    print('Criando a coluna Pontualidade para armazenar os resultados das comparacoes...\n')
    pontualidade = lista_pontualidade
    df1['Pontualidade'] = pontualidade
    print('Coluna criada e resultados atribuidos!\n')
    p.sleep(1)
    
    # # Salvando um novo arquivo XLSX
    print('Criando um novo arquivo XLSX...\n')
    nova_planilha = 'C:\\RPA\\arquivos\\Agendamento_de_Servicos_Jangada_Renault.xlsx'
    df1.to_excel(nova_planilha, index = False)
    print('Novo arquivo XLSX criado!\n')
    p.sleep(1)               
    
# def main():
#     gera_nova_planilha()

# if __name__ == '__main__':
#     main()

