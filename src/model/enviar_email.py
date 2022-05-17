# -*- coding: UTF-8 -*-
import sys
import os
import pyautogui as p
from datetime import date
# Importando os pacotes para automaçao de email
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email import encoders
# Imports do Código Oauth 2.0
import base64
import imaplib
import json
import urllib.parse
import urllib.request
import lxml.html
import pandas as pd
import pyautogui as p
import logging
import time

# Links Úteis:
# Exemplo de Uso do Oauth para envio de email com Gmail -> https://blog.macuyiko.com/post/2016/how-to-send-html-mails-with-oauth2-and-gmail-in-python.html
# Documentação do Protocolo Oauth 2.0 -> : https://developers.google.com/gmail/imap/xoauth2-protocol
# Repositório do GitHub do Google com uso do Oauth 2.0 em Python -> https://github.com/google/gmail-oauth2-tools/blob/master/python/oauth2.py

credenciais = open('caminho_do_arquivo\\credenciais_gmail_oauth.txt', 'r')
chaves = credenciais.readlines()
credenciais.close()

GOOGLE_ACCOUNTS_BASE_URL = 'https://accounts.google.com'
REDIRECT_URI = 'urn:ietf:wg:oauth:2.0:oob'

GOOGLE_CLIENT_ID = chaves[1][:-1]
GOOGLE_CLIENT_SECRET = chaves[3][:-1]
GOOGLE_REFRESH_TOKEN = chaves[5]
# GOOGLE_REFRESH_TOKEN = None


def command_to_url(command):
    return '%s/%s' % (GOOGLE_ACCOUNTS_BASE_URL, command)


def url_escape(text):
    return urllib.parse.quote(text, safe='~-._')


def url_unescape(text):
    return urllib.parse.unquote(text)


def url_format_params(params):
    param_fragments = []
    for param in sorted(params.items(), key=lambda x: x[0]):
        param_fragments.append('%s=%s' % (param[0], url_escape(param[1])))
    return '&'.join(param_fragments)


def generate_permission_url(client_id, scope='https://mail.google.com/'):
    params = {}
    params['client_id'] = client_id
    params['redirect_uri'] = REDIRECT_URI
    params['scope'] = scope
    params['response_type'] = 'code'
    return '%s?%s' % (command_to_url('o/oauth2/auth'), url_format_params(params))


def call_authorize_tokens(client_id, client_secret, authorization_code):
    params = {}
    params['client_id'] = client_id
    params['client_secret'] = client_secret
    params['code'] = authorization_code
    params['redirect_uri'] = REDIRECT_URI
    params['grant_type'] = 'authorization_code'
    request_url = command_to_url('o/oauth2/token')
    response = urllib.request.urlopen(request_url, urllib.parse.urlencode(
        params).encode('UTF-8')).read().decode('UTF-8')
    return json.loads(response)


def call_refresh_token(client_id, client_secret, refresh_token):
    params = {}
    params['client_id'] = client_id
    params['client_secret'] = client_secret
    params['refresh_token'] = refresh_token
    params['grant_type'] = 'refresh_token'
    request_url = command_to_url('o/oauth2/token')
    response = urllib.request.urlopen(request_url, urllib.parse.urlencode(
        params).encode('UTF-8')).read().decode('UTF-8')
    return json.loads(response)


def generate_oauth2_string(username, access_token, as_base64=False):
    auth_string = 'user=%s\1auth=Bearer %s\1\1' % (username, access_token)
    if as_base64:
        auth_string = base64.b64encode(
            auth_string.encode('ascii')).decode('ascii')
    return auth_string


def test_imap(user, auth_string):
    imap_conn = imaplib.IMAP4_SSL('imap.gmail.com')
    imap_conn.debug = 4
    imap_conn.authenticate('XOAUTH2', lambda x: auth_string)
    imap_conn.select('INBOX')


def test_smpt(user, base64_auth_string):
    smtp_conn = smtplib.SMTP('smtp.gmail.com', 587)
    smtp_conn.set_debuglevel(True)
    smtp_conn.ehlo('test')
    smtp_conn.starttls()
    smtp_conn.docmd('AUTH', 'XOAUTH2 ' + base64_auth_string)


def get_authorization(google_client_id, google_client_secret):
    scope = "https://mail.google.com/"
    print('Navigate to the following URL to auth:',
          generate_permission_url(google_client_id, scope))
    authorization_code = input('Enter verification code: ')
    response = call_authorize_tokens(
        google_client_id, google_client_secret, authorization_code)
    return response['refresh_token'], response['access_token'], response['expires_in']


def refresh_authorization(google_client_id, google_client_secret, refresh_token):
    response = call_refresh_token(
        google_client_id, google_client_secret, refresh_token)
    return response['access_token'], response['expires_in']


# Tuplas com os emails para envio da confirmação da ocorrência
nome_tupla = ('email@email.com')

TESTE_emails = ('seuemail@seuemail.com')

dict_envio = {'KEY NAME': nome_tupla, 'TESTE': TESTE_emails}

def envia_email():
    # Obtendo o período da mês em formato de string
    data_atual = date.today()
    data_em_texto = data_atual.strftime('%d/%m/%Y')

    setData = date.fromordinal(data_atual.toordinal() - 4)
    setData_barra = setData.strftime('%d/%m/%Y')
    dia_setData = setData_barra[:2]
    mes_setData = setData_barra[3:5]
    ano_setData = setData_barra[8:]

    # Inserindo o nome do bot no email
    body = f"""
    <p><b>RPA - Agendamento vs Abertura de OS</b></p>
    <br>
    <br>
    """

    # Gerando a tabela HTML para inserir no corpo do email
    print('Convertendo a planilha XLSX em HTML...\n')
    df2 = pd.read_excel('caminho_do_arquivo\\nome_do_arquivo.xlsx', dtype=str)
    tabela = df2.to_html(justify = 'center', index = False, na_rep='')
    with open('caminho_do_arquivo\\tabela.html', 'w') as tabela_1:
        tabela_1.write(tabela)
    print('Arquivo HTML gerado!\n')
    p.sleep(1)

    # Inserindo email e senha
    print('Enviando email...\n')
    remetente = 'seu email aqui'
    # get the password in the gmail (manage your google account, click on the avatar on the right)
    # then go to security (right) and app password (center)
    # insert the password and then choose mail and this computer and then generate
    # copy the password generated here
    password = 'sua senha aqui'

    receiver = dict_envio['TESTE']

    access_token, expires_in = refresh_authorization(GOOGLE_CLIENT_ID, GOOGLE_CLIENT_SECRET, GOOGLE_REFRESH_TOKEN)
    auth_string = generate_oauth2_string(remetente, access_token, as_base64=True)

    # Configurando o objeto MIME (Remetente, destinatário e assunto do email com a data do dia)
    message = MIMEMultipart('related')
    message['From'] = remetente
    message['To'] = receiver
    message['Subject'] = 'RPA - Agendamento vs Abertura de OS' + ' ' + f'({dia_setData}/{mes_setData}/{ano_setData} - {data_em_texto})'

    # Anexando a mensagem alternativa
    msgAlternative = MIMEMultipart('alternative')
    # Mensagem Alternativa está sendo jogada no message (root-principal)
    message.attach(msgAlternative)

    msgText = MIMEText(f'RPA - Agendamento vs Abertura de OS')
    msgAlternative.attach(msgText)  # msgText é jogada na mensagem alternativa

    # Inserindo a tabela no formato HTML no corpo do email
    msgText = MIMEText(
        f'{tabela}', 'html')
    msgAlternative.attach(msgText)

    # use gmail with port
    session = smtplib.SMTP('smtp.gmail.com', 587)

    session.ehlo(GOOGLE_CLIENT_ID)
    # enable security
    session.starttls()

    session.docmd('AUTH', 'XOAUTH2 ' + auth_string)

    ######### ANTIGA FORMA DE ENVIAR EMAIL - APPS MENOS SEGURO ####################
    # # login with mail_id and password
    # session.login(remetente, password)

    text = message.as_string()
    session.sendmail(remetente, receiver.split(","), text)
    session.quit()

    os.remove('caminho_do_arquivo\\tabela.html')
    os.remove('caminho_do_arquivo\\nome_do_arquivo.xlsx')
    os.remove('caminho_do_arquivo\\nome_do_arquivo.XLS')

    logging.debug('Email enviado e arquivos excluidos!')
    print('Email enviado e arquivos excluidos!')

