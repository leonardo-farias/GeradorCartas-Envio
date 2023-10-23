import win32com.client as win32
import pandas as pd
import time
import pyautogui as pg

# Caminho BD de e-mails
df = pd.read_excel(r'C:/Users/leonardo.farias/Desktop/Projetos/Gerador de Cartas/Modelo de base de dados email.xlsx')

# Caminho para image
logo_path = r'C:/Users/leonardo.farias/Desktop/Projetos/Gerador de Cartas/img/logo jpg.jpg'

# Crição de loop pelo for que percorre com o enumerate, lendo linha a linha de acordo com o i(número da linha)
for i, contato in enumerate(df['EMAIL']):
    indice = df.loc[i, "ID"]
    arquivo = df.loc[i, "ANEXO"]
    cliente = df.loc[i, "Cliente"]
    CCemail = df.loc[i, "CC"]
    
    outlook = win32.Dispatch('outlook.application')

    email = outlook.CreateItem(0)
    email.To = contato
    #email.CC = CCemail
    email.Subject = f'Teste email {cliente} - {indice}'
    email.HTMLBody = f"""
    <img src="{logo_path}">
    <p>corpo do email</p>
    """
    email.Attachments.Add(arquivo)

    start_time = time.time()  # Registre o tempo de início

    email.display()
    time.sleep(1)

    # Simule o atalho de teclado Ctrl+Enter+Enter para enviar o e-mail
    #pg.hotkey('ctrl', 'enter', 'enter')
    
    # Aguarde um momento para que o e-mail seja enviado antes de prosseguir
    time.sleep(2)


    end_time = time.time()  # Registre o tempo de término
    send_time = end_time - start_time  # Calcule o tempo gasto para enviar o e-mail

    print(f'E-mail {indice} enviado em {send_time:.2f} segundos')

print('Todos os e-mails foram enviados com sucesso!')
