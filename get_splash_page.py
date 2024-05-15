####################################################################################
# Criado por: Diego Fukayama
# Função: Enviar um relatório mensalmente dos usuários conectados na splash page
# O envio é programado no Agendador de Tarefas do Windows ou no Crontab do Linux
####################################################################################

import os
import requests
import pandas as pd
from datetime import datetime, timedelta
import pytz
import calendar
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from dotenv import load_dotenv

# Carregar variáveis de ambiente do arquivo .env
load_dotenv()

def get_splash_login_attempts(network_id, timespan, bearer_token):
    url = f"https://api.meraki.com/api/v1/networks/{network_id}/splashLoginAttempts?timespan={timespan}"
    headers = {
        "Authorization": f"Bearer {bearer_token}",
        "Content-Type": "application/json"
    }

    response = requests.get(url, headers=headers)

    if response.status_code == 200:
        return response.json()
    else:
        print(f"Failed to fetch data. Status code: {response.status_code}")
        return None

def convert_to_gmt_minus_3(time_string):
    # Converte a string de tempo em formato datetime
    time = datetime.strptime(time_string, "%Y-%m-%dT%H:%M:%S.%fZ")

    # Define o fuso horário para GMT-3
    timezone = pytz.timezone('America/Sao_Paulo')

    # Converte o tempo para GMT-3
    time_gmt_minus_3 = time.replace(tzinfo=pytz.utc).astimezone(timezone)

    return time_gmt_minus_3.strftime("%d/%m/%Y %H:%M:%S")

def send_email(file_path, recipient_email, sender_email, sender_password, smtp_server, smtp_port):
    # Configuração da mensagem
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = recipient_email
    msg['Subject'] = "Arquivo de login Splash"

    # Anexando o arquivo Excel
    with open(file_path, 'rb') as attachment:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename= {os.path.basename(file_path)}')
        msg.attach(part)

    # Conectando-se ao servidor SMTP
    server = smtplib.SMTP(smtp_server, smtp_port)
    server.starttls()

    # Login no servidor SMTP
    server.login(sender_email, sender_password)

    # Enviando e-mail
    server.send_message(msg)
    server.quit()

if __name__ == "__main__":
    network_id = "<NETWORK_ID>"
    timespan = 2592000 # 1 mês em segundos
    bearer_token = os.getenv("BEARER_TOKEN")

    splash_login_attempts = get_splash_login_attempts(network_id, timespan, bearer_token)

    if splash_login_attempts:
        # Filtrar dados por 'authorization' igual a 'success'
        success_attempts = [attempt for attempt in splash_login_attempts if attempt['authorization'] == 'success']

        # Aplicar a função de conversão para 'loginAt'
        for attempt in success_attempts:
            attempt['loginAt'] = convert_to_gmt_minus_3(attempt['loginAt'])

        # Criar DataFrame pandas
        df = pd.DataFrame(success_attempts)

        # Obter o nome do mês, dia do mês e ano do mês passado
        last_month = datetime.now() - timedelta(days=30)
        month_name = calendar.month_name[last_month.month]
        day_last_month = last_month.day
        year_last_month = last_month.year

        # Definir o caminho completo para o diretório
        directory = r"<DIRECTORY_PATH_TO_SAVE>"
        
        # Criar o diretório se não existir
        if not os.path.exists(directory):
            os.makedirs(directory)

        # Definir o nome completo do arquivo com o caminho
        file_name = f"splash_login_success_{day_last_month:02d}_{month_name.lower()}_{year_last_month}.xlsx"
        file_path = os.path.join(directory, file_name)

        # Salvar em formato Excel
        df.to_excel(file_path, index=False)
        print(f"Dados salvos com sucesso em {file_path}")

        # Enviar o arquivo Excel por e-mail
        recipient_email = os.getenv("RECIPIENT_EMAIL")
        sender_email = os.getenv("SENDER_EMAIL")
        sender_password = os.getenv("SENDER_PASSWORD")
        smtp_server = os.getenv("SMTP_SERVER")
        smtp_port = int(os.getenv("SMTP_PORT"))

        send_email(file_path, recipient_email, sender_email, sender_password, smtp_server, smtp_port)
        print("E-mail enviado com sucesso.")
