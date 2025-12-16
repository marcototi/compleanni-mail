import pandas as pd
import smtplib
import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import os # Serve per leggere le "variabili segrete"

# --- CONFIGURAZIONE ---
# Leggiamo le credenziali dai "Segreti" di GitHub (spiegato dopo nella guida)
email_user = os.environ.get('EMAIL_USER') 
email_password = os.environ.get('EMAIL_PASSWORD')
email_send = os.environ.get('EMAIL_USER') # Si manda la mail a se stessi
smtp_server = 'smtp.gmail.com'
smtp_port = 587

# --- FUNZIONI ---
def calculate_age(born):
    today = datetime.date.today()
    return today.year - born.year - ((today.month, today.day) < (born.month, born.day))

def send_mail():
    print("Inizio procedura invio mail...")
    
    # Lettura del file Excel (Ora cerca il file nella stessa cartella dello script)
    # Assicurati che il file si chiami ESATTAMENTE 'lista_compleanni.xlsx'
    try:
        df = pd.read_excel("lista_compleanni.xlsx", index_col=0)
    except FileNotFoundError:
        print("Errore: File lista_compleanni.xlsx non trovato!")
        return

    # Selezione delle persone che compiono gli anni oggi
    today = datetime.date.today()
    df['Data di nascita'] = pd.to_datetime(df['Data di nascita'], format='%d/%m/%Y')
    
    # Filtro magico
    today_birthdays = df[df['Data di nascita'].notnull() & df['Data di nascita'].apply(lambda x: x.strftime('%d/%m') == today.strftime('%d/%m') if not pd.isnull(x) else False)]

    # Se non ci sono compleanni e non vuoi ricevere mail vuote, togli il commento sotto:
    # if len(today_birthdays) == 0:
    #     print("Nessun compleanno oggi. Nessuna mail inviata.")
    #     return

    # Creazione del corpo della mail
    message = MIMEMultipart()
    message['From'] = email_user
    message['To'] = email_send
    message['Subject'] = 'Compleanni del giorno ' + today.strftime('%d/%m/%Y')

    if len(today_birthdays) == 0:
        body = 'Buongiorno Marco,\n\noggi non ci sono compleanni da visualizzare.\n\nPassa una buona giornata!'
    else:
        body = 'Buongiorno Marco,\n\noggi ' + today.strftime('%d/%m/%Y') + ' compiono gli anni le seguenti persone:\n\n'
        for index, row in today_birthdays.iterrows():
            age = calculate_age(row['Data di nascita'])
            body += '- ' + row['Nome'] + ' ' + row['Cognome'] + ' compie ' + str(age) + ' anni\n'

    message.attach(MIMEText(body, 'plain'))

    # Invio della mail
    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.ehlo()
        server.starttls()
        server.login(email_user, email_password)
        text = message.as_string()
        server.sendmail(email_user, email_send, text)
        server.quit()
        print('Mail inviata con successo.')
    except Exception as e:
        print(f"Errore durante l'invio della mail: {e}")

# Eseguiamo la funzione direttamente, senza schedule o while loop
if __name__ == "__main__":
    send_mail()