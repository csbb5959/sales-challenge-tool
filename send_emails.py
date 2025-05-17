# bulk_email_sender.py
# Skript zum Einlesen einer Google Sheets-Tabelle und Versand personalisierter Mails über Gmail

import pandas as pd
import smtplib
import time
import os
import logging
import streamlit as st
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import gspread
from google.oauth2.service_account import Credentials

os.makedirs("mail_log", exist_ok=True)

# --- Konfiguration ---
SPREADSHEET_ID = "1ghB0Okyu3MEQizb2qyIPTTIlr29eF6ljJoQOvJM4PME"
WORKSHEET_NAME = "Kontaktliste all"
LOG_FILE = 'mail_log/mail_log.txt'
DELAY_SECONDS = 3  # Wartezeit zwischen den Mails (anpassbar)

# Gmail-Zugangsdaten aus st.secrets laden
td = {
    'GMAIL_USER': st.secrets["GMAIL_USER"],
    'GMAIL_PASS': st.secrets["GMAIL_PASS"],
    'SMTP_HOST': st.secrets.get("SMTP_HOST", "smtp.gmail.com"),
    'SMTP_PORT': int(st.secrets.get("SMTP_PORT", 587))
}

# Google Sheets Setup (Streamlit Cloud: Service Account aus st.secrets)
SERVICE_ACCOUNT_FILE = "service_account.json"
with open(SERVICE_ACCOUNT_FILE, "w") as f:
    f.write(st.secrets["GOOGLE_SERVICE_ACCOUNT_JSON"])

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]
CREDS = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
gc = gspread.authorize(CREDS)
sh = gc.open_by_key(SPREADSHEET_ID)
worksheet = sh.worksheet(WORKSHEET_NAME)

# --- Daten einlesen (ohne Filter) ---
data = worksheet.get_all_records()
df = pd.DataFrame(data)

# Logging einrichten
logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)

DEFAULT_MAIL_HTML = """
<html>
  <body>
    <p>Sehr geehrte Damen und Herren,</p>
    <p>Ich bin Lucas von <b>icons – consulting by students</b>, Österreichs führende studentische Unternehmensberatung, und wir unterstützen Unternehmen mit maßgeschneiderten Lösungen in Bereichen wie Strategie, Marketing, HR, IT und ESG.</p>
    <p>Wir suchen nach neuen Herausforderungen – gemeinsam mit <b>{company}</b> möchten wir ein spannendes Projekt verwirklichen und dabei einen weiteren Meilenstein setzen.</p>
    <p>Mit über <b>400 erfolgreichen Projekten</b> – von Start-ups bis DAX-Konzernen – bieten wir kreative und praxisnahe Lösungen, die echten Mehrwert schaffen.</p>
    <p>Falls wir Ihre Aufmerksamkeit geweckt haben, würde ich Ihnen gerne ein unverbindliches Gespräch anbieten, auf dessen Basis wir ein maßgeschneidertes, kostenloses Proposal für Sie ausarbeiten werden.</p>
    <p>Wir sind gespannt auf Ihre Antwort.</p>
    <p>Mit bestem Dank und freundlichen Grüßen</p>
    <p>Lucas Freigang</p>
  </body>
</html>
"""

SIGNATURE_HTML = """
<br><br>
<span style="color:#888; font-size:13px;">
  <b style="color:#888; font-size:15px;">Lucas Freigang</b><br>
  <span style="font-size:11px;">Head of ESG</span><br>
  <a href="https://icons.at" style="color:#1a73e8; text-decoration:none; font-size:14px;">icons – consulting by students Innsbruck</a><br>
  <br>
  <span>Bürgerstraße 2 | 6020 Innsbruck | Österreich</span><br><br>
  <span>
    +436607197960 | <a href="mailto:lucas.freigang@icons.at" style="color:#888;">lucas.freigang@icons.at</a> | <a href="https://icons.at" style="color:#888;">icons.at</a>
  </span>
  <br>
  <hr style="border:0; border-top:1px solid #ccc;">
  <span style="font-size:10px;">Vereinsbehörde: LPD Tirol | ZVR-Zahl: 542695411</span><br>
  <span style="font-size:9px; color:gray;">
    This e-mail message may contain confidential and/or privileged information. If you are not an addressee or otherwise authorized to receive this message, you should not use, copy, disclose or take any action based on this e-mail or any information contained in the message. If you have received this material in error, please advise the sender immediately by reply e-mail and delete this message. Thank you.
  </span>
</span>
"""

def convert_text_to_html(text, company):
    # Ersetze {company} und wandle Zeilenumbrüche in <br> um, dann in <p> Absätze
    text = text.format(company=company)
    paragraphs = text.split('\n')
    html_paragraphs = [f"<p>{line.strip()}</p>" for line in paragraphs if line.strip()]
    return "<html><body>" + "\n".join(html_paragraphs) + "</body></html>"

def add_signature_to_html(html, signature):
    import re
    return re.sub(r"</body\s*>", signature + "</body>", html, flags=re.IGNORECASE)

def send_mail(recipient, company, mail_text=None, mail_subject=None, attachment=None, add_signature=True):
    try:
        msg = MIMEMultipart()
        msg['From'] = td['GMAIL_USER']
        msg['To'] = recipient
        subject = mail_subject.format(company=company) if mail_subject else f"Maßgeschneiderte Lösungen für {company}"
        msg['Subject'] = subject

        # Verwende eigenen Text oder Default
        if mail_text:
            html = convert_text_to_html(mail_text, company)
        else:
            html = DEFAULT_MAIL_HTML.format(company=company)

        # Signatur optional anhängen
        if add_signature:
            html = add_signature_to_html(html, SIGNATURE_HTML)

        msg.attach(MIMEText(html, 'html'))

        # Optionaler Anhang
        if attachment is not None:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename="{attachment.name}"')
            msg.attach(part)

        # Verbindung zum SMTP-Server aufbauen
        server = smtplib.SMTP(td['SMTP_HOST'], td['SMTP_PORT'])
        server.starttls()
        server.login(td['GMAIL_USER'], td['GMAIL_PASS'])
        server.send_message(msg)
        server.quit()

        logging.info(f"Erfolgreich gesendet an {recipient} ({company})")
        print(f"Mail gesendet an {recipient}")
    except Exception as e:
        logging.error(f"Fehler beim Senden an {recipient} ({company}): {e}")
        print(f"Fehler beim Senden an {recipient}: {e}")

# Hauptfunktion: Alle Zeilen durchlaufen (ohne Filter)
if __name__ == '__main__':
    if df.empty:
        print("Keine Unternehmen zum Kontaktieren gefunden.")
    else:
        for idx, row in df.iterrows():
            email = row['E-Mail']
            company = row['Unternehmen']
            send_mail(email, company)
            time.sleep(DELAY_SECONDS)

        print("Alle Mails wurden bearbeitet. Details im Log-File.")

