# -*- coding: utf-8 -*-
import pandas as pd
import logging
import os
from openai import OpenAI
import re
import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import requests
from datetime import datetime
from hubspot_api import get_last_hubspot_contact, annotate_companies_with_hubspot, get_last_company_activity

os.makedirs("mail_log", exist_ok=True)

# --- Konfiguration ---
SPREADSHEET_ID = "1ghB0Okyu3MEQizb2qyIPTTIlr29eF6ljJoQOvJM4PME"
WORKSHEET_NAME = "Kontaktliste all"
LOG_FILE = 'mail_log/mail_log.txt'

# OpenAI-Client initialisieren (Streamlit Cloud: Key aus st.secrets)
client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])

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

# Logging einrichten
logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)

SIGNATURE_HTML = """
<br><br>
<span style="color:#888; font-size:13px;">
  <b style="color:#888; font-size:15px;">Lucas Freigang</b><br>
  <span style="font-size:11px;">Head of ESG</span><br>
  <a href="https://icons.at" style="color:#1a73e8; text-decoration:none; font-size:14px;">icons – consulting by students Innsbruck</a><br>
  <br>
  <span style="color:#888;">Bürgerstraße 2 | 6020 Innsbruck | Österreich</span><br><br>
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

def get_prompt(prompt_type=None, custom_prompt=None):
    if custom_prompt:
        return custom_prompt
    if prompt_type == "mittelständisch":
        path = "resources/prompt_mittelständisch.txt"
    elif prompt_type == "klein":
        path = "resources/prompt_klein.txt"
    else:
        raise ValueError("prompt_type muss 'mittelständisch', 'klein' oder custom_prompt gesetzt sein.")
    with open(path, "r", encoding="utf-8") as f:
        return f.read()

def get_companies_via_openai_prompt(prompt):
    response = client.chat.completions.create(
        model="gpt-4.1",
        messages=[{"role": "user", "content": prompt}]
    )
    return response.choices[0].message.content

def parse_openai_response(response_text):
    companies = []
    # Splitte an Gedankenstrich (–, U+2013) oder Bindestrich (-)
    import re
    pattern = re.compile(
        r'^\s*(.*?)\s*[-–]\s*(.*?)\s*[-–]\s*(.*?)\s*[-–]\s*([\w\.-]+@[\w\.-]+\.\w+)\s*$'
    )
    lines = response_text.strip().split('\n')
    for line in lines:
        match = pattern.match(line)
        if match:
            name, website, region, email = match.groups()
            companies.append({
                'Name': name.strip(),
                'Website': website.strip(),
                'Region': region.strip(),
                'E-Mail': email.strip()
            })
    return companies

def load_company_data():
    """
    Liest die gesamte Google Sheet-Tabelle ein (ohne Filter).
    """
    data = worksheet.get_all_records()
    df = pd.DataFrame(data)
    return df

def update_sheet(companies):
    """
    Fügt neue Unternehmen als Zeilen in das Google Sheet ein.
    Die Werte werden direkt aus dem company-Dict übernommen.
    Gibt eine Liste der übersprungenen Unternehmensnamen zurück.
    """
    existing_names = set(row['Unternehmen'] for row in worksheet.get_all_records())
    new_count = 0
    skipped_names = []
    for company in companies:
        name = company.get('Name', '').strip()
        email = company.get('E-Mail', '').strip()
        region = company.get('Region', '').strip()
        website = company.get('Website', '').strip()
        gruppe = company.get('Gruppe', '').strip()
        mitglied = company.get('Name icons Mitglied', '').strip()
        letzter_kontakt_orga = company.get('Letzter Kontakt Organisation', '').strip()
        # Passe die Reihenfolge und Anzahl der Felder an dein Sheet an!
        if not name or name in existing_names:
            skipped_names.append(name)
            continue
        new_row = [
            gruppe, region, mitglied, name, email,
            '', '', '', '', '', '', letzter_kontakt_orga, '', 'Nein', ''
        ] + [''] * 4 + [website]
        worksheet.append_row(new_row)
        new_count += 1
        existing_names.add(name)
    if new_count > 0:
        print(f"{new_count} Unternehmen hinzugefügt.")
    else:
        print("Keine neuen Unternehmen hinzugefügt.")
    return skipped_names

# --- Testaufruf für die Kommandozeile ---
if __name__ == "__main__":
    firmen = ["Deloitte Österreich", "PwC Österreich"]
    for name in firmen:
        print(f"\nSuche nach Unternehmen: {name}")
        last_activity = get_last_company_activity(name)
        if last_activity:
            print(f"Letzte Aktivität für '{name}': {last_activity}")
        else:
            print(f"Kein Kontakt für '{name}' gefunden.")



