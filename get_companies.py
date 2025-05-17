# -*- coding: utf-8 -*-
import pandas as pd
import logging
import os
from openai import OpenAI
import re
import streamlit as st
import gspread
from google.oauth2.service_account import Credentials

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
    # Akzeptiere Bindestrich, Gedankenstrich, verschiedene Leerzeichen als Trenner
    pattern = re.compile(
        r'^\s*(.*?)\s*[-–]\s*([^\s]+?\.[^\s]+?)\s*[-–]\s*(.*?)\s*[-–—]\s*([\w\.-]+@[\w\.-]+\.\w+)\s*$',
        re.IGNORECASE
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
    """
    existing_names = set(row['Unternehmen'] for row in worksheet.get_all_records())
    new_count = 0
    for company in companies:
        name = company.get('Name', '').strip()
        email = company.get('E-Mail', '').strip()
        region = company.get('Region', '').strip()
        website = company.get('Website', '').strip()
        gruppe = company.get('Gruppe', '').strip()
        mitglied = company.get('Name icons Mitglied', '').strip()
        # Passe die Reihenfolge und Anzahl der Felder an dein Sheet an!
        if not name or name in existing_names:
            continue
        new_row = [
            gruppe, region, mitglied, name, email,
            '', '', '', '', '', '', '', '', 'Nein'
        ] + [''] * 5 + [website]
        worksheet.append_row(new_row)
        new_count += 1
        existing_names.add(name)
    if new_count > 0:
        print(f"{new_count} Unternehmen hinzugefügt.")
    else:
        print("Keine neuen Unternehmen hinzugefügt.")

