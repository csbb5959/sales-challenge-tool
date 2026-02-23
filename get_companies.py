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

SPREADSHEET_ID = "1o4vY8j2hrHyKfs1wwvyxYF172BOvgPRl5qpj_9-GsJE"
WORKSHEET_NAME = "Team Gabriel"
LOG_FILE = 'mail_log/mail_log.txt'

client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])

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
    pattern = re.compile(
        r'^\s*(.*?)\s*[-–]\s*(.*?)\s*[-–]\s*(.*?)\s*[-–]\s*([\w\.-]+@[\w\.-]+\.\w+)\s*$'
    )
    lines = response_text.strip().split('\n')
    for line in lines:
        match = pattern.match(line)
        if match:
            company_name, website, region, email = match.groups()
            companies.append({
                'Name': company_name.strip(), 
                'Website': website.strip(),
                'Region': region.strip(),
                'E-Mail': email.strip()
            })
    return companies

def update_sheet(companies):
    # ANGEPASST: Zieht den echten Spaltennamen für den Abgleich
    existing_names = set(str(row.get('Unternehmensname (laut Handelsregister)', '')) for row in worksheet.get_all_records())
    new_count = 0
    skipped_names = []
    
    for company in companies:
        company_name = company.get('Name', '').strip()
        email = company.get('E-Mail', '').strip()
        region = company.get('Region', '').strip()
        website = company.get('Website', '').strip()
        gruppe = company.get('Gruppe', '').strip()
        mitglied = company.get('Name icons Mitglied', '').strip()
        letzter_kontakt_orga = company.get('Letzter Kontakt Organisation', '').strip()
        name = company.get('Name Kontaktperson', '').strip()
        last_contact_person = company.get('Letzter Kontakt Person', '').strip()
        
        if not company_name or company_name in existing_names:
            skipped_names.append(company_name)
            continue
            
        # ANGEPASST: Exaktes Mapping auf die 17 sichtbaren Spalten im Sheet.
        # Zusatzinfos von OpenAI werden ganz ans Ende gehängt.
        new_row = [
            company_name,         # 1: Unternehmensname (laut Handelsregister)
            '',                   # 2: Tr
            name,                 # 3: Name, Nachname
            '',                   # 4: Tr
            email,                # 5: E-Mail
            '',                   # 6: Alternative E-Mail
            '',                   # 7: Alternative E-Mail 2
            '',                   # 8: Tr
            '',                   # 9: Telefonnummer
            '',                   # 10: Tr
            0,                    # 11: Anzahl Mails/ LinkedIn Nachrichten
            'FALSE',              # 12: Cold-Call?
            'FALSE',              # 13: Persönlich?
            'FALSE',              # 14: Get 2 Gether?
            'FALSE',              # 15: Proposal?
            'FALSE',              # 16: Projekt?
            0,                    # 17: Punkte
            # Angehängte Metadaten:
            website, region, gruppe, mitglied, last_contact_person, letzter_kontakt_orga
        ]
        
        worksheet.append_row(new_row)
        new_count += 1
        existing_names.add(company_name)
        
    if new_count > 0:
        print(f"{new_count} Unternehmen hinzugefügt.")
    else:
        print("Keine neuen Unternehmen hinzugefügt.")
    return skipped_names