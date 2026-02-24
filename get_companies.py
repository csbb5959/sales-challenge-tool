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
    # Spalte A (Unternehmensnamen) abrufen
    try:
        col_a = worksheet.col_values(1)
    except Exception:
        col_a = []
        
    existing_names = set(str(name).strip() for name in col_a[5:] if str(name).strip())
    
    # --- FIX 1: Exakt die erste freie Zeile ab Zeile 6 finden ---
    insert_row_idx = 6
    for i in range(5, len(col_a)):
        if not str(col_a[i]).strip(): # Wenn die Zelle leer ist
            insert_row_idx = i + 1
            break
    else:
        # Wenn zwischendrin keine frei ist, hänge es ganz unten an
        insert_row_idx = max(6, len(col_a) + 1)
        
    new_count = 0
    skipped_names = []
    rows_to_insert = [] # Wir sammeln alle neuen Zeilen hier
    
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
            
        new_row = [
            company_name, name, email, '', '', '', 
            0, 'FALSE', 'FALSE', 'FALSE', 'FALSE', 'FALSE', 0,
            website, region, gruppe, mitglied, last_contact_person, letzter_kontakt_orga
        ]
        
        rows_to_insert.append(new_row)
        existing_names.add(company_name)
        new_count += 1
        
    # --- FIX 1b: Alle gesammelten Zeilen exakt an der Lücke einfügen ---
    if rows_to_insert:
        worksheet.update(range_name=f"A{insert_row_idx}", values=rows_to_insert)
        print(f"{new_count} Unternehmen ab Zeile {insert_row_idx} hinzugefügt.")
    else:
        print("Keine neuen Unternehmen hinzugefügt.")
        
    return skipped_names