import streamlit as st
import pandas as pd
import time
import openai
import os
import gspread
from google.oauth2.service_account import Credentials

# Passwortschutz (ganz am Anfang der Datei einfügen)
def check_password():
    def password_entered():
        if st.session_state["password"] == st.secrets["APP_PASSWORD"]:
            st.session_state["password_correct"] = True
            del st.session_state["password"]  # Passwort aus dem Speicher löschen
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        st.text_input("Passwort", type="password", on_change=password_entered, key="password")
        st.stop()
    elif not st.session_state["password_correct"]:
        st.text_input("Passwort", type="password", on_change=password_entered, key="password")
        st.error("Falsches Passwort")
        st.stop()

check_password()

# Neue Importe für get_companies und send_emails:
from get_companies import get_companies_via_openai_prompt, parse_openai_response, update_sheet, get_prompt
from send_emails import send_mail, df, td, DELAY_SECONDS, LOG_FILE

# Google Sheets Setup
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

# Ordner für Logfile anlegen, falls nicht vorhanden
os.makedirs("mail_log", exist_ok=True)

# Schreibe die JSON-Datei aus dem Secret
SERVICE_ACCOUNT_FILE = "service_account.json"
with open(SERVICE_ACCOUNT_FILE, "w") as f:
    f.write(st.secrets["GOOGLE_SERVICE_ACCOUNT_JSON"])

# Dann verwende SERVICE_ACCOUNT_FILE für Credentials:
CREDS = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)

SHEET_JSON = 'sales-challenge-600f686a3b9b.json'
SPREADSHEET_ID = "1ghB0Okyu3MEQizb2qyIPTTIlr29eF6ljJoQOvJM4PME"
WORKSHEET_NAME = "Kontaktliste all"
gc = gspread.authorize(CREDS)
sh = gc.open_by_key(SPREADSHEET_ID)
worksheet = sh.worksheet(WORKSHEET_NAME)

@st.cache_data(ttl=60)
def load_company_data():
    data = worksheet.get_all_records()
    return pd.DataFrame(data)

def save_company_data(df):
    worksheet.clear()
    worksheet.update([df.columns.values.tolist()] + df.values.tolist())

st.title("Unternehmensakquise & Mailing Tool")

# OpenAI-Key laden
openai.api_key = st.secrets["OPENAI_API_KEY"]

# 1) Unternehmen suchen
st.header("1. Unternehmen suchen (OpenAI)")

prompt_option = st.radio(
    "Prompt auswählen:",
    ("Eigener Prompt", "Mittelständische Unternehmen", "Kleine Unternehmen")
)

# Nur bei vorgefertigten Prompts die Anzahl anzeigen
if prompt_option in ("Mittelständische Unternehmen", "Kleine Unternehmen"):
    anzahl = st.number_input(
        "Wie viele Unternehmen sollen recherchiert werden?",
        min_value=1, max_value=100, value=20, step=1
    )

with open('resources/prompt_structure.txt', 'r', encoding='utf-8') as f:
    prompt_structure = f.read()

if prompt_option == "Eigener Prompt":
    prompt = st.text_area("Eigener Prompt:") + prompt_structure
elif prompt_option == "Mittelständische Unternehmen":
    prompt = get_prompt(prompt_type="mittelständisch")
    prompt = prompt.replace("{anzahl}", str(anzahl))
    st.code(prompt)
elif prompt_option == "Kleine Unternehmen":
    prompt = get_prompt(prompt_type="klein")
    prompt = prompt.replace("{anzahl}", str(anzahl))
    st.code(prompt)

if 'companies' not in st.session_state:
    st.session_state['companies'] = []

if st.button("Unternehmen suchen"):
    if prompt:
        response_text = get_companies_via_openai_prompt(prompt)
        companies = parse_openai_response(response_text)
        st.session_state['companies'] = companies
        st.write("Gefundene Unternehmen:")
        st.write(companies)
    else:
        st.warning("Bitte gib einen Prompt ein.")

if st.session_state['companies']:
    with st.expander("Optionale Zusatzinfos für neue Unternehmen"):
        gruppe = st.text_input("Gruppe (optional)")
        region = st.text_input("Region (optional)")
        mitglied = st.text_input("Name icons Mitglied (optional)")
        # Füge weitere optionale Felder nach Bedarf hinzu

    if st.button("Gefundene Unternehmen in Tabelle eintragen"):
        # Übernehme optionale Felder, falls sie gesetzt sind
        companies_to_add = []
        for company in st.session_state['companies']:
            company = company.copy()
            if gruppe:
                company['Gruppe'] = gruppe
            if mitglied:
                company['Name icons Mitglied'] = mitglied
            companies_to_add.append(company)
        update_sheet(companies_to_add)
        st.success("Unternehmen wurden in Tabelle eingetragen.")
        st.session_state['companies'] = []

# 2) Excel-Liste anzeigen und filtern
st.header("2. Tabelle anzeigen & filtern")
st.caption(
    "💡 **Tipp:** Wenn du mit der Maus über die Tabelle fährst, erscheint oben rechts eine kleine Suchlupe. "
    "Damit kannst du in jeder Spalte direkt nach Text filtern!"
)
try:
    df = load_company_data()
except Exception as e:
    st.error("Google Sheets API-Limit erreicht. Bitte warte eine Minute und lade die Seite neu.")
    st.stop()

# Zeilen ohne Unternehmensnamen entfernen
df = df[df["Unternehmen"].notna() & (df["Unternehmen"].str.strip() != "")]

# Freitext-Filter für bestimmte Spalten
mitglied_filter = st.text_input("Filter für 'Name icons Mitglied':")
unternehmen_filter = st.text_input("Filter für 'Unternehmen':")
region_filter = st.text_input("Filter für 'Region':")

# Multiselect für weitere Spalten
other_cols = [col for col in df.columns if col not in ["Name icons Mitglied", "Unternehmen", "Region"]]
filter_cols = st.multiselect("Weitere Spalten zum Filtern auswählen:", other_cols)

filtered_df = df.copy()

# Anwenden der Freitext-Filter (case-insensitive, enthält)
if mitglied_filter:
    filtered_df = filtered_df[filtered_df["Name icons Mitglied"].str.contains(mitglied_filter, case=False, na=False)]
if unternehmen_filter:
    filtered_df = filtered_df[filtered_df["Unternehmen"].str.contains(unternehmen_filter, case=False, na=False)]
if region_filter:
    filtered_df = filtered_df[filtered_df["Region"].str.contains(region_filter, case=False, na=False)]

# Anwenden der Multiselect-Filter für andere Spalten
for col in filter_cols:
    if df[col].dtype == object:
        unique_vals = df[col].dropna().unique().tolist()
        selected_vals = st.multiselect(f"Werte für '{col}' auswählen:", unique_vals, default=unique_vals)
        filtered_df = filtered_df[filtered_df[col].isin(selected_vals)]
    else:
        min_val, max_val = int(df[col].min()), int(df[col].max())
        selected_range = st.slider(f"Wertebereich für '{col}' auswählen:", min_val, max_val, (min_val, max_val))
        filtered_df = filtered_df[(filtered_df[col] >= selected_range[0]) & (filtered_df[col] <= selected_range[1])]

# Bearbeitbare Tabelle
edit_df = filtered_df.copy()
edit_df = edit_df.dropna(axis=1, how='all')
edited_df = st.data_editor(
    edit_df,
    num_rows="dynamic",
    use_container_width=True,
    key="excel_editor"
)

if st.button("Änderungen speichern"):
    save_company_data(edited_df)
    st.success("Änderungen gespeichert!")

# 3) E-Mails senden (nur an gefilterte Unternehmen)
st.header("3. E-Mails senden (an gefilterte Auswahl)")

mail_text_option = st.radio(
    "Welchen E-Mail-Text möchtest du verwenden?",
    ("Standard-Text verwenden", "Eigenen Text eingeben")
)
if mail_text_option == "Eigenen Text eingeben":
    custom_mail_subject = st.text_input("Eigener E-Mail-Betreff (nutze {company} als Platzhalter für den Firmennamen):", value="Maßgeschneiderte Lösungen für {company}")
    custom_mail_text = st.text_area("Eigener E-Mail-Text (nutze {company} als Platzhalter für den Firmennamen):")
else:
    custom_mail_subject = None
    custom_mail_text = None

# Auswahl für E-Mail-Signatur
add_signature = st.checkbox("E-Mail-Signatur anhängen (empfohlen)", value=True)

# Optionaler Anhang per Drag & Drop
uploaded_file = st.file_uploader("Optional: Anhang (z.B. PDF) per Drag & Drop hinzufügen", type=["pdf"])

if not filtered_df.empty:
    options = [
        f"{row['Unternehmen']} ({row['E-Mail']})"
        for idx, row in filtered_df.iterrows()
    ]
    selected = st.multiselect(
        "Wähle die Unternehmen aus, die du kontaktieren möchtest:",
        options
    )
    if st.button("Ausgewählten Unternehmen E-Mails senden"):
        for idx, row in filtered_df.iterrows():
            label = f"{row['Unternehmen']} ({row['E-Mail']})"
            if label in selected:
                send_mail(
                    row['E-Mail'],
                    row['Unternehmen'],
                    mail_text=custom_mail_text if mail_text_option == "Eigenen Text eingeben" else None,
                    mail_subject=custom_mail_subject if mail_text_option == "Eigenen Text eingeben" else None,
                    attachment=uploaded_file,
                    add_signature=add_signature
                )
                st.success(f"Mail gesendet an {row['E-Mail']}")
                time.sleep(DELAY_SECONDS)
        st.info("Alle ausgewählten Mails wurden bearbeitet. Details siehe Log.")
else:
    st.info("Keine Unternehmen für den Versand gefunden.")

st.write("---")
st.write("© Lucas Freigang")