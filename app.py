import streamlit as st
import pandas as pd
import time
import openai
import os
import gspread
from google.oauth2.service_account import Credentials

# Passwortschutz
def check_password():
    def password_entered():
        if st.session_state["password"] == st.secrets["APP_PASSWORD"]:
            st.session_state["password_correct"] = True
            del st.session_state["password"]
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

from get_companies import get_companies_via_openai_prompt, parse_openai_response, update_sheet, get_prompt
from send_emails import send_mail, td, DELAY_SECONDS, LOG_FILE
from hubspot_api import annotate_companies_with_hubspot, get_last_company_activity, get_last_hubspot_contact

# Google Sheets Setup
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

os.makedirs("mail_log", exist_ok=True)

SERVICE_ACCOUNT_FILE = "service_account.json"
with open(SERVICE_ACCOUNT_FILE, "w") as f:
    f.write(st.secrets["GOOGLE_SERVICE_ACCOUNT_JSON"])

CREDS = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
SPREADSHEET_ID = "1o4vY8j2hrHyKfs1wwvyxYF172BOvgPRl5qpj_9-GsJE"
WORKSHEET_NAME = "Team Gabriel" 
gc = gspread.authorize(CREDS)
sh = gc.open_by_key(SPREADSHEET_ID)
worksheet = sh.worksheet(WORKSHEET_NAME)

@st.cache_data(ttl=60)
def load_company_data():
    data = worksheet.get("A6:Q")
    
    if not data:
        return pd.DataFrame()
        
    headers = data[0]
    rows = data[1:]
    
    seen = {}
    clean_headers = []
    for i, h in enumerate(headers):
        h = str(h).replace('\n', ' ').strip()
        h = " ".join(h.split())
        
        if not h:
            h = f"Leer_{i}"
        if h in seen:
            seen[h] += 1
            clean_headers.append(f"{h}_{seen[h]}")
        else:
            seen[h] = 0
            clean_headers.append(h)
            
    max_cols = len(clean_headers)
    padded_rows = []
    for row in rows:
        padded_row = row + [''] * (max_cols - len(row))
        padded_rows.append(padded_row[:max_cols])
            
    df = pd.DataFrame(padded_rows, columns=clean_headers)
    
    for col in df.columns:
        df[col] = pd.to_numeric(df[col], errors='ignore')
        
    return df

def save_company_data(df):
    worksheet.batch_clear(["A6:Q1000"]) 
    data_to_upload = [df.columns.values.tolist()] + df.values.tolist()
    worksheet.update(range_name="A6", values=data_to_upload)

st.title("Unternehmensakquise & Mailing Tool")
openai.api_key = st.secrets["OPENAI_API_KEY"]

st.header("1. Unternehmen suchen (OpenAI)")

prompt_option = st.radio(
    "Prompt auswählen:",
    ("Eigener Prompt", "Mittelständische Unternehmen", "Kleine Unternehmen")
)

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

search_contacts = st.checkbox("Auch nach Kontaktpersonen in HubSpot suchen", value=False)
only_new_hubspot = st.checkbox("Nur Unternehmen suchen, die nicht bereits in HubSpot sind")
st.caption("⚠️ **Achtung:** Dieses Feature sucht nach konkreten Personen in HubSpot...")

if st.button("Unternehmen suchen (normal)"):
    if prompt:
        response_text = get_companies_via_openai_prompt(prompt)
        companies = parse_openai_response(response_text)
        
        filtered_companies = [] 

        for company in companies:
            hub_org = get_last_company_activity(company["Name"])
            
            if only_new_hubspot and hub_org is not None:
                continue 
            
            company["Letzter Kontakt Organisation"] = hub_org["last_activity_date"] if hub_org else "Keinen Kontakt gefunden"

            if search_contacts:
                hub_contact = get_last_hubspot_contact(email=company["E-Mail"], company_name=company["Name"])
                if hub_contact:
                    company["Name Kontaktperson"] = hub_contact.get("name", company["Name"])
                    company["E-Mail"] = hub_contact.get("email", company["E-Mail"])
                    company["Letzter Kontakt Person"] = hub_contact.get("date", "")
                else:
                    company["Letzter Kontakt Person"] = "Keine Kontaktperson gefunden"
            else:
                company["Letzter Kontakt Person"] = ""
                
            filtered_companies.append(company)

        st.session_state['companies'] = filtered_companies
        companies_df = pd.DataFrame(filtered_companies)

        def highlight_last_contact(val):
            if val and val != "Keinen Kontakt gefunden":
                return 'background-color: orange'
            else:
                return 'background-color: lightgreen'

        if not companies_df.empty:
            styled_df = companies_df.style.map(
                highlight_last_contact, subset=["Letzter Kontakt Organisation"]
            )
            st.write("Gefundene Unternehmen:")
            st.dataframe(styled_df, use_container_width=True)
        else:
            if only_new_hubspot and len(companies) > 0:
                st.warning("Alle von OpenAI vorgeschlagenen Unternehmen waren bereits in HubSpot und wurden herausgefiltert. Klicke erneut auf Suchen!")
            else:
                st.write("Keine Unternehmen gefunden.")
    else:
        st.warning("Bitte gib einen Prompt ein.")

if st.session_state['companies']:
    with st.form("unternehmen_eintragen_form"):
        submit = st.form_submit_button("Gefundene Unternehmen in Tabelle eintragen")
        if submit:
            companies_to_add = []
            for company in st.session_state['companies']:
                company = company.copy()
                companies_to_add.append(company)
            skipped = update_sheet(companies_to_add)
            if skipped:
                st.info(f"Folgende Unternehmen waren bereits im Google Sheet und wurden übersprungen:\n\n- " + "\n- ".join(skipped))
            else:
                st.success("Alle Unternehmen wurden in die Tabelle eingetragen.")
            st.session_state['companies'] = []

st.header("2. Tabelle anzeigen & filtern")
st.caption("💡 **Tipp:** Wenn du mit der Maus über die Tabelle fährst, erscheint oben rechts eine kleine Suchlupe.")

try:
    df = load_company_data()
except Exception as e:
    st.error(f"{type(e).__name__} - {e}")
    st.stop()

df = df[df["Unternehmensname (laut Handelsregister)"].notna() & (df["Unternehmensname (laut Handelsregister)"].str.strip() != "")]

unternehmen_filter = st.text_input("Filter für 'Unternehmensname':")
name_filter = st.text_input("Filter für 'Name, Nachname':")
email_filter = st.text_input("Filter für 'E-Mail':")

other_cols = [col for col in df.columns if col not in ["Unternehmensname (laut Handelsregister)", "Name, Nachname", "E-Mail"]]
filter_cols = st.multiselect("Weitere Spalten zum Filtern auswählen:", other_cols)

filtered_df = df.copy()

if unternehmen_filter:
    filtered_df = filtered_df[filtered_df["Unternehmensname (laut Handelsregister)"].str.contains(unternehmen_filter, case=False, na=False)]
if name_filter:
    filtered_df = filtered_df[filtered_df["Name, Nachname"].str.contains(name_filter, case=False, na=False)]
if email_filter:
    filtered_df = filtered_df[filtered_df["E-Mail"].str.contains(email_filter, case=False, na=False)]

for col in filter_cols:
    if df[col].dtype == object:
        unique_vals = df[col].dropna().unique().tolist()
        selected_vals = st.multiselect(f"Werte für '{col}' auswählen:", unique_vals, default=unique_vals)
        filtered_df = filtered_df[filtered_df[col].isin(selected_vals)]
    else:
        min_val, max_val = int(df[col].min()), int(df[col].max())
        selected_range = st.slider(f"Wertebereich für '{col}' auswählen:", min_val, max_val, (min_val, max_val))
        filtered_df = filtered_df[(filtered_df[col] >= selected_range[0]) & (filtered_df[col] <= selected_range[1])]

edit_df = filtered_df.copy()
edit_df = edit_df.dropna(axis=1, how='all')
edited_df = st.data_editor(edit_df, num_rows="dynamic", use_container_width=True, key="excel_editor")

if st.button("Änderungen speichern"):
    save_company_data(edited_df)
    st.success("Änderungen gespeichert!")

st.header("3. E-Mails senden (an gefilterte Auswahl)")

mail_text_option = st.radio("Welchen E-Mail-Text möchtest du verwenden?", ("Standard-Text verwenden", "Eigenen Text eingeben"))
if mail_text_option == "Eigenen Text eingeben":
    custom_mail_subject = st.text_input("Eigener E-Mail-Betreff (nutze {company} als Platzhalter):", value="Maßgeschneiderte Lösungen für {company}")
    custom_mail_text = st.text_area("Eigener E-Mail-Text (nutze {company} als Platzhalter):")
else:
    custom_mail_subject = None
    custom_mail_text = None

add_signature = st.checkbox("E-Mail-Signatur anhängen (empfohlen)", value=True)
uploaded_file = st.file_uploader("Optional: Anhang (z.B. PDF) per Drag & Drop hinzufügen", type=["pdf"])

# --- NEU: Eingabefeld für die CC-E-Mail-Adresse ---
cc_email_input = st.text_input("Optional: CC-Adresse hinzufügen (z.B. für eine Kopie an dich selbst oder das CRM):")

if not filtered_df.empty:
    options = [
        f"{row['Unternehmensname (laut Handelsregister)']} ({row['E-Mail']})"
        for idx, row in filtered_df.iterrows()
    ]
    selected = st.multiselect("Wähle die Unternehmen aus, die du kontaktieren möchtest:", options)
    
    if st.button("Ausgewählten Unternehmen E-Mails senden"):
        for idx, row in filtered_df.iterrows():
            label = f"{row['Unternehmensname (laut Handelsregister)']} ({row['E-Mail']})"
            if label in selected:
                send_mail(
                    row['E-Mail'],
                    row['Unternehmensname (laut Handelsregister)'],
                    mail_text=custom_mail_text if mail_text_option == "Eigenen Text eingeben" else None,
                    mail_subject=custom_mail_subject if mail_text_option == "Eigenen Text eingeben" else None,
                    attachment=uploaded_file,
                    add_signature=add_signature,
                    cc_email=cc_email_input if cc_email_input.strip() else None # <-- NEU übergeben
                )
                st.success(f"Mail gesendet an {row['E-Mail']}")
                time.sleep(DELAY_SECONDS)
        st.info("Alle ausgewählten Mails wurden bearbeitet. Details siehe Log.")
else:
    st.info("Keine Unternehmen für den Versand gefunden.")

st.write("---")
st.write("© Lucas Freigang")