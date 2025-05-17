# Unternehmensakquise & Mailing Tool

Dieses Tool unterstützt dich dabei, Unternehmen automatisiert zu recherchieren, zu verwalten und personalisierte E-Mails zu versenden. Es basiert auf [Streamlit](https://streamlit.io/) und nutzt Google Sheets, OpenAI und optionale PDF-Anhänge.

---

## Features

- **Unternehmensrecherche** via OpenAI (mit anpassbaren Prompts)
- **Datenverwaltung** direkt in Google Sheets (Bearbeiten, Filtern, Hinzufügen)
- **Personalisierter E-Mail-Versand** (Standard- oder eigener Text, eigener Betreff, optionaler Anhang)
- **Intuitive Weboberfläche** (kein Coding nötig)

---

## Schnellstart

### 1. Voraussetzungen

- Python 3.8+
- Google-Service-Account (JSON) mit Zugriff auf das gewünschte Google Sheet
- OpenAI API-Key
- Gmail-Account für den Versand (mit App-Passwort)

### 2. Installation

```sh
git clone <REPO-URL>
cd email
python -m venv venv
source venv/bin/activate  # Windows: venv\Scripts\activate
pip install -r requirements.txt
```

### 3. Konfiguration

- Lege eine `.env`-Datei im `email/`-Ordner an (siehe `.env.example`):

    ```
    OPENAI_API_KEY=dein-openai-key
    GMAIL_USER=deine.email@gmail.com
    GMAIL_PASS=dein-app-passwort
    SMTP_HOST=smtp.gmail.com
    SMTP_PORT=587
    ```

- Lege deine Google-Service-Account-JSON (z.B. `sales-challenge-xxxx.json`) in den `email/`-Ordner.

- Passe ggf. die Sheet-ID und den Sheet-Namen in `app.py` und `get_companies.py` an.

### 4. Starten

```sh
streamlit run app.py
```

Die App öffnet sich im Browser.

---

## Hinweise

- **Sensible Daten** wie `.env` und Service-Account-JSON **niemals öffentlich teilen**!
- Die App funktioniert am besten mit Google Chrome oder Firefox.
- Für den E-Mail-Versand muss ggf. ein [App-Passwort](https://support.google.com/accounts/answer/185833?hl=de) für Gmail erstellt werden.

---

## Deployment für Nicht-Developer

- **Streamlit Cloud:** Lade das Projekt auf GitHub hoch und deploye es auf [streamlit.io/cloud](https://streamlit.io/cloud).
- **Lokale Nutzung:** Gib Nutzern die ZIP-Datei und diese Anleitung.

---

## Ordnerstruktur

```
email/
    app.py
    get_companies.py
    send_emails.py
    requirements.txt
    .env.example
    resources/
    mail_log/
    confidential/
```

---

## Lizenz

MIT License

---

**Fragen oder Feedback?**  
Melde dich bei [Lucas Freigang](mailto:lucasfreigang10@gmail.com)
