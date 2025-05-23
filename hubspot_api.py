import requests
import streamlit as st
from datetime import datetime
import difflib
import dateutil.parser

def get_last_hubspot_contact(email=None, company_name=None):
    """
    Prüft, ob ein Kontakt mit dieser E-Mail ODER einem zur Firma passenden E-Mail-Domain in HubSpot existiert.
    Gibt (Name, E-Mail, Datum letzter Kontakt) zurück oder None.
    """
    url = "https://api.hubapi.com/crm/v3/objects/contacts/search"
    headers = {
        "Authorization": f"Bearer {st.secrets['HUBSPOT_TOKEN']}",
        "Content-Type": "application/json"
    }

    # 1. Suche nach exakter E-Mail
    if email:
        data = {
            "filterGroups": [{
                "filters": [{
                    "propertyName": "email",
                    "operator": "EQ",
                    "value": email
                }]
            }],
            "properties": ["firstname", "lastname", "email", "lastmodifieddate", "last_contacted"]
        }
        response = requests.post(url, headers=headers, json=data)
        if response.status_code == 200:
            results = response.json().get("results", [])
            if results:
                props = results[0].get("properties", {})
                name = f"{props.get('firstname', '')} {props.get('lastname', '')}".strip()
                email_addr = props.get("email", "")
                last_contacted = props.get("last_contacted") or props.get("lastmodifieddate")
                if last_contacted:
                    try:
                        last_contacted = int(last_contacted)
                        last_contacted = datetime.utcfromtimestamp(last_contacted / 1000).strftime("%Y-%m-%d")
                    except Exception:
                        pass
                return {"name": name, "email": email_addr, "date": last_contacted}
    
    # 2. Suche nach Kontakten, deren E-Mail-Domain zum Unternehmen passt
    if company_name:
        # Extrahiere einen "Kern" des Firmennamens für die Suche (z.B. "PwC" aus "PwC Österreich")
        company_token = company_name.split()[0].lower()
        # Suche alle Kontakte, filtere nach Domain
        # Achtung: HubSpot erlaubt keine Wildcard-Suche, daher paginiere alle Kontakte (max. 100 pro Seite)
        url_all = "https://api.hubapi.com/crm/v3/objects/contacts"
        after = None
        best_match = None
        while True:
            params = {
                "limit": 100,
                "properties": "firstname,lastname,email,lastmodifieddate,last_contacted"
            }
            if after:
                params["after"] = after
            resp = requests.get(url_all, headers=headers, params=params)
            if resp.status_code != 200:
                break
            data = resp.json()
            contacts = data.get("results", [])
            for c in contacts:
                props = c.get("properties", {})
                email_addr = props.get("email", "") or ""
                if company_token in email_addr.lower():
                    name = f"{props.get('firstname', '')} {props.get('lastname', '')}".strip()
                    last_contacted = props.get("last_contacted") or props.get("lastmodifieddate")
                    if last_contacted:
                        try:
                            last_contacted = int(last_contacted)
                            last_contacted = datetime.utcfromtimestamp(last_contacted / 1000).strftime("%Y-%m-%d")
                        except Exception:
                            pass
                    match = {"name": name, "email": email_addr, "date": last_contacted}
                    # Nimm den ersten besten Treffer (oder verbessere das Matching nach Bedarf)
                    if not best_match or (match["date"] and match["date"] > best_match.get("date", "")):
                        best_match = match
            after = data.get("paging", {}).get("next", {}).get("after")
            if not after:
                break
        if best_match:
            return best_match

    return None

def get_last_company_activity(company_name):
    """
    Sucht im HubSpot-Companies-Endpoint nach dem Unternehmensnamen (unscharf) und gibt den besten Treffer zurück.
    Wählt zuerst den längsten Namensmatch, dann das neueste Aktivitätsdatum.
    """
    url = "https://api.hubapi.com/crm/v3/objects/companies/search"
    headers = {
        "Authorization": f"Bearer {st.secrets['HUBSPOT_TOKEN']}",
        "Content-Type": "application/json"
    }
    main_token = company_name.split()[0]
    data = {
        "filterGroups": [{
            "filters": [{
                "propertyName": "name",
                "operator": "CONTAINS_TOKEN",
                "value": main_token
            }]
        }],
        "properties": ["name", "last_activity_date", "lastmodifieddate", "createdate"]
    }
    response = requests.post(url, headers=headers, json=data)
    if response.status_code == 200:
        results = response.json().get("results", [])
        if not results:
            return None

        # 1. Finde alle mit maximalem Namensmatch
        def match_len(r):
            hub_name = r["properties"].get("name", "").lower()
            tokens = set(company_name.lower().split())
            return len(tokens & set(hub_name.split()))
        max_match = max(match_len(r) for r in results)
        best_matches = [r for r in results if match_len(r) == max_match]

        # 2. Wähle aus diesen das mit dem neuesten Datum
        def get_best_date(r):
            props = r["properties"]
            for key in ["last_activity_date", "lastmodifieddate", "createdate"]:
                if props.get(key):
                    try:
                        return dateutil.parser.parse(props[key])
                    except Exception:
                        pass
            return dateutil.parser.parse("1900-01-01")
        best = max(best_matches, key=get_best_date)
        best_name = best["properties"].get("name", "")
        # Fallback-Logik für Datum
        last_activity = (
            best["properties"].get("last_activity_date")
            or best["properties"].get("lastmodifieddate")
            or best["properties"].get("createdate")
            or ""
        )
        print(f"OpenAI: {company_name} | HubSpot: {best_name} | last_activity: {last_activity}")
        return {
            "name": best_name,
            "last_activity_date": last_activity if last_activity else "Kein Datum gefunden"
        }
    return None

def annotate_companies_with_hubspot(companies):
    """
    Ergänzt jedes Unternehmen mit dem letzten HubSpot-Kontakt in 'Letzter Kontakt Organisation' (Spalte L).
    """
    for company in companies:
        email = company.get("E-Mail", "")
        name = company.get("Name", "")
        last_contact = get_last_hubspot_contact(email=email, company_name=name)
        if last_contact:
            company["Letzter Kontakt Organisation"] = last_contact.get("date", "Keinen Kontakt gefunden")
        else:
            company["Letzter Kontakt Organisation"] = "Keinen Kontakt gefunden"
    return companies