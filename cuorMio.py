import os
import pickle
import datetime
import json
import sys
from zoneinfo import ZoneInfo
import time
import requests
from dateutil import parser

# ============== CARICAMENTO VARIABILI D'AMBIENTE ==============
SHOP_URL = os.environ.get("SHOP_URL", "")  
API_VERSION = os.environ.get("API_VERSION", "")      # es: "2023-04"
ACCESS_TOKEN = os.environ.get("ACCESS_TOKEN", "")     # Token Shopify

AC_BASE_URL = os.environ.get("AC_BASE_URL", "")  
AC_API_KEY = os.environ.get("AC_API_KEY", "")         # Chiave ActiveCampaign

# ============== CONFIGURAZIONE GOOGLE SHEETS ==============
SPREADSHEET_ID = "1vqX3vOoQgIeJu9nSwLw11Y3UU_-YTFVu_V8wwHLUvkA"
SHEET_NAME = os.environ.get("SHEET_NAME", "")         # Nome del foglio da environment
CLIENT_SECRET_FILE = "client_secret.json"
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# Dati per client_secret (Google). In molti casi potresti prendere anche questi da env, ma qui li lasciamo così.
client_secret_data = {
    "installed": {
        "client_id": "1023871584063-q6d2c00ea3ig0u3d7b5tj2a43do5bif5.apps.googleusercontent.com",
        "project_id": "trascrizione-intervista-rds",
        "auth_uri": "https://accounts.google.com/o/oauth2/auth",
        "token_uri": "https://oauth2.googleapis.com/token",
        "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
        "client_secret": "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",
        "redirect_uris": ["http://localhost"]
    }
}

if not os.path.exists(CLIENT_SECRET_FILE):
    with open(CLIENT_SECRET_FILE, "w") as f:
        json.dump(client_secret_data, f, indent=4)
    print(f"{CLIENT_SECRET_FILE} creato con i dati in chiaro.")


# ============== IMPORT GOOGLE E SHOPIFY ==============
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
import shopify


# ============== FUNZIONI SUPPORTO ACTIVECAMPAIGN ==============
def create_field_value(contact_id, field_id, value):
    """
    Crea un fieldValue (record custom field) su ActiveCampaign.
    """
    headers = {"Api-Token": AC_API_KEY, "Content-Type": "application/json"}
    url = f"{AC_BASE_URL}/api/3/fieldValues"
    payload = {
        "fieldValue": {
            "contact": contact_id,
            "field": field_id,
            "value": value
        }
    }
    response = requests.post(url, headers=headers, json=payload)
    if response.status_code in [200, 201]:
        data = response.json()
        created_id = data.get("fieldValue", {}).get("id")
        print(f"[OK] FieldValue creato per contatto {contact_id}, campo {field_id} con valore '{value}'. Nuovo ID: {created_id}")
        return created_id
    else:
        print(f"[ERRORE] Creazione fallita per contatto {contact_id}, campo {field_id}: {response.text}")
        return None

def get_contact_by_email(email):
    """
    Cerca il contatto su ActiveCampaign via email.
    """
    headers = {"Api-Token": AC_API_KEY}
    url = f"{AC_BASE_URL}/api/3/contacts?filters[email]={email}"
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        data = response.json()
        contacts = data.get("contacts", [])
        if contacts:
            return contacts[0]
        else:
            print("Nessun contatto trovato per email:", email)
            return None
    else:
        print("Errore nella richiesta del contatto:", response.text)
        return None


# ============== FUNZIONI SUPPORTO GOOGLE SHEETS ==============
def init_google_sheets_service():
    creds = None
    if os.path.exists("token.pickle"):
        # Carica le credenziali già salvate
        with open("token.pickle", "rb") as token:
            creds = pickle.load(token)
        print("Caricato token.pickle esistente")

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
            print("Token rinnovato con successo, salvato in token.pickle.")
        else:
            # Su Render non possiamo completare l'OAuth interattivo,
            # quindi possiamo sollevare un errore o loggare un messaggio:
            raise RuntimeError("Nessun token valido e non posso fare OAuth su Render.")
        
        with open("token.pickle", "wb") as token:
            pickle.dump(creds, token)
    
    service = build("sheets", "v4", credentials=creds)
    print("Google Sheets service inizializzato.")
    return service


def read_sheet(service):
    sheet = service.spreadsheets()
    result = sheet.values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=SHEET_NAME
    ).execute()
    return result.get('values', [])

def cerca_data_primo_ordine(service, email):
    """
    Cerca nel Google Sheet la data del primo ordine (riga dove colonna B=1 e colonna C = email).
    """
    rows = read_sheet(service)
    for row in rows:
        if len(row) >= 3 and row[2].strip().lower() == email.strip().lower() and row[1] == "1":
            return row[0]
    return "non trovata"

def insert_row_to_sheet(service, values_list):
    """
    Inserisce una riga in coda al foglio.
    """
    current_data = read_sheet(service)
    next_row = len(current_data) + 1
    range_to_update = f"{SHEET_NAME}!A{next_row}:K{next_row}"
    body = {'values': [values_list]}
    result = service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=range_to_update,
        valueInputOption="USER_ENTERED",
        body=body
    ).execute()
    print(f"Riga inserita in {range_to_update}: {values_list}")
    return result


# ============== FUNZIONI SUPPORTO SHOPIFY ==============
def init_shopify_session():
    session = shopify.Session(SHOP_URL, API_VERSION, ACCESS_TOKEN)
    shopify.ShopifyResource.activate_session(session)
    print("Shopify session inizializzata.")

def get_last_week_range():
    """
    Calcola l'intervallo da mercoledì 00:00 (della settimana scorsa) a martedì 23:59 (di questa settimana).
    Supponendo che lo script giri di martedì sera.
    """
    now = datetime.datetime.now()  # es. martedì alle 23:00/23:59
    # Voglio calcolare:
    #  - mercoledì 00:00 (6 giorni fa)
    #  - martedì 23:59 di oggi

    # "today" come data (senza orario)
    today_date = now.date()
    # se ad esempio today_date = 2025-04-22 (un martedì)
    # mercoledì scorso = today_date - 6 giorni (2025-04-16)
    wednesday_date = today_date - datetime.timedelta(days=6)

    # mercoledì 00:00
    start = datetime.datetime(
        wednesday_date.year, wednesday_date.month, wednesday_date.day,
        0, 0, 0
    )

    # martedì 23:59
    end = datetime.datetime(
        today_date.year, today_date.month, today_date.day,
        23, 59, 59
    )

    # Convertiamo in stringhe ISO con suffisso Z
    created_at_min = start.isoformat() + "Z"
    created_at_max = end.isoformat() + "Z"
    return created_at_min, created_at_max

def get_orders_in_range():
    """
    Recupera gli ordini dall'ultima settimana:
      - da mercoledì 00:00 a martedì 23:59 (ora locale del server).
    """
    created_at_min, created_at_max = get_last_week_range()
    orders = shopify.Order.find(
        created_at_min=created_at_min,
        created_at_max=created_at_max,
        status='any',
        limit=250
    )
    orders.reverse()
    print(f"Trovati {len(orders)} ordini da {created_at_min} a {created_at_max}.")
    return orders

def extract_order_info(order):
    dataAcquisto = order.created_at
    email = getattr(order, 'email', None)
    if not email and hasattr(order, 'customer') and order.customer:
        email = order.customer.email

    landing = getattr(order, 'landing_site', "") or ""
    landing_lower = ""
    campagnaDiProvenienza = ""
    if landing and "utm_campaign=" in landing:
        try:
            campagnaDiProvenienza = landing.split("utm_campaign=")[1].split("&")[0]
            landing_lower = landing.lower()
        except Exception:
            campagnaDiProvenienza = ""

    canale_di_provenienza = ""
    if "facebook" in landing_lower or "fbclid" in landing_lower:
        canale_di_provenienza = "FB"
    elif "instagram" in landing_lower:
        canale_di_provenienza = "IG"
    else:
        canale_di_provenienza = getattr(order, 'source_name', "")

    totaleOrdine = getattr(order, 'total_price', "")
    nomeProdotto = ""
    if order.line_items and len(order.line_items) > 0:
        nomeProdotto = order.line_items[0].title

    return {
        "dataAcquisto": dataAcquisto,
        "email": email,
        "campagnaDiProvenienza": campagnaDiProvenienza,
        "canale_di_provenienza": canale_di_provenienza,
        "totaleOrdine": totaleOrdine,
        "nomeProdotto": nomeProdotto
    }



def check_run_script():
    # Ora corrente in fuso "Europe/Rome"
    rome_now = datetime.datetime.now(ZoneInfo("Europe/Rome"))
    # Controlliamo se e' martedi e ora=23:59
    # schedule.run_pending() gira ogni tot secondi, quindi dobbiamo
    # concedere un "range" di orari (23:59 ± 1 minuto) oppure esatto

    if rome_now.weekday() == 1:  # 0=lunedì, 1=martedì, ...
        if rome_now.hour == 23 and rome_now.minute == 59:
            print("** E’ martedi 23:59 in Italia! Eseguo run_script() **")
            run_script()
        else:
            print(f"Ora Roma {rome_now}, non e’ martedi 23:59, skip.")
    else:
        print(f"Ora Roma {rome_now}, non e’ martedi, skip.")


# ============== FUNZIONE PRINCIPALE DI LAVORO ==============
def run_script():
    """
    Esegue la logica Shopify -> Google Sheets -> ActiveCampaign
    prendendo gli ordini dall'ultima settimana (mercoledì -> martedì).
    """
    init_shopify_session()
    gs_service = init_google_sheets_service()

    orders = get_orders_in_range()
    
    for order in orders:
        time.sleep(4)  # Per evitare rate limit
        if not hasattr(order, 'customer') or order.customer is None:
            continue

        orders_count = getattr(order.customer, 'orders_count', 0)
        if orders_count == 0 and hasattr(order.customer, 'id'):
            customer = shopify.Customer.find(order.customer.id)
            orders_count = customer.orders_count

        if orders_count not in [1, 2]:
            continue

        order_info = extract_order_info(order)
        raw_date = order_info["dataAcquisto"]
        date_obj = parser.parse(raw_date)
        formatted_date = date_obj.strftime("%d/%m/%Y")

        ac_contact = get_contact_by_email(order_info["email"])
        if not ac_contact:
            print(f"Contatto ActiveCampaign non trovato per {order_info['email']}")
            continue

        ac_contact_id = ac_contact.get("id")

        if orders_count == 1:
            tot = f"€ {order_info['totaleOrdine']}"
            new_row = [
                formatted_date,
                "1",
                order_info["email"],
                order_info["campagnaDiProvenienza"],
                order_info["canale_di_provenienza"],
                None, None, None, None,
                tot,
                order_info["nomeProdotto"]
            ]
            insert_row_to_sheet(gs_service, new_row)

            # PRIMO ACQUISTO => field 39 (Data primo acquisto) + field 38 (Data ultimo acquisto)
            id_field39 = create_field_value(ac_contact_id, "39", formatted_date)
            id_field38 = create_field_value(ac_contact_id, "38", formatted_date)
            print(f"FieldValue creato: id {id_field39} (Data primo acquisto), id {id_field38} (Data ultimo acquisto).")

        elif orders_count == 2:
            tot = f"€ {order_info['totaleOrdine']}"
            data_primo_ordine = cerca_data_primo_ordine(gs_service, order_info["email"])
            new_row = [
                formatted_date,
                "2",
                order_info["email"],
                order_info["campagnaDiProvenienza"],
                order_info["canale_di_provenienza"],
                None, None,
                data_primo_ordine,
                None,
                tot,
                order_info["nomeProdotto"]
            ]
            insert_row_to_sheet(gs_service, new_row)

            # SECONDO ACQUISTO => field 38 (Data ultimo acquisto) + field 40 (Data secondo acquisto)
            id_field38 = create_field_value(ac_contact_id, "38", formatted_date)
            id_field40 = create_field_value(ac_contact_id, "40", formatted_date)
            print(f"FieldValue creato: id {id_field38} (Data ultimo acquisto), id {id_field40} (Data secondo acquisto).")

    print("Processo completato.")


# ============== SCHEDULER (main) ==============
def main():
    import schedule
    # Pianifichiamo di controllare ogni minuto
    schedule.every(1).minutes.do(check_run_script)
    print("Scheduler avviato (controllo ogni minuto l’ora in Europe/Rome).")

    while True:
        schedule.run_pending()
        time.sleep(10)
        


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print("Si è verificato un errore:", e)
        sys.exit(1)
