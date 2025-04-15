import os
import pickle
import datetime
import json
import sys
import time
import requests  # Per le chiamate HTTP a ActiveCampaign
from dateutil import parser

# ===== LEGGIAMO LE VARIABILI SENSIBILI DALL'AMBIENTE =====
SHOP_URL = os.environ.get("SHOP_URL", "")  # es. cuormio.myshopify.com
API_VERSION = os.environ.get("API_VERSION", "")                   # Versione API Shopify (fisso)
ACCESS_TOKEN = os.environ.get("ACCESS_TOKEN", "")  # Token Shopify

AC_BASE_URL = os.environ.get("AC_BASE_URL", "")  # es. https://petwellnessdilucaderiu.api-us1.com
AC_API_KEY = os.environ.get("AC_API_KEY", "")    # Chiave ActiveCampaign

# ----- CONFIGURAZIONE GOOGLE SHEETS -----
SPREADSHEET_ID = "1vqX3vOoQgIeJu9nSwLw11Y3UU_-YTFVu_V8wwHLUvkA"
SHEET_NAME = os.environ.get("SHEET_NAME", "")
CLIENT_SECRET_FILE = "client_secret.json"
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']


# ----- CONFIGURAZIONE CLIENT SECRET (Google) -----
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

# ----- IMPORT PER GOOGLE SHEETS -----
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

# ----- IMPORT PER SHOPIFY -----
import shopify

# ====================================================
# FUNZIONE PER CREARE FIELDVALUE SU ACTIVECAMPAIGN
# ====================================================

def create_field_value(contact_id, field_id, value):
    """
    Esegue una POST per creare un fieldValue con:
      - contact: l'ID del contatto
      - field: l'ID del campo custom
      - value: il valore da assegnare (ad esempio la data dell'acquisto)
    Stampa e restituisce l'ID del record creato (se creato correttamente).
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

# ====================================================
# FUNZIONE PER CERCARE IL CONTATTO DA EMAIL
# ====================================================

def get_contact_by_email(email):
    """
    Recupera il contatto da ActiveCampaign filtrando per email.
    Restituisce il primo contatto trovato oppure None.
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

# ====================================================
# FUNZIONI PER GOOGLE SHEETS
# ====================================================

def init_google_sheets_service():
    """
    Inizializza il servizio per Google Sheets utilizzando OAuth.
    """
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CLIENT_SECRET_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    service = build('sheets', 'v4', credentials=creds)
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
    rows = read_sheet(service)
    for row in rows:
        if len(row) >= 3 and row[2].strip().lower() == email.strip().lower() and row[1] == "1":
            return row[0]
    return "non trovata"

def insert_row_to_sheet(service, values_list):
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

# ====================================================
# FUNZIONI PER SHOPIFY E ORDINI
# ====================================================

def init_shopify_session():
    """
    Inizializza la sessione per Shopify utilizzando il dominio, la versione API e l'ACCESS_TOKEN.
    """
    session = shopify.Session(SHOP_URL, API_VERSION, ACCESS_TOKEN)
    shopify.ShopifyResource.activate_session(session)
    print("Shopify session inizializzata.")

def get_orders_in_range():
    """
    Recupera gli ordini da Shopify in un intervallo di date definito.
    Modifica le date in base alle tue necessità.
    """
    created_at_min = "2025-04-11T00:00:00Z"
    created_at_max = "2025-04-14T23:59:59Z"
    orders = shopify.Order.find(created_at_min=created_at_min, created_at_max=created_at_max, status='any', limit=250)
    orders.reverse()  # Ordine cronologico ascendente
    print(f"Trovati {len(orders)} ordini dal 11/04/2025 al 14/04/2025.")
    return orders

def extract_order_info(order):
    """
    Estrae informazioni rilevanti dall'ordine.
    """
    dataAcquisto = order.created_at
    email = getattr(order, 'email', None)
    if not email and hasattr(order, 'customer') and order.customer:
        email = order.customer.email
    landing = getattr(order, 'landing_site', "") or ""
    landing_lower = ""
    if landing and "utm_campaign=" in landing:
        try:
            campagnaDiProvenienza = landing.split("utm_campaign=")[1].split("&")[0]
            landing_lower = landing.lower()
        except Exception as e:
            campagnaDiProvenienza = ""
    else:
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

# ====================================================
# FUNZIONE MAIN
# ====================================================

def main():
    init_shopify_session()
    gs_service = init_google_sheets_service()
    
    orders = get_orders_in_range()
    
    for order in orders:
        time.sleep(4)  # Per evitare di superare i rate limit
        if not hasattr(order, 'customer') or order.customer is None:
            continue

        orders_count = getattr(order.customer, 'orders_count', 0)
        if orders_count == 0 and hasattr(order.customer, 'id'):
            customer = shopify.Customer.find(order.customer.id)
            orders_count = customer.orders_count

        # Ci interessa solo il 1° o 2° acquisto
        if orders_count not in [1, 2]:
            continue

        order_info = extract_order_info(order)
        raw_date = order_info["dataAcquisto"]
        date_obj = parser.parse(raw_date)
        # Formato data (dd/mm/YYYY). Se necessario, cambia in %Y-%m-%d
        formatted_date = date_obj.strftime("%d/%m/%Y")
        
        # Cerchiamo il contatto in AC via email
        ac_contact = get_contact_by_email(order_info["email"])
        if not ac_contact:
            print(f"Contatto ActiveCampaign non trovato per {order_info['email']}")
            continue
        
        ac_contact_id = ac_contact.get("id")
        
        # Aggiorniamo Google Sheet
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
            
            # PRIMO acquisto => POST a field 39 (Data primo acquisto) e field 38 (Data ultimo acquisto)
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
            
            # SECONDO acquisto => POST a field 38 (Data ultimo acquisto) e field 40 (Data secondo acquisto)
            id_field38 = create_field_value(ac_contact_id, "38", formatted_date)
            id_field40 = create_field_value(ac_contact_id, "40", formatted_date)
            print(f"FieldValue creato: id {id_field38} (Data ultimo acquisto), id {id_field40} (Data secondo acquisto).")
    
    print("Processo completato.")

def main():
    import schedule
    
    # Pianifichiamo run_script() ogni martedì alle 23:00
    schedule.every().tuesday.at("23:00").do(run_script)
    
    print("Scheduler avviato. Ogni martedì alle 23:00 partirà il job.")
    
    # Ciclo infinito per tenere viva l'app e gestire i job
    while True:
        schedule.run_pending()
        time.sleep(60)  # Controlla ogni 60 secondi se è ora di eseguire un job
        print("ancora no")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print("Si è verificato un errore:", e)
        sys.exit(1)
