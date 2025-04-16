import os
import datetime
import json
import sys
import time
import requests
from dateutil import parser
from zoneinfo import ZoneInfo
import schedule

# ======== Variabili d'Ambiente e Configurazioni Principali ========
SHOP_URL = os.environ.get("SHOP_URL", "")
API_VERSION = os.environ.get("API_VERSION", "")  # es. "2023-04"
ACCESS_TOKEN = os.environ.get("ACCESS_TOKEN", "")

AC_BASE_URL = os.environ.get("AC_BASE_URL", "")
AC_API_KEY = os.environ.get("AC_API_KEY", "")

SPREADSHEET_ID = "1vqX3vOoQgIeJu9nSwLw11Y3UU_-YTFVu_V8wwHLUvkA"
SHEET_NAME = os.environ.get("SHEET_NAME", "")  # Nome del foglio

# Il file di Service Account (inserito come stringa multilinea; meglio usarlo come file separato)
SERVICE_ACCOUNT_FILE = """
{
  "type": "service_account",
  "project_id": "trascrizione-intervista-rds",
  "private_key_id": "f73a9d20338b08a010866c56a6f735cea6b56d7f",
  "private_key": "-----BEGIN PRIVATE KEY-----\\nMIIEvgIBADANBgkqhkiG9w0BAQEFAASCBKgwggSkAgEAAoIBAQCgy6cY/uROJQKo\\nYeNAcx+caiZU1ydZtILbr4s62GJXemHHP3nEaclbK2mIS5wqKc5jjWAcqbMTAI4F\\nq939kM4DeMNmEnMqdkhr0fhHkulq+QNmCA3/Fc8o9Qwc4rD4g0NHMBHKR4o+Jj5W\\nfKxsZL+2NBk7h6BAmIjGBXfjgCaTt2OrmsiPPyhDa5nQlelTVwXiRdLf5l/t7+pb\\n4O3A8dmhyzuLhukSv1YDJAZDUICTcclCiOvri5F2w6eippP5MwSTyUsEY+0GVeI5\\nJL7IlgEoWiuCtqrCS5gA515a6RKsqoM1/69K1bCeYRtk243GwLdSOAbxtgqLqA8V\\ntG6S+d+5AgMBAAECggEABISkbdnfrWhx0ixp98okVb9P02t2QhmF4clldqJU5RNd\\nwv0AHWpBi6vFG9zQBwlEsNxsmnGURBDsbLFfG/xhJYzTpL8Y+FT5hPoR6WTx5R0Z\\nINlSF1xUBVkZXYhrI5iAn/P0VAQ9mLB3aPO43pTYJDUDjn4pnRcMJNBLhZt4ugbO\\nN+QRVJyYhku4GcecFyzEhKa0eug28lV0g+RUlDCtnFHbYhsZ5PijC8NF6uk0dRUC\\nwgWeb91ZD/6beKvfaEA73V8YjjzhFuHBMV+hsQ4WaquJ1WssQ1teQvKq5Dp+4hq1\\nEQtn/ZSubA9VHljCJU76iM+CI9B12RSw7kRXfZA2gQKBgQDcRkzyAdLzqDvVw8s6\\nyEQlqpFifx+SSWJWoKLwUfpKGsO0UEKNULlcV45cIN4/6eCM6/7SyZhLFCa0qdeA\\nhqlnI7Slql1u+byO5jd6r2hYLybJ0PLf2Opd9FWGq3fE40qFqrQzz9gWogb6hiJW\\n65uAi71PsWGkK1d/LlOOXSoMGQKBgQC6387QsFOC0YIl2UGy0C9Y9xmx1un+tYbb\\nEYYTDDqqArCLegkRTd7FIr9YOlcX0iotQAkqX37J6tOEn4MXmcpPtm3EGPly5Dxz\\noKlqpk+5IhP2lnJxPdB+VQHpogsKjtYPyjQxLXb2iaCnn0d/1KnteRxCtckDQFXX\\n2dsbBxLkoQKBgHsoCylb/7gfnaS9Hcm14vQ0U6kAboR55zOMCM3Y59m68STFoxAj\\nzB9nDL9R2TFe8B+aaxUrhaykjaeBNm4z3E9AVWYyxJ6hnt0+tlIv9GUpp8Q6wTCK\\ntS7mx1LOV96LPkVR1gMJ+EVfPgugJ171yDGs76G5CWCiov8GxczZJgMxAoGBAJYc\\nWeFBApQ+/zCwCBo/KQlp1HYKkQRNhPpMZUq/tBAFARPI/6eqyZvJgbK5imRUKhUX\\nL0WeWBaSTz5lc8RtgRnvDNVMynQD6ptnHy/QUJICUc7uoxdb9DLGzjaCOCRPAJzG\\nbI5kWv9HJon/ZEvG5IkhlBXyOHooH8y3700SrZaBAoGBAMuPUpJRl356thYU3/vP\\nVu4XDwwPy+mRyjE1jw+3STlmW96JbDUT+nmA2TysM0MVLlrEMKsnZl6Xq5piVjB9\\n9TfE3LCEBntFGxUgRRTInI6AEjv9D43xl323S6pXKivRTzkYlXQEH6v4yPY70Lc3\\nYk07YhItfZPe3pjYHvYmuTD6\\n-----END PRIVATE KEY-----\\n",
  "client_email": "kpi-marketing-automation@trascrizione-intervista-rds.iam.gserviceaccount.com",
  "client_id": "111553854552653368425",
  "auth_uri": "https://accounts.google.com/o/oauth2/auth",
  "token_uri": "https://oauth2.googleapis.com/token",
  "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
  "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/kpi-marketing-automation%40trascrizione-intervista-rds.iam.gserviceaccount.com",
  "universe_domain": "googleapis.com"
}
"""
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# ======== Import Librerie Google e Shopify ========
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import shopify

# ======== ActiveCampaign ============
def create_field_value(contact_id, field_id, value):
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
        print(f"[OK] FieldValue creato per contatto {contact_id}, campo {field_id} con valore '{value}'. Nuovo ID: {created_id}", flush=True)
        return created_id
    else:
        print(f"[ERRORE] Creazione fallita per contatto {contact_id}, campo {field_id}: {response.text}", flush=True)
        return None

def get_contact_by_email(email):
    headers = {"Api-Token": AC_API_KEY}
    url = f"{AC_BASE_URL}/api/3/contacts?filters[email]={email}"
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        data = response.json()
        contacts = data.get("contacts", [])
        if contacts:
            return contacts[0]
        else:
            print("Nessun contatto trovato per email:", email, flush=True)
            return None
    else:
        print("Errore nella richiesta del contatto:", response.text, flush=True)
        return None

# ======== Google Sheets con Service Account ============
def init_google_sheets_service():
    json_data = json.loads(SERVICE_ACCOUNT_FILE)
    creds = Credentials.from_service_account_info(json_data, scopes=SCOPES)
    service = build("sheets", "v4", credentials=creds)
    print("Google Sheets service inizializzato (Service Account).", flush=True)
    return service

def read_sheet(service):
    sheet = service.spreadsheets()
    result = sheet.values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=SHEET_NAME
    ).execute()
    return result.get("values", [])

def insert_row_to_sheet(service, values_list):
    current_data = read_sheet(service)
    next_row = len(current_data) + 1
    range_to_update = f"{SHEET_NAME}!A{next_row}:K{next_row}"
    body = {"values": [values_list]}
    result = service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=range_to_update,
        valueInputOption="USER_ENTERED",
        body=body
    ).execute()
    print(f"Riga inserita in {range_to_update}: {values_list}", flush=True)
    return result

def cerca_data_primo_ordine(service, email):
    rows = read_sheet(service)
    for row in rows:
        if len(row) >= 3 and row[2].strip().lower() == email.strip().lower() and row[1] == "1":
            return row[0]
    return "non trovata"

# ======== Funzione per aggiornare il numero di telefono del contatto su AC ========
def update_contact_phone(contact_id, phone):
    headers = {"Api-Token": AC_API_KEY, "Content-Type": "application/json"}
    url = f"{AC_BASE_URL}/api/3/contacts/{contact_id}"
    payload = {
        "contact": {
            "phone": phone
        }
    }
    response = requests.put(url, headers=headers, json=payload)
    if response.status_code in [200, 201]:
        print(f"[OK] Contatto {contact_id} aggiornato con numero di telefono {phone}.", flush=True)
        return True
    else:
        print(f"[ERRORE] Aggiornamento telefono fallito per contatto {contact_id}: {response.text}", flush=True)
        return False

# ======== Shopify ============
def init_shopify_session():
    session = shopify.Session(SHOP_URL, API_VERSION, ACCESS_TOKEN)
    shopify.ShopifyResource.activate_session(session)
    print("Shopify session inizializzata.", flush=True)

def get_last_week_range():
    now = datetime.datetime.now()
    today_date = now.date()
    wednesday_date = today_date - datetime.timedelta(days=6)
    start = datetime.datetime(wednesday_date.year, wednesday_date.month, wednesday_date.day, 0, 0, 0)
    end = datetime.datetime(today_date.year, today_date.month, today_date.day, 23, 59, 59)
    created_at_min = start.isoformat() + "Z"
    created_at_max = end.isoformat() + "Z"
    return created_at_min, created_at_max

def get_orders_in_range():
    created_at_min, created_at_max = get_last_week_range()
    orders = shopify.Order.find(
        created_at_min=created_at_min,
        created_at_max=created_at_max,
        status='any',
        limit=250
    )
    orders.reverse()
    print(f"Trovati {len(orders)} ordini da {created_at_min} a {created_at_max}.", flush=True)
    return orders

def extract_order_info(order):
    dataAcquisto = order.created_at
    email = getattr(order, 'email', None)
    if not email and hasattr(order, 'customer') and order.customer:
        email = order.customer.email
    # Estrae telefono: cerca in order.phone o in order.customer.phone
    phone = getattr(order, 'phone', None)
    if not phone and hasattr(order, 'customer') and order.customer:
        phone = getattr(order.customer, 'phone', None)

    landing = getattr(order, 'landing_site', "") or ""
    landing_lower = ""
    campagnaDiProvenienza = ""
    if landing and "utm_campaign=" in landing:
        try:
            campagnaDiProvenienza = landing.split("utm_campaign=")[1].split("&")[0]
            landing_lower = landing.lower()
        except:
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
        "phone": phone,
        "campagnaDiProvenienza": campagnaDiProvenienza,
        "canale_di_provenienza": canale_di_provenienza,
        "totaleOrdine": totaleOrdine,
        "nomeProdotto": nomeProdotto
    }

# ======== LOGICA PRINCIPALE ============
def run_script():
    """
    Esegue la logica (Shopify -> Google Sheets -> ActiveCampaign).
    Per ogni ordine dell’ultima settimana:
      - Aggiorna il Google Sheet
      - Crea i fieldValue in ActiveCampaign per Data (primo, ultimo, secondo acquisto)
      - Aggiorna il numero di telefono del contatto con quello preso da Shopify (se presente)
    """
    init_shopify_session()
    gs_service = init_google_sheets_service()

    orders = get_orders_in_range()
    for order in orders:
        time.sleep(4)
        if not hasattr(order, "customer") or order.customer is None:
            continue

        orders_count = getattr(order.customer, "orders_count", 0)
        if orders_count == 0 and hasattr(order.customer, "id"):
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
            print(f"Contatto ActiveCampaign non trovato per {order_info['email']}", flush=True)
            continue

        ac_contact_id = ac_contact.get("id")

        # Aggiornamento del numero di telefono su ActiveCampaign (se presente)
        if order_info.get("phone"):
            update_contact_phone(ac_contact_id, order_info["phone"])

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
            print(f"FieldValue creato: id {id_field39} (Data primo acquisto), id {id_field38} (Data ultimo acquisto).", flush=True)

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
            print(f"FieldValue creato: id {id_field38} (Data ultimo acquisto), id {id_field40} (Data secondo acquisto).", flush=True)

    print("Processo completato.", flush=True)

def check_run_script():
    """
    Ogni minuto, controlliamo se è martedì 23:59 in ora italiana,
    e se sì, eseguiamo run_script().
    """
    rome_now = datetime.datetime.now(ZoneInfo("Europe/Rome"))
    # Selezioniamo MARTEDÌ: in Python weekday() restituisce: 0 lunedì, 1 martedì, etc.
    if rome_now.weekday() == 2:  
        # Esempio di condizione: per testing ho messo un orario fittizio (es. 16:15)
        if rome_now.hour == 16 and rome_now.minute == 48:
            print("** E’ martedì 23:59 in Italia! Eseguo run_script() **", flush=True)
            run_script()
        else:
            print(f"Ora Roma {rome_now}, non è martedì 23:59, skip.", flush=True)
    else:
        print(f"Ora Roma {rome_now}, non è martedì, skip.", flush=True)

def main():
    schedule.every(1).minutes.do(check_run_script)
    print("Scheduler avviato (controllo ogni minuto l’ora in Europe/Rome).", flush=True)

    while True:
        schedule.run_pending()
        print("A", flush=True)
        time.sleep(10)

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print("Si è verificato un errore:", e, flush=True)
        sys.exit(1)
