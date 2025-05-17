import os
import datetime
import json
import sys
import time
import re
import requests
from dateutil import parser
from zoneinfo import ZoneInfo
import schedule

# ======== Variabili d'Ambiente e Configurazioni Principali ========
SHOP_URL     = os.environ.get("SHOP_URL", "")
API_VERSION  = os.environ.get("API_VERSION", "")   # es. "2023-04"
ACCESS_TOKEN = os.environ.get("ACCESS_TOKEN", "")

AC_BASE_URL = os.environ.get("AC_BASE_URL", "")
AC_API_KEY  = os.environ.get("AC_API_KEY", "")

SPREADSHEET_ID = "1vqX3vOoQgIeJu9nSwLw11Y3UU_-YTFVu_V8wwHLUvkA"
SHEET_NAME     = os.environ.get("SHEET_NAME", "")  # Nome del foglio

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

# ======== ActiveCampaign ========

def create_field_value(contact_id, field_id, value):
    headers = {"Api-Token": AC_API_KEY, "Content-Type": "application/json"}
    url = f"{AC_BASE_URL}/api/3/fieldValues"
    payload = {"fieldValue": {"contact": contact_id, "field": field_id, "value": value}}
    response = requests.post(url, headers=headers, json=payload)
    if response.status_code in (200, 201):
        created_id = response.json().get("fieldValue", {}).get("id")
        print(f"[OK] FieldValue creato: contatto={contact_id}, campo={field_id}, valore={value} -> id {created_id}", flush=True)
        return created_id
    else:
        print(f"[ERRORE] create_field_value: {response.text}", flush=True)
        return None

def get_contact_by_email(email):
    headers = {"Api-Token": AC_API_KEY}
    url = f"{AC_BASE_URL}/api/3/contacts?filters[email]={email}"
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        contacts = response.json().get("contacts", [])
        if contacts:
            return contacts[0]
        print(f"[INFO] Nessun contatto trovato per email: {email}", flush=True)
        return None
    print(f"[ERRORE] get_contact_by_email: {response.text}", flush=True)
    return None

def add_contact_to_automation(contact_id, automation_id):
    headers = {"Api-Token": AC_API_KEY, "Content-Type": "application/json"}
    url = f"{AC_BASE_URL}/api/3/contactAutomations"
    payload = {
        "contactAutomation": {
            "contact": contact_id,
            "automation": automation_id
        }
    }
    response = requests.post(url, headers=headers, json=payload)
    if response.status_code in (200, 201):
        print(f"[OK] Aggiunto a automazione ID {automation_id}", flush=True)
        return True
    else:
        print(f"[ERRORE] add_contact_to_automation: {response.text}", flush=True)
        return False

def automation_is_active_for_contact(contact_id, automation_id):
    headers = {"Api-Token": AC_API_KEY}
    url = (f"{AC_BASE_URL}/api/3/contactAutomations"
           f"?filters[automation]={automation_id}&filters[contact]={contact_id}&limit=100")
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        automations = response.json().get("contactAutomations", [])
        return len(automations) > 0
    print(f"[ERRORE] automation_is_active_for_contact: {response.text}", flush=True)
    return False

# ======== Google Sheets ========

def init_google_sheets_service():
    data = json.loads(SERVICE_ACCOUNT_FILE)
    creds = Credentials.from_service_account_info(data, scopes=SCOPES)
    service = build("sheets", "v4", credentials=creds)
    print("[OK] Google Sheets service inizializzato.", flush=True)
    return service

def read_sheet(service):
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=SHEET_NAME).execute()
    return result.get("values", [])

def insert_row_to_sheet(service, values_list):
    current = read_sheet(service)
    next_row = len(current) + 1
    range_to_update = f"{SHEET_NAME}!A{next_row}:K{next_row}"
    body = {"values": [values_list]}
    service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=range_to_update,
        valueInputOption="USER_ENTERED",
        body=body
    ).execute()
    print(f"[OK] Riga inserita {range_to_update}: {values_list}", flush=True)

def cerca_data_primo_ordine(service, email):
    rows = read_sheet(service)
    for row in rows:
        if len(row) >= 3 and row[2].strip().lower() == email.strip().lower() and row[1] == "1":
            return row[0]
    return "non trovata"

# ======== Normalize Phone ========

def normalize_phone(raw_phone):
    digits = re.sub(r"\D+", "", raw_phone or "")
    if not digits:
        return None
    digits = digits.lstrip('0') if digits.startswith('00') else digits
    if not digits.startswith("39"):
        digits = "39" + digits
    if not digits.startswith("+"):
        digits = "+" + digits
    return digits

# ======== ActiveCampaign: update phone ========

def update_contact_phone(contact_id, phone):
    headers = {"Api-Token": AC_API_KEY, "Content-Type": "application/json"}
    url = f"{AC_BASE_URL}/api/3/contacts/{contact_id}"
    payload = {"contact": {"phone": phone}}
    response = requests.put(url, headers=headers, json=payload)
    if response.status_code in (200, 201):
        print(f"[OK] Contatto {contact_id} aggiornato con telefono {phone}.", flush=True)
        return True
    print(f"[ERRORE] update_contact_phone: {response.text}", flush=True)
    return False

# ======== Shopify ========

def init_shopify_session():
    session = shopify.Session(SHOP_URL, API_VERSION, ACCESS_TOKEN)
    shopify.ShopifyResource.activate_session(session)
    print("[OK] Shopify session inizializzata.", flush=True)

def get_last_week_range():
    today = datetime.datetime.now().date()
    start = datetime.datetime.combine(today - datetime.timedelta(days=7), datetime.time(0,0))
    end = datetime.datetime.combine(today - datetime.timedelta(days=1), datetime.time(23,59,59))
    return start.isoformat() + "Z", end.isoformat() + "Z"

def get_orders_in_range():
    created_at_min, created_at_max = get_last_week_range()
    orders = shopify.Order.find(
        created_at_min=created_at_min,
        created_at_max=created_at_max,
        status='any',
        limit=250
    )
    orders.reverse()
    print(f"[OK] Trovati {len(orders)} ordini da {created_at_min} a {created_at_max}.", flush=True)
    return orders

def extract_order_info(order):
    email = order.email or (order.customer.email if order.customer else None)
    phone = (
        getattr(order, 'phone', None)
        or (order.customer.phone if order.customer else None)
        or (getattr(order, 'shipping_address', None) and order.shipping_address.phone)
        or (getattr(order, 'billing_address', None)  and order.billing_address.phone)
    )
    phone = normalize_phone(phone)
    landing = getattr(order, 'landing_site', '') or ''
    campagna = ''
    if 'utm_campaign=' in landing:
        campagna = landing.split('utm_campaign=')[1].split('&')[0]
    canale = ''
    l = landing.lower()
    if 'facebook' in l or 'fbclid' in l:
        canale = 'FB'
    elif 'instagram' in l:
        canale = 'IG'
    else:
        canale = getattr(order, 'source_name', '')

    if not order.line_items:
        prodotto = None
    else:
        titles = [li.title.lower() for li in order.line_items]
        has_cofanetto = any("cofanetto degustazione" in t for t in titles)
        num_items     = len(titles)
        if has_cofanetto and num_items == 1:
            prodotto = "degustazione"
        elif has_cofanetto and num_items > 1:
            prodotto = "più cose"
        else:
            prodotto = "primo acquisto no cofanetto"
    return {
        'created_at': order.created_at,
        'email': email,
        'phone': phone,
        'campagna': campagna,
        'canale': canale,
        'total_price': order.total_price,
        'product': prodotto
    }

# ======== Logica Principale ========

def run_script():
    init_shopify_session()
    gs_service = init_google_sheets_service()
    orders = get_orders_in_range()

    for order in orders:
        time.sleep(4)
        if not getattr(order, 'customer', None):
            continue

        orders_count = getattr(order.customer, 'orders_count', 0)
        if orders_count == 0 and getattr(order.customer, 'id', None):
            cust = shopify.Customer.find(order.customer.id)
            orders_count = cust.orders_count

        if orders_count not in [1, 2]:
            continue

        info = extract_order_info(order)
        print(f"[INFO] Order {order.id}: email={info['email']}, phone={info['phone']}", flush=True)

        date_obj = parser.parse(info['created_at'])
        formatted_date = date_obj.strftime("%d/%m/%Y")

        ac_contact = get_contact_by_email(info['email'])
        if not ac_contact:
            continue
        ac_id = ac_contact.get('id')

        if info['phone']:
            update_contact_phone(ac_id, info['phone'])

        tot = f"€ {info['total_price']}"
        tot = tot.replace('.', ',')
        if orders_count == 1:
            row = [formatted_date, '1', info['email'], info['campagna'], info['canale'], None, None, None, None, tot, info['product']]
            insert_row_to_sheet(gs_service, row)
            create_field_value(ac_id, '39', formatted_date)
            create_field_value(ac_id, '38', formatted_date)

            # === INTEGRAZIONE AUTOMAZIONE ACTIVE CAMPAIGN (PRIMO ACQUISTO) ===
            if info['product'] == "primo acquisto no cofanetto":
                automation_id = 52
                automation_label = "Primo Acquisto No Cofanetto Degustazione"
            else:
                automation_id = 32
                automation_label = "Acquisto Cofanetto Degustazione"

            if automation_is_active_for_contact(ac_id, automation_id):
                print(f"[INFO] Automazione '{automation_label}' già attiva (ID {automation_id})", flush=True)
            else:
                if add_contact_to_automation(ac_id, automation_id):
                    time.sleep(2)
                    if automation_is_active_for_contact(ac_id, automation_id):
                        print(f"[INFO] Automazione '{automation_label}' aggiunta e ora attiva (ID {automation_id})", flush=True)
                    else:
                        print(f"[ERRORE] Automazione '{automation_label}' NON trovata dopo aggiunta (ID {automation_id})", flush=True)
                else:
                    print(f"[ERRORE] Fallita aggiunta automazione '{automation_label}' (ID {automation_id})", flush=True)

        else:
            prima = cerca_data_primo_ordine(gs_service, info['email'])
            row = [formatted_date, '2', info['email'], info['campagna'], info['canale'], None, None, prima, None, tot, info['product']]
            insert_row_to_sheet(gs_service, row)
            create_field_value(ac_id, '38', formatted_date)
            create_field_value(ac_id, '40', formatted_date)

    print("[OK] Processo completato.", flush=True)

# ======== Scheduler ========

def check_run_script():
    now_rome = datetime.datetime.now(ZoneInfo('Europe/Rome'))
    if now_rome.weekday() == 4 and now_rome.hour == 6 and now_rome.minute == 0:
        print("[SCHED] Eseguo run_script() alle 6:00 di venerdì", flush=True)
        run_script()
    else:
        print(f"[SCHED] Ora Roma {now_rome}, skip.", flush=True)

def main():
    schedule.every(1).minutes.do(check_run_script)
    print("[OK] Scheduler avviato.", flush=True)
    while True:
        schedule.run_pending()
        time.sleep(10)

if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        print("[CRITICAL] Errore:", e, flush=True)
        sys.exit(1)
