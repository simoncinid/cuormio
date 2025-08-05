#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import sys
import time
import schedule
import datetime
from zoneinfo import ZoneInfo
import shopify
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import requests
import json

# Configurazione Shopify
SHOP_URL = "cuormio.myshopify.com"
ACCESS_TOKEN = os.environ.get('SHOPIFY_ACCESS_TOKEN')

# Configurazione Google Sheets
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SERVICE_ACCOUNT_FILE = 'client_secret.json'
SPREADSHEET_ID = '1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms'

def init_shopify_session():
    shopify.ShopifyResource.set_site(f"https://{ACCESS_TOKEN}@{SHOP_URL}/admin/api/2023-10")
    shopify.ShopifyResource.set_headers({'X-Shopify-Access-Token': ACCESS_TOKEN})

def init_google_sheets_service():
    creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    service = build('sheets', 'v4', credentials=creds)
    return service

def get_orders_in_range():
    start_date, end_date = get_last_week_range()
    orders = shopify.Order.find(
        created_at_min=start_date,
        created_at_max=end_date,
        status='any',
        limit=250
    )
    print(f"[INFO] Trovati {len(orders)} ordini tra {start_date} e {end_date}", flush=True)
    return orders

def get_last_week_range():
    now = datetime.datetime.now(ZoneInfo('Europe/Rome'))
    start_of_week = now - datetime.timedelta(days=now.weekday() + 7)
    end_of_week = start_of_week + datetime.timedelta(days=6)
    
    start_date = start_of_week.strftime('%Y-%m-%d')
    end_date = end_of_week.strftime('%Y-%m-%d')
    
    return start_date, end_date

def extract_order_info(order):
    email = order.email or (order.customer.email if order.customer else None)
    phone = (
        getattr(order, 'phone', None)
        or (order.customer.phone if order.customer else None)
        or (getattr(order, 'shipping_address', None) and order.shipping_address.phone)
        or (getattr(order, 'billing_address', None) and order.billing_address.phone)
    )
    
    # Normalizza il telefono
    if phone:
        phone = phone.replace(' ', '').replace('-', '').replace('(', '').replace(')', '')
        if phone.startswith('+39'):
            phone = phone[3:]
        if phone.startswith('39'):
            phone = phone[2:]
        if phone.startswith('0'):
            phone = phone[1:]
    
    return {
        'created_at': order.created_at,
        'email': email,
        'phone': phone,
        'total_price': order.total_price,
        'order_id': order.id
    }

def insert_row_to_sheet(service, values_list):
    try:
        range_name = 'Sheet1!A:K'
        body = {
            'values': [values_list]
        }
        result = service.spreadsheets().values().append(
            spreadsheetId=SPREADSHEET_ID,
            range=range_name,
            valueInputOption='RAW',
            insertDataOption='INSERT_ROWS',
            body=body
        ).execute()
        print(f"[INFO] Riga inserita in Google Sheets", flush=True)
        return True
    except Exception as e:
        print(f"[ERRORE] Errore inserimento Google Sheets: {e}", flush=True)
        return False

def run_script():
    init_shopify_session()
    gs_service = init_google_sheets_service()
    orders = get_orders_in_range()

    for order in orders:
        time.sleep(4)
        if not getattr(order, 'customer', None):
            continue

        info = extract_order_info(order)
        print(f"[INFO] Order {order.id}: email={info['email']}, phone={info['phone']}", flush=True)

        date_obj = datetime.datetime.fromisoformat(info['created_at'].replace('Z', '+00:00'))
        formatted_date = date_obj.strftime("%d/%m/%Y")

        tot = f"€ {info['total_price']}"
        tot = tot.replace('.', ',')
        
        row = [formatted_date, info['email'], info['phone'], tot, info['order_id']]
        insert_row_to_sheet(gs_service, row)

    print("[OK] Processo completato.", flush=True)

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