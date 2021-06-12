import shopify
import json
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials
from googleapiclient import discovery

# Get raw data from the Shopify webpage
API_KEY = "*********"
PASSWORD = "*********"
API_VERSION = "2021-04"
SHOP_NAME = '************'

shopify.ShopifyResource.set_user(API_KEY)
shopify.ShopifyResource.set_password(PASSWORD)

SHOP_URL = "https://%s:%s@%s.myshopify.com/admin/api/%s" % (API_KEY, PASSWORD, SHOP_NAME, API_VERSION)
shopify.ShopifyResource.set_site(SHOP_URL)

page = shopify.Order.find()
product_list = []
product_list += page

while page.has_next_page():
    page = page.next_page()
    product_list = product_list + page

# Convert every single product's data to the appropriate format
all_rows = []
dict_data = dict()
for product in product_list:
    json_data = product.to_json()
    data = json.loads(json_data)
    keys_list = []
    values_list = []
    for key, value in data['order'].items():
        keys_list.append(key)
        values_list.append(value)
    dict_data['ShopOrderNumber'] = values_list[keys_list.index('name')]
    dict_data['CreatedAt'] = values_list[keys_list.index('created_at')]
    dict_data['Title'] = values_list[keys_list.index('line_items')][0]['title']
    dict_data['Variant_Title'] = values_list[keys_list.index('line_items')][0]['variant_title']
    dict_data['tracking'] = values_list[keys_list.index('landing_site_ref')]
    dict_data['Quantity'] = values_list[keys_list.index('line_items')][0]['quantity']
    dict_data['Full_Name'] = values_list[keys_list.index('shipping_address')]['name']
    dict_data['Address 1&2'] = values_list[keys_list.index('shipping_address')].get('address1', 'address2')
    dict_data['City'] = values_list[keys_list.index('shipping_address')]['city']
    dict_data['Province'] = values_list[keys_list.index('shipping_address')]['province']
    dict_data['ZIP'] = values_list[keys_list.index('shipping_address')]['zip']
    dict_data['CountryCode'] = values_list[keys_list.index('shipping_address')]['country_code']
    dict_data['Phone'] = values_list[keys_list.index('shipping_address')]['phone']
    dict_data['Email'] = values_list[keys_list.index('contact_email')]
    if values_list[keys_list.index('note_attributes')]:
        if values_list[keys_list.index('note_attributes')][0]['value'][0] == "{":
            dictionary = eval(values_list[keys_list.index('note_attributes')][0]['value'])
            dict_data['url'] = dictionary["thumb"]
        else:
            dict_data['url'] = values_list[keys_list.index('note_attributes')][0]['value']
    else:
        dict_data['url'] = eval(values_list[keys_list.index('line_items')][0]['properties'][-1]['value'])[0]["url"]
    dict_data['Note'] = values_list[keys_list.index('note')]
    dict_data['FinancialStatus'] = values_list[keys_list.index('financial_status')]
    dict_data['PaymentGatewayNames'] = str(values_list[keys_list.index('payment_gateway_names')])
    all_rows.append(list(dict_data.values()))

# Create DataFrame from with dictionary keys and values are columns and row respectively
columns = list(dict_data.keys())
df = pd.DataFrame(all_rows, columns=columns)

# Automately upload data to Google Sheet
SCOPE = ["https://www.googleapis.com/auth/spreadsheets"]

creds = ServiceAccountCredentials.from_json_keyfile_name("creds.json", SCOPE)

service = discovery.build("sheets", "v4", credentials=creds)

SPREADSHEET_URL = "***************************************"

sheet = service.spreadsheets()

# Clear the sheet and update it again
sheet.values().clear(spreadsheetId=SPREADSHEET_URL, range="sheet1!A2:R").execute()

sheet.values().update(spreadsheetId=SPREADSHEET_URL, range="sheet1!A2", valueInputOption="USER_ENTERED",
                      body={"values": all_rows}).execute()
