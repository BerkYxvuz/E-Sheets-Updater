import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
from pathlib import Path

key = {
  // Your Key
}

# Kimlik doğrulama kapsamlarını ayarlayın
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']

# Kimlik doğrulama bilgilerini yükle
credentials = ServiceAccountCredentials.from_json_keyfile_dict(key, scope)

# Google Sheets'e bağlan
gc = gspread.authorize(credentials)

# Dosya adı ve sayfa adını belirtin
excel_file = Path(input("Excel dosyasının yolunu giriniz: ").replace('"',''))
gs = 'Genel Stok'
od = 'Okut Depo'

# Belirli bir sayfayı oku
df_gs = pd.read_excel(excel_file, sheet_name=gs)
df_gs.fillna('', inplace=True)

df_od = pd.read_excel(excel_file, sheet_name=od)
df_od.fillna('', inplace=True)

# Google Sheets dosyasını açın
sh = gc.open("Yeni Depo & Refill")

# Sayfa verilerini temizle
worksheet = sh.worksheet(gs)
worksheet.clear()

# Sayfa verilerini temizle
worksheet = sh.worksheet(od)
worksheet.clear()

# Verileri string olarak Google Sheets'e yükle
values = df_gs.astype(str).values.tolist()
sh.values_update(
    gs, 
    params={'valueInputOption': 'RAW'}, 
    body={'values': values}
)

print("Genel Stok Güncellendi.")

# Verileri string olarak Google Sheets'e yükle
values = df_od.astype(str).values.tolist()
sh.values_update(
    od, 
    params={'valueInputOption': 'RAW'}, 
    body={'values': values}
)
print("Okut Depo Güncellendi.")
