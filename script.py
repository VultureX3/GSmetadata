import gspread
from oauth2client.service_account import ServiceAccountCredentials

scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']

credentials = ServiceAccountCredentials.from_json_keyfile_name('GSmetadata-b4b4a0eaa7d3.json', scope)
gc = gspread.authorize(credentials)

metadata = gc.open_by_key('1eCVqx5GhgY_LXcYcLIcC0_2Rw4cAi8Elyx7i9o3l5ko').sheet1

print(metadata.get_all_records())
