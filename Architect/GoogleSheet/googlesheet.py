import done
from oauth2client.service_account import ServiceAccountCredentials

scope = ['https://www.googleapis.com/auth/spreadsheets']
creds = ServiceAccountCredentials.from_json_keyfile_name('.json', scope)
client = gspread.authorize(creds)

doc = client.open_by_url('https://docs.google.com/spreadsheets/d/138QjV6skBObCz1TaWOwgtBLdIgYTnCK9nRZtNMcFCyU/edit#gid=0')

sheet1 = doc.worksheet('원주1')

cnt = int(sheet1.cell(1,2).value)
print('기존 행수 : ', cnt)
