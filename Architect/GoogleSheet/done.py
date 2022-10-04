import os.path

import gspread
from google.oauth2.service_account import Credentials






scopes = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive',
    'https://spreadsheets.google.com/feeds'
]
credentials = Credentials.from_service_account_file(
    'C:\\Users\\ckddn\\PycharmProjects\\c21166a86163.json',
    scopes=scopes
)
gc = gspread.authorize(credentials)
spreadsheet_url = 'https://docs.google.com/spreadsheets/d/138QjV6skBObCz1TaWOwgtBLdIgYTnCK9nRZtNMcFCyU/edit#gid=396813348'
doc = gc.open_by_url(spreadsheet_url)



before_worksheet = doc.worksheet('원주(전)')

# after_worksheet = doc.add_worksheet(title="결과", rows="100", cols="6")

items = []
worksheetvalue = before_worksheet.get_all_values()


for row in worksheetvalue:
    items.append(worksheetvalue)

print(items)

# for row in worksheetvalue:
#     try:
#         if ( row[0].value is not None
#         ):
#             temp_windows_name = f"{row[0].value} | {float(row[1].value)}*{float(row[2].value)} | M2"
#         if (row[15].value is not None
#         ):
#             temp_windows_name = f"{row[0].value}(시스템도어포함) | {float(row[1].value)}*{float(row[2].value)} | M2"
#     except:
#         print('창호이름오류')




# worksheetvalue = worksheet.get_all_values()
# items = []
# for row in worksheetvalue:
#     items.append(row)

