import gspread
from oauth2client.service_account import ServiceAccountCredentials

scope = ['https://drive.google.com/drive/my-drive']


credentials = ServiceAccountCredentials.from_json_keyfile_name('C:\dev\My Project-sample.json', scope)

gc = gspread.authorize(credentials).open("Google Sheet Name")

wks = gc.get_worksheet(0)

gc.get_worksheet(-1)## integer position으로 접근
gc.worksheet('Sheet Name) ## sheet name으로 접근

wks.update_acell('D1', 'It's Work!')

# Select a range
cell_list = wks.range('A1:C7')

for cell in cell_list:
    cell.value = 'test'

# Update in batch
wks.update_cells(cell_list)