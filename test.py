import gspread
import calendar

def init_sheets(auth, filename):
    gc = gspread.service_account(filename=auth)
    return gc.open(filename)


ss = init_sheets("test.json", "test")
#sht = ss.add_worksheet(title="10월", rows="70", cols="20")
sht = ss.worksheet("10월")
selected_cell = "B6:C9"
data = [[1, 2], [2, 3], [3, 4], [4, 5]]
sht.update(selected_cell, data)
