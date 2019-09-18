from openpyxl import load_workbook
from tkinter import Tk
from tkinter.filedialog import askopenfilename

Tk().withdraw()
filename = askopenfilename()

path = filename

wb = load_workbook(path)

ws = wb.get_sheet_by_name('딜리스트')

dealInfo = ws.rows

for row in dealInfo:
    date = row[1].value
    channelName = row[2].value

