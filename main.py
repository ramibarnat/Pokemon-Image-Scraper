from bs4 import BeautifulSoup
import requests
from xlutils.copy import copy
from xlrd import open_workbook

# w = copy('book1.xls')
# w.get_sheet(0).write(0,0,"foo")
# w.save('book2.xls')
#1010
old_wb = open_workbook('Num_and_image.xls')
old_sheet = old_wb.sheet_by_index(0)
new_wb = copy(old_wb)
new_sheet = new_wb.get_sheet(0)
for i in range(old_sheet.nrows):
    req = requests.get('https://pokemon.fandom.com/wiki/' + str(old_sheet.cell(i, 1).value)).text
    soup = BeautifulSoup(req,'html.parser')
    link = soup.find('img')
    new_sheet.write(i, 2, list(enumerate(soup.find_all('img')))[3][1]["src"])

new_wb.save('test.xls')
