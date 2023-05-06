from urllib.request import urlopen, Request
from bs4 import BeautifulSoup
import openpyxl as xl
from openpyxl.styles import Font
import keys
from twilio.rest import Client

url = 'https://finance.yahoo.com/crypto/'

headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 6.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/41.0.2228.0 Safari/537.3'}
req = Request(url, headers=headers)

webpage = urlopen(req).read()

soup = BeautifulSoup(webpage,'html.parser')
print(soup.title.text)

client = Client(keys.accountSID, keys.authToken)

MyCellNum = "+16613901505"
TwilioNum = "+15074163593"

tables = soup.findAll('table')
updated_tables = tables[0]
rows = updated_tables.findAll('tr')
wb = xl.Workbook()
ws = wb.active
ws.title = "Crypto Dashboard"

ws['B1'] = "Symbol"
ws['C1'] = "Currency Name"
ws['D1'] = "Current Price"
ws['E1'] = "Percent Change 24hrs"
ws['F1'] = "Price t-1"

ws.column_dimensions['B2'].width = 30
ws.column_dimensions['C2'].width = 30
ws.column_dimensions['D2'].width = 30
ws.column_dimensions['E2'].width = 30
ws.column_dimensions['F2'].width = 30


HeaderFont = Font(name='Verdana',size=16,bold=True)

ws['B2'].font = HeaderFont
ws['C2'].font = HeaderFont
ws['D2'].font = HeaderFont
ws['E2'].font = HeaderFont
ws['F2'].font = HeaderFont

cell_range = ws[1:6]
for row in cell_range:
    for cell in row:
        cell.font = Font(name='Verdana')



for row in range(1,6):
    td = rows[row].findAll('td')
    symbol = td[1].text
    curr_name = td[2].text
    price = float(td[3].text.replace(",", "").replace("$", ""))
    percent_change = float(td[5].text.replace("%","").replace("B",""))
    price_tminusone = round(price*(1+percent_change),2)
    change_detector = price - price_tminusone

    if change_detector >= 5 or change_detector <= -5:
        message = "A price change of " + str(change_detector) + " has occurred."
        alert = client.messages.create(to=MyCellNum, from_=TwilioNum, body=message)

    ws['B'+str(row+1)] = symbol
    ws['C'+str(row+1)] = curr_name
    ws['D'+str(row+1)] = '$' + str(format(price,',.2f'))
    ws['E'+str(row+1)] = str(format(percent_change, ',.2f')+'%') 
    ws['F'+str(row+1)] = '$' + format(price_tminusone,',.2f')

wb.save('CryptoDashboard.xlsx')





