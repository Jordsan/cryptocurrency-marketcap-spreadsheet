import json
import urllib.request
import xlsxwriter
import time

# sorting callback function for json data
def position_sort(json):
    try:
        return int(json['position'])
    except KeyError:
        return -1

urlData = "https://api.coinmarketcap.com/v1/ticker/?limit=100"
webURL = urllib.request.urlopen(urlData)
data = webURL.read()
encoding = webURL.info().get_content_charset('utf-8')
JSON_Data = json.loads(data.decode(encoding))


# Extract relevant info, then sort data in object to be lowest pos first
JSON_Data.sort(key=position_sort)
dataArray = []
dataArray2 = []
dataArray3 = []
for index, item in enumerate(JSON_Data):
    dataArray2.append(item["percent_change_24h"])
    dataArray3.append(item["24h_volume_usd"])

maxCharsPos = 0
maxCharsName = 0
maxCharsSymbol = 0
maxCharsMCUSD = 0
maxCharsPriceUSD = 0
maxCharsPriceBTC = 0
maxCharsAvailSup = 0
maxCharsVol24HrUSD = 0
maxCharsVol1Hr = 0
maxCharsVol24Hr = 0
maxCharsVol7Day = 0

for index, item in enumerate(JSON_Data):

    # Getting data for top 100
    if index + 1 > 100:
        break
    itemPosition = item.get("rank")
    itemName = item.get("name")
    itemSymbol = item.get("symbol")
    itemMarketCapUSD = item["market_cap_usd"]
    itemPriceUSD = item["price_usd"]
    itemPriceBTC = item["price_btc"]
    itemAvailableSupply = item.get("available_supply")
    itemChange1Hr = item["percent_change_1h"]
    itemChange7Day = item["percent_change_7d"]

    # Calculating column widths
    maxCharsPos = max(len(itemPosition) + 1, maxCharsPos)
    maxCharsName = max(len(itemName) + 1, maxCharsName)
    maxCharsSymbol = max(len(itemSymbol) + 1, maxCharsSymbol)
    maxCharsMCUSD = max(len(itemMarketCapUSD) + 1, maxCharsMCUSD)
    maxCharsPriceUSD = max(len(itemPriceUSD) + 1, maxCharsPriceUSD)
    maxCharsPriceBTC = max(len(itemPriceBTC) + 1, maxCharsPriceBTC)
    maxCharsAvailSup = max(len(str(itemAvailableSupply)) + 1, maxCharsAvailSup)
    maxCharsVol24HrUSD = max(len(dataArray3[index]) + 1, maxCharsVol24HrUSD)
    maxCharsVol1Hr = max(len(itemChange1Hr) + 1, maxCharsVol1Hr)
    maxCharsVol7Day = max(len(itemChange7Day) + 1, maxCharsVol7Day)

    # Storing retrieved data in array
    dataArray.append([itemPosition, itemName, itemSymbol, itemMarketCapUSD,
                      itemPriceUSD, itemPriceBTC, itemAvailableSupply, dataArray3[index],
                      itemChange1Hr, dataArray2[index], itemChange7Day])

# Max char lengths for columns (basically Auto Fit)
maxCharsPos = max(len("#") + 1, maxCharsPos)
maxCharsName = max(len("Name") + 1, maxCharsName)
maxCharsSymbol = max(len("Symbol") + 1, maxCharsSymbol)
maxCharsMCUSD = max(len("Market Cap (USD)") + 1, maxCharsMCUSD)
maxCharsPriceUSD = max(len("Unit Price (USD)") + 1, maxCharsPriceUSD)
maxCharsPriceBTC = max(len("Unit Price (BTC)") + 1, maxCharsPriceBTC)
maxCharsAvailSup = max(len("Available Supply") + 1, maxCharsAvailSup)
maxCharsVol24HrUSD = max(len("Volume 24 Hrs (USD)") + 1, maxCharsVol24HrUSD)
maxCharsVol1Hr = max(len("Change 1 Hr (%)") + 1, maxCharsVol1Hr)
maxCharsVol7Day = max(len("Change 7 Days (%)") + 1, maxCharsVol7Day)

# Create excel worksheet from the data
fileName = time.strftime("%m-%d-%y") + ".xlsx"
workbook = xlsxwriter.Workbook("Sheets/" + fileName)
worksheet = workbook.add_worksheet()
format = workbook.add_format({'bold': True})
format.set_align('center')

# Column Headers
worksheet.write('A1', '#', format)
worksheet.write('B1', 'Name', format)
worksheet.write('C1', 'Symbol', format)
worksheet.write('D1', 'Market Cap (USD)', format)
worksheet.write('E1', 'Unit Price (USD)', format)
worksheet.write('F1', 'Unit Price (BTC)', format)
worksheet.write('G1', 'Available Supply', format)
worksheet.write('H1', 'Volume 24 Hrs (USD)', format)
worksheet.write('I1', 'Change 1 Hr (%)', format)
worksheet.write('J1', 'Change 24 Hrs (%)', format)
worksheet.write('K1', 'Change 7 Days (%)', format)

# Column Width
worksheet.set_column('A:A', maxCharsPos)
worksheet.set_column('B:B', maxCharsName)
worksheet.set_column('C:C', maxCharsSymbol)
worksheet.set_column('D:D', maxCharsMCUSD)
worksheet.set_column('E:E', maxCharsPriceUSD)
worksheet.set_column('F:F', maxCharsPriceBTC)
worksheet.set_column('G:G', maxCharsAvailSup)
worksheet.set_column('H:H', maxCharsVol24HrUSD)
worksheet.set_column('I:I', maxCharsVol1Hr)
worksheet.set_column('J:J', maxCharsVol7Day)
worksheet.set_column('K:K', maxCharsVol7Day)

# Column Number Formatting
format0 = workbook.add_format({'num_format': '##', 'align': 'left'})
format1 = workbook.add_format()
format2 = workbook.add_format()
format3 = workbook.add_format({'num_format': 0x03})
format4 = workbook.add_format({'num_format': 0x04})
format5 = workbook.add_format({'num_format': '###,###,###,##0.00000000'})
format6 = workbook.add_format({'num_format': 0x03})
format7 = workbook.add_format({'num_format': 0x03})
format8 = workbook.add_format({'num_format': '###,##0.00"%"'})
format9 = workbook.add_format({'num_format': '###,##0.00"%"'})
format10 = workbook.add_format({'num_format': '###,##0.00"%"'})

# Populate worksheet with data
for index, item in enumerate(dataArray):
    for subIndex, subItem in enumerate(item):
        if subIndex == 0:
            worksheet.write_number(index + 1, subIndex, float(subItem), format0)
        elif subIndex == 1:
            worksheet.write_string(index + 1, subIndex, subItem)
        elif subIndex == 2:
            worksheet.write_string(index + 1, subIndex, subItem)
        elif subIndex == 3:
            worksheet.write_number(index + 1, subIndex, float(subItem), format3)
        elif subIndex == 4:
            worksheet.write_number(index + 1, subIndex, float(subItem), format4)
        elif subIndex == 5:
            worksheet.write_number(index + 1, subIndex, float(subItem), format5)
        elif subIndex == 6:
            worksheet.write_number(index + 1, subIndex, float(subItem), format6)
        elif subIndex == 7:
            worksheet.write_number(index + 1, subIndex, float(subItem), format7)
        elif subIndex == 8:
            worksheet.write_number(index + 1, subIndex, float(subItem), format8)
        elif subIndex == 9:
            worksheet.write_number(index + 1, subIndex, float(subItem), format9)
        elif subIndex == 10:
            worksheet.write_number(index + 1, subIndex, float(subItem), format10)

workbook.close()