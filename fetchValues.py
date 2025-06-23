import pandas as pd
from spreadsheetTemplate import columns

import re

from datetime import datetime
import pytz

igTransactions = pd.read_csv('transactions.csv', index_col=False)
igTransactions.reset_index(drop=True, inplace=True)

portfolioFile = pd.read_excel('portfolio.xlsx')
pRowIndex = len(portfolioFile)


def enterVal(fieldName,value):
    portfolioFile.at[pRowIndex + transaction, fieldName] = value

noOfRows = igTransactions.shape[0]

for transaction in range(noOfRows):
    qty = 1
    currRow = igTransactions.loc[transaction]

    description = currRow[2]

    transactNameAddr = re.search(r'^(.*?)\s*-\s*',description)
    if transactNameAddr == None:
        transactNameAddr = re.search(r'^(.*?)\s*\(',description)
    transactName = transactNameAddr.group(1)

    dateTimeUTC = currRow[13]

    dateTimeUTC_noT = dateTimeUTC.replace('T', ' ')

    formatted_DTUtc = datetime.strptime(dateTimeUTC_noT, "%Y-%m-%d %H:%M:%S")

    utcDatetimeLocal = pytz.utc.localize(formatted_DTUtc)

    ukTimeWithTZ = utcDatetimeLocal.astimezone(pytz.timezone("Europe/London"))

    ukTime = ukTimeWithTZ.replace(tzinfo=None)

    dateTimeUtcMatch = re.search(r'T(.*)', dateTimeUTC)
    dateTimeUtc = dateTimeUtcMatch.group(1) if dateTimeUtcMatch else ""


    if transactName == "Cash Interest Paid To Client":
        transactType = "CASH"
    else:
        transactType = "STOCKS"

    transactAction = currRow[1]

    inOrOut = currRow[5]

    if transactAction == "Dividend" or transactAction == "Client Consideration":
        qtyAddr = re.search(r'(\d+(?:\.\d+)?)@',description)
        if qtyAddr != None:
            qty = qtyAddr.group(1)
    
    if transactAction == "Dividend":
        transactAction = "DIVIDEND"
    elif transactAction == "Client Consideration":
        transactAction = "BUY/SELL"
        if inOrOut == "DEPO":
            transactAction = "BUY"
        elif inOrOut == "WITH":
            transactAction = "SELL"
    else:
        transactAction = "DEPOSIT/WITHRAWAL/TRANSFER"

    gbp_pAndL = currRow[11]
    amtPerUnit = float(gbp_pAndL) / float(qty)
    if float(gbp_pAndL) < 0:
        amtPerUnit *= -1

    fieldsInRow = {
        "DATE": currRow[0], #date
        #"period": currRow[3], #period
        "GBP_pAndL": currRow[4], #GBP_pAndL
        #"curr": currRow[10], #curr
        "ACCOUNT TYPE": "Regular Stocks & Shares",
        "ACTION": "BUY/SELL/DIVIDEND/DEPOSIT/WITHRAWAL/TRANSFER",
        "PLATFORM": "IG Trading",
        #Useful fields
        "NOTES": currRow[1], #summary
        "AMOUNT PER UNIT": amtPerUnit,
        "FEES": "N/D",

        "NAME": transactName, #description

        "ACTION": transactAction,

        "QUANTITY": qty,

        "inOrOut": currRow[5], #inOrOut
        "TYPE": transactType,
        "uniqueRef": currRow[6], #uniqueRef
        "openLevel": currRow[7], #openLevel
        "closeLevel": currRow[8], #closeLevel
        "size": currRow[9], #size
        "PROFIT/LOSS (APPROX)": gbp_pAndL, #pAndLAmt
        "TIME (UTC)": dateTimeUtc, #dateUTC
        "TIME (UK)": ukTime,
        "opendateUTC": currRow[14], #opendateUTC
        "CURRENCY": currRow[15] #currCode
    }

    for col in columns:
        if col in fieldsInRow:
            enterVal(col,fieldsInRow[col])
        else:
            enterVal(col,"")
    portfolioFile.to_excel("portfolio.xlsx", index=False)