import pandas as pd
from spreadsheetTemplate import columns

import re

from datetime import datetime
import pytz

#read csv file
igTransactions = pd.read_csv('transactions.csv', index_col=False)
igTransactions.reset_index(drop=True, inplace=True)

portfolioFile = pd.read_excel('portfolio.xlsx')
pRowIndex = len(portfolioFile)


def enterVal(fieldName,value):
    portfolioFile.at[pRowIndex + transaction, fieldName] = value

noOfRows = igTransactions.shape[0]

rate_USD_GBP = 0.0

#iterate over each transaction in the original sheet
for transaction in range(noOfRows):
    qty = 1.0
    currRow = igTransactions.loc[transaction]

    description = currRow[2]

    transactNameAddr = re.search(r'^(.*?)\s*-\s*',description)
    if transactNameAddr == None:
        transactNameAddr = re.search(r'^(.*?)\s*\(',description)
    elif transactNameAddr != None:
        transactName = transactNameAddr.group(1)
    else:
        transactName = None
        

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

    match_USD_GBP = re.search(r"converted at\s*([0-9]+(?:\.[0-9]+)?)", description, re.IGNORECASE)
    if match_USD_GBP != None:
        rate_USD_GBP = float(match_USD_GBP.group(1))
        rate_USD_GBP = round(rate_USD_GBP,4)

    totalPriceNoFee = float(amtPerUnit) * float(qty)

    ttlFees = 0

    totalConsideration = totalPriceNoFee + ttlFees

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
        "FEES": ttlFees,

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
        "CURRENCY": currRow[15], #currCode

        "USD/GBP": rate_USD_GBP,
        "TOTAL PRICE/COST (WITHOUT FEES)": totalPriceNoFee,
        "TOTAL CONSIDERATION": totalConsideration
    }

    for col in columns:
        if col in fieldsInRow:
            enterVal(col,fieldsInRow[col])
        else:
            enterVal(col,"")
    portfolioFile.to_excel("portfolio.xlsx", index=False)