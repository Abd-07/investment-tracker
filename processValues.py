import pandas as pd

#read csv file
igTransactions = pd.read_csv('transactions.csv', index_col=False)
igTransactions.reset_index(drop=True, inplace=True)

sectors = [
    "Energy",
    "Materials",
    "Industrials",
    "Consumer Discretionary",
    "Consumer Staples",
    "Consumer Staples",
    "Healthcare",
    "Financials",
    "Information Technology (IT)",
    "Communication Services",
    "Utilities",
    "Real Estate",
    "Consumer Technology"
]