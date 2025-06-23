import pandas as pd

portfolioFile.at[transaction, "DATE"] = dateUTC

portfolioFile.to_excel("portfolio.xlsx", index=False)
