import yfinance as yf
import pandas as pd
import datetime as dt
import logging

logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)
ch = logging.StreamHandler()
ch.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
logger.addHandler(ch)

tickers = [""]    #Insert Tickers

logger.info("Starting earnings/calendar data download for %d tickers", len(tickers))

earnings_dict = {}

for ticker in tickers:
    
    logger.info("Fetching calendar for %s", ticker)
    
    info = yf.Ticker(ticker)
    
    data = {'Next Earnings': None, 'Eps Pred': None, 'Rev Pred': None}

    try:
    
        cal = info.calendar  
    
        ed = cal.get('Earnings Date')
    
        if isinstance(ed, (list, tuple)) and ed:
    
            data['Next Earnings'] = ed[0]
    
        elif isinstance(ed, dt.date):
    
            data['Next Earnings'] = ed

        data['Eps Pred'] = cal.get('Earnings Average')

        data['Rev Pred'] = cal.get('Revenue Average')

    except Exception as e:
    
        logger.error("Error parsing calendar for %s: %s", ticker, e)

    earnings_dict[ticker] = data

df = (
    pd.DataFrame.from_dict(earnings_dict, orient='index').sort_values(by='Next Earnings', na_position='last')
)

df.index.name = 'Ticker'

logger.info("Download complete; writing to Excel")

with pd.ExcelWriter('Earnings_Date.xlsx', engine='openpyxl') as writer:
    df.to_excel(writer)

logger.info("Data written to Earnings_Date.xlsx")
