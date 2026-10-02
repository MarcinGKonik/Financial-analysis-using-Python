import yfinance as yf
#yf.enable_debug_mode()

#Downloading stock data from yfinance
def YfinanceDownload():
   ticker = "AAPL" 
   print(f"Downloading historic data on {ticker}")
   stock = yf.Ticker(ticker)
   hist = stock.history(period="30d")
   print(hist[['Open', 'Close']])

   return ticker, hist
