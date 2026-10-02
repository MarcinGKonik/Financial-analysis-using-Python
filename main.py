import pandas as pd 
import os
import openai 
from langchain_community.tools.yahoo_finance_news import YahooFinanceNewsTool
from langgraph.prebuilt import create_react_agent

import YfinanceDownload
import WordGen
import ChartGen

ticker, hist = YfinanceDownload.YfinanceDownload()

chart_path = ChartGen.ChartGen(
    ticker,
    hist
)

WordGen.WordGen(
    ticker,
    hist,
    chart_path
)

#os.environ["OPENAI_API_KEY"] = "YOUR_API_KEY"





