from docx import Document
from docx.shared import Inches
import openpyxl


def WordGen(ticker, hist, chart_path):
#Generating word report
    if ticker:
        doc = Document()
        
        doc.add_heading(f'Report for {ticker}', 0)

        doc.add_heading('Latest 10 Days - Open & Close prices', level=1)
        table = doc.add_table(rows=1, cols=3)
        table.style = 'Table Grid'
        
        hdr_cells = table.rows[0].cells
        
        hdr_cells[0].text = 'Date'
        hdr_cells[1].text = 'Open Price'
        hdr_cells[2].text = 'Close Price'

        for date, row in hist.iterrows():
            row_cells = table.add_row().cells
            row_cells[0].text = str(date.date())
            row_cells[1].text = f"{row['Open']:.2f}"
            row_cells[2].text = f"{row['Close']:.2f}"
        
        #adding data visualization to the report using docx Inches
        doc.add_heading('Price Chart (Last 30 days)', level=1)
        doc.add_picture(chart_path, width=Inches(6))
        print("Visualization added to the report")

        #adding news to the report
        # oc.add_heading('recent news headlines', level=1)
        # doc.add_paragraph(response_text)
        
        doc.save("report.docx")
        print("Report.docx generated")
        #TODO rewrite hist so its compatible with excel
        #export to excel
        hist.index = hist.index.tz_localize(None)
        hist[['Open', 'Close']].to_excel("reportExcel.xlsx")
        print("Report.xlsx generated")
    else: 
        print("Couldn't export docs due to missing ticker")