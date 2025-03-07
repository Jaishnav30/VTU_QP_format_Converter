from win32com import client 
    
def convert_to_pdf(ex_path):
    excel = client.Dispatch("Excel.Application") 
    excel.Visible = False 

    sheets = excel.Workbooks.Open(ex_path) 
    ws = sheets.Worksheets[0]  
    
    index = ex_path.rfind('\\')
    pdf_path = ex_path[:index+1] + "op.pdf"
    
    ws.PageSetup.PaperSize = 9  # A4 Paper Size
    ws.PageSetup.Zoom = False  
    ws.PageSetup.FitToPagesWide = 1  
    ws.PageSetup.FitToPagesTall = 1  

    ws.ExportAsFixedFormat(0, pdf_path)

    sheets.Close(SaveChanges=False)
    excel.Quit()

    print(f"PDF saved at: {pdf_path}")
