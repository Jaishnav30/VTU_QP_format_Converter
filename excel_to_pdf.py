from win32com import client 
    
def convert_to_pdf(ex_path):
    excel = client.Dispatch("Excel.Application") # This selects and runs the excel application
    excel.Visible = False                        # this prevents the excel to run in background

    sheets = excel.Workbooks.Open(ex_path) 
    ws = sheets.Worksheets[0]  # selecting the first worksheet
    
    index = ex_path.rfind('\\')             # .rfind method is used here to set the output file path same as input excel file's path
    pdf_path = ex_path[:index+1] + "op.pdf" # Please enter the name of the output PDF file or manually enter the path where you want to store the output PDF
    
    ws.PageSetup.PaperSize = 9  # 9 is for A4 Paper Size
    ws.PageSetup.Zoom = False  
    ws.PageSetup.FitToPagesWide = 1  
    ws.PageSetup.FitToPagesTall = False  

    ws.ExportAsFixedFormat(0, pdf_path)

    sheets.Close(SaveChanges=False)
    excel.Quit()

    print(f"PDF saved at: {pdf_path}")
