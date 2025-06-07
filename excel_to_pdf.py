import os
import sys
import win32com.client

def convert_excel_to_pdf(excel_path, pdf_path=None):
    if not pdf_path:
        pdf_path = os.path.splitext(excel_path)[0] + '.pdf'
    
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False

        wb = excel.Workbooks.Open(os.path.abspath(excel_path))

        try:
            wb.ExportAsFixedFormat(0, os.path.abspath(pdf_path))
        except Exception:
            ws = wb.Worksheets[0]
            ws.ExportAsFixedFormat(0, os.path.abspath(pdf_path))

        wb.Close(False)
        excel.Quit()

        return pdf_path
    except Exception as e:
        print(f"PDF donusturme hatasi: {str(e)}")
        return None
