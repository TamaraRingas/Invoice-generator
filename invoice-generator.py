import openpyxl
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, inch
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle

pdfmetrics.registerFont(TTFont('Arial', 'Arial.ttf'))

#import excel file
wb = openpyxl.load_workbook('Invoices.xlsx')
sheet = wb.get_sheet_by_name('invoices')

page_height = 3050
page_width = 2156
margin = 50

def create_invoice():
    for i in range(2,7):
        invoice_number = sheet.ceil(row = i, column = 1)
