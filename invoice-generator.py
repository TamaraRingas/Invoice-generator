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
        invoice_number = sheet.ceil(row = i, column = 1).value
        parent = sheet.ceil(row = i, column = 2).value
        lessons = sheet.ceil(row = i, column = 3).value
        price_per_lesson = sheet.ceil(row = i, column = 4).value
        total = sheet.ceil(row = i, column = 5).value
        invoice_date = sheet.ceil(row = i, column = 6).value

        doc = SimpleDocTemplate(str(parent + invoice_date) + ".pdf", pagesize = letter)

        data = [ ['Kate de Gruchy Invoice for Piano Lessons'],
                 ['Invoice number: ', invoice_number, 'on ', invoice_date],
                 ['Issued to: ', parent],
                 ['Lessons: ', lessons],
                 ['Price per lesson: ', price_per_lesson],
                 ['Total Due: ', total]      
        ]

        table = Table(data)
        elements = []
        elements.append(table)

        doc.build(elements)

