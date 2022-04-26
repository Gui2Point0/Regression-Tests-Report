from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl import load_workbook
from update_spreadsheet import filename

#Make titles font bold
def MakeTextBold(cell,ws):
    ws[cell].font = Font(bold=True)

#Align titles to the center
def CenterText(cell, ws):
    ws[cell].alignment = Alignment(horizontal='center', vertical='center')

#Pass and Fail background
def CellBackgroung(ws): 
    green = '#90EE90' 
    red = '#FF6961'

    for rows in ws.iter_rows(min_row=1, max_row=58, min_col=1, max_col=3):
        for cell in rows:
                if cell.row == 'Pass':
                    cell.fill = PatternFill(start_color=green, end_color=green, fill_type='solid')
                elif cell.row == 'Fail':
                    cell.fill = PatternFill(start_color=red, end_color=red, fill_type='solid')

#Styling columns info title
def ColumnsTitle(ws):
    CenterText('A1', ws)
    MakeTextBold('A1', ws)
    
    CenterText('B1', ws)
    MakeTextBold('B1', ws)
    
    CenterText('C1', ws)
    MakeTextBold('C1', ws)

##########################################################################################################################

#Main
def SheetStyles():
    #Loading existing Excel File into work_book Object
    wb = load_workbook(filename + '.xlsx') 
    ws = wb.active
    
    #Call methods
    ColumnsTitle(ws)
    CellBackgroung(ws)

    #Save file
    wb.save(filename + '.xlsx')