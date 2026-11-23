"""弘法 Excel 导出统一格式；仅用于报表，不读写数据库。"""
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins

GREEN='176B39'
LIGHT='EDF6EF'
BORDER='DCE8DF'

def polish(ws, header_rows=1, freeze='A2', filter_rows=True, fixed_widths=None):
    ws.freeze_panes=freeze
    ws.sheet_view.showGridLines=False
    ws.sheet_properties.pageSetUpPr.fitToPage=True
    ws.page_setup.fitToWidth=1
    ws.sheet_properties.outlinePr.summaryBelow=True
    ws.print_options.horizontalCentered=True
    ws.page_margins=PageMargins(left=.25,right=.25,top=.45,bottom=.45,header=.2,footer=.2)
    ws.print_title_rows=f'1:{header_rows}'
    if filter_rows and ws.max_column and ws.max_row:
        ws.auto_filter.ref=f'A1:{get_column_letter(ws.max_column)}{max(ws.max_row,header_rows+1)}'
    edge=Side(style='hair',color=BORDER)
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment=Alignment(vertical='center',wrap_text=(cell.row<=header_rows))
            cell.border=Border(bottom=edge)
            if cell.row<=header_rows:
                cell.fill=PatternFill('solid',fgColor=GREEN)
                cell.font=Font(name='Calibri',size=11,bold=True,color='FFFFFF')
            else:
                cell.font=Font(name='Calibri',size=11)
                if cell.row%2==0:
                    cell.fill=PatternFill('solid',fgColor=LIGHT)
        ws.row_dimensions[row[0].row].height=30 if row[0].row<=header_rows else 22
    for col in ws.columns:
        letter=get_column_letter(col[0].column)
        max_len=max((len(str(c.value or '')) for c in col[:300]),default=10)
        ws.column_dimensions[letter].width=min(45,max(13,max_len+3))
    for key,width in (fixed_widths or {}).items():
        ws.column_dimensions[key].width=width
