"""행사내역 엑셀 생성."""
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
import os


_THIN = Side(style='thin')
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)


def _cell(ws, row, col, value, bold=False, align='left', number_format=None):
    c = ws.cell(row=row, column=col, value=value)
    c.font = Font(name='맑은 고딕', bold=bold, size=10)
    c.alignment = Alignment(horizontal=align, vertical='center', wrap_text=True)
    c.border = _BORDER
    if number_format:
        c.number_format = number_format
    return c


def generate_exercise_excel(round_name: str, exercise_date: str,
                             applicants: list, output_path: str) -> str:
    """
    행사내역 엑셀 생성.
    applicants: list of dicts with name, exercise_price, quantity, broker, account_number
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '행사내역'

    # 제목
    ws.merge_cells('A1:G1')
    title_cell = ws['A1']
    title_cell.value = f'스톡옵션 행사내역 ({exercise_date})'
    title_cell.font = Font(name='맑은 고딕', bold=True, size=13)
    title_cell.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 28

    # 헤더
    headers = ['이름', '행사가액(원)', '수량(주)', '납입금액(원)', '증권사', '계좌번호', '비고']
    for ci, h in enumerate(headers, 1):
        _cell(ws, 2, ci, h, bold=True, align='center')
    ws.row_dimensions[2].height = 20

    # 데이터
    row = 3
    price_totals = {}  # {price: {'qty': 0, 'amount': 0}}

    for ap in applicants:
        price = ap.get('exercise_price') or 0
        qty   = ap.get('quantity') or 0
        amt   = price * qty

        _cell(ws, row, 1, ap.get('name', ''))
        _cell(ws, row, 2, price,  align='right', number_format='#,##0')
        _cell(ws, row, 3, qty,    align='right', number_format='#,##0')
        _cell(ws, row, 4, amt,    align='right', number_format='#,##0')
        _cell(ws, row, 5, ap.get('broker', '') or '')
        _cell(ws, row, 6, ap.get('account_number', '') or '')
        _cell(ws, row, 7, '')
        ws.row_dimensions[row].height = 18

        if price not in price_totals:
            price_totals[price] = {'qty': 0, 'amount': 0}
        price_totals[price]['qty']    += qty
        price_totals[price]['amount'] += amt

        row += 1

    # 가격별 소계
    row += 1
    ws.merge_cells(f'A{row}:G{row}')
    ws.cell(row=row, column=1).value = '[ 행사가액별 집계 ]'
    ws.cell(row=row, column=1).font = Font(name='맑은 고딕', bold=True, size=11)
    ws.row_dimensions[row].height = 22
    row += 1

    sub_headers = ['행사가액(원)', '수량(주)', '납입금액(원)']
    for ci, h in enumerate(sub_headers, 1):
        _cell(ws, row, ci, h, bold=True, align='center')
    row += 1

    total_qty = 0
    total_amt = 0
    for price in sorted(price_totals.keys()):
        info = price_totals[price]
        _cell(ws, row, 1, price,          align='right', number_format='#,##0')
        _cell(ws, row, 2, info['qty'],    align='right', number_format='#,##0')
        _cell(ws, row, 3, info['amount'], align='right', number_format='#,##0')
        total_qty += info['qty']
        total_amt += info['amount']
        row += 1

    # 합계
    _cell(ws, row, 1, '합계', bold=True, align='center')
    _cell(ws, row, 2, total_qty, bold=True, align='right', number_format='#,##0')
    _cell(ws, row, 3, total_amt, bold=True, align='right', number_format='#,##0')

    # 열 너비
    ws.column_dimensions['A'].width = 12
    ws.column_dimensions['B'].width = 14
    ws.column_dimensions['C'].width = 12
    ws.column_dimensions['D'].width = 16
    ws.column_dimensions['E'].width = 14
    ws.column_dimensions['F'].width = 22
    ws.column_dimensions['G'].width = 12

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    wb.save(output_path)
    return output_path
