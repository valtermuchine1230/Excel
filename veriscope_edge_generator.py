"""
VERISCOPE EDGE - 3-Tier Generator (Basic / Pro / Elite)
Python 3.10+ / openpyxl only.
Fixes the two known bugs:
  1) Risk Manager: Broker Name kept fully separate from numeric calc cells (no #VALUE!).
  2) Drawdown Tracker / Equity Curve: previously referenced Tbl_Journal[Underwater] and
     Tbl_Journal[Drawdown %], columns that DO NOT EXIST in the Journal table -> #REF!.
     Fixed by building local helper columns (INDEX-based, not structured refs to
     nonexistent columns) inside each sheet that needs them.
Output: ./output/Veriscope_Edge_Basic.xlsx / _Pro.xlsx / _Elite.xlsx
"""
import os
from datetime import datetime
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.chart import LineChart, BarChart, Reference
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
from openpyxl.workbook.defined_name import DefinedName

from styles import COLORS, StylePresets, apply_style

PASSWORD = 'veriscope2026'
N_JOURNAL_ROWS = 500          # Journal rows 5..504
JOURNAL_FIRST_ROW = 5
JOURNAL_LAST_ROW = JOURNAL_FIRST_ROW + N_JOURNAL_ROWS - 1  # 504

# ---------------------------------------------------------------------------
# HELPERS
# ---------------------------------------------------------------------------

def set_sheet_appearance(ws):
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = 'landscape'
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True


def add_title_band(ws, title, subtitle='', row=1, colspan=10):
    rng = f"A{row}:{get_column_letter(colspan)}{row}"
    ws.merge_cells(rng)
    cell = ws[f"A{row}"]
    cell.value = title
    ws.row_dimensions[row].height = 30
    apply_style(cell, StylePresets.title_band())
    if subtitle:
        rng2 = f"A{row+1}:{get_column_letter(colspan)}{row+1}"
        ws.merge_cells(rng2)
        c2 = ws[f"A{row+1}"]
        c2.value = subtitle
        c2.font = Font(size=9, italic=True, color=COLORS['text_dark'])
        c2.alignment = Alignment(horizontal='center')


def named_range(wb, name, sheet, cell):
    wb.defined_names[name] = DefinedName(name, attr_text=f"'{sheet}'!${cell.replace('$','')}"
                                          .replace('$', '$', 0))
    # simpler explicit form:
    ref = cell
    col = ''.join(ch for ch in ref if ch.isalpha())
    row = ''.join(ch for ch in ref if ch.isdigit())
    wb.defined_names[name] = DefinedName(name, attr_text=f"'{sheet}'!${col}${row}")


def protect(ws):
    ws.protection.sheet = True
    ws.protection.password = PASSWORD
    ws.protection.enable()


def home_link(ws, target='03_Dashboard'):
    # Column AB (28) is never covered by any title-band merge in this workbook
    # (max colspan used is 26 -> column Z), so it's always a safe, unmerged cell.
    ws['AB1'] = f'=HYPERLINK("#{target}!A1","Home")'
    ws['AB1'].font = Font(size=9, color=COLORS['accent_gold'], underline='single')


# ---------------------------------------------------------------------------
# SHEET BUILDERS
# ---------------------------------------------------------------------------

def build_lists_sheet(wb):
    ws = wb.create_sheet('99_Lists')
    ws.sheet_state = 'hidden'
    set_sheet_appearance(ws)
    data = {
        'SETUPS': ['Breakout', 'Pullback', 'Support/Resistance', 'Fibonacci', 'MA Cross', 'Pin Bar', 'Divergence', 'Custom'],
        'SYMBOLS': ['EURUSD', 'GBPUSD', 'USDJPY', 'GOLD', 'BTCUSD', 'AAPL', 'MSFT', 'Other'],
        'EMOTIONS': ['Calm', 'Focused', 'Anxious', 'FOMO', 'Tilt', 'Confident', 'Bored'],
        'RULES_BROKEN': ['Entry Criteria', 'Position Sizing', 'Stop Placement', 'Early Exit', 'Overtrading', 'Revenge Trade', 'None'],
        'ASSET_CLASSES': ['Forex', 'Stocks', 'Futures', 'Crypto', 'Options'],
    }
    col = 1
    for name, items in data.items():
        ws.cell(row=1, column=col, value=name).font = Font(bold=True, color=COLORS['accent_gold'])
        for i, item in enumerate(items, start=2):
            ws.cell(row=i, column=col, value=item)
        col += 1
    return data


def build_settings_sheet(wb):
    ws = wb.create_sheet('02_Settings')
    set_sheet_appearance(ws)
    add_title_band(ws, 'Settings & Configuration', 'Your control panel — feeds every other sheet automatically', 1, 6)
    rows = [
        ('Starting Capital ($)', 10000, 'C5'),
        ('Base Currency', 'USD', 'C6'),
        ('Currency Symbol', '$', 'C7'),
        ('Default Risk % per Trade', 1, 'C8'),
        ('Daily Loss Limit ($)', 500, 'C9'),
        ('Trader Name', '', 'C10'),
        ('Broker(s)', '', 'C11'),
        ('Timezone', 'UTC', 'C12'),
    ]
    r = 4
    for label, default, ref in rows:
        ws.cell(row=r, column=1, value=label).font = Font(bold=True, size=10)
        c = ws[ref]
        c.value = default
        apply_style(c, StylePresets.input_cell())
        r += 1
    named_range(wb, 'AccountBalance', '02_Settings', 'C5')
    named_range(wb, 'CurrencySymbol', '02_Settings', 'C7')
    named_range(wb, 'RiskPercent', '02_Settings', 'C8')
    named_range(wb, 'DailyLossLimit', '02_Settings', 'C9')
    ws.column_dimensions['A'].width = 26
    ws.column_dimensions['C'].width = 22
    protect(ws)


def build_cover_sheet(wb, tier, sheet_links):
    ws = wb.create_sheet('00_Cover')
    set_sheet_appearance(ws)
    for row in range(1, 40):
        for col in range(1, 12):
            ws.cell(row=row, column=col).fill = PatternFill('solid', fgColor=COLORS['bg_dark'])
    ws.merge_cells('A4:K5')
    t = ws['A4']
    t.value = 'VERISCOPE EDGE'
    t.font = Font(size=30, bold=True, color=COLORS['accent_gold'])
    t.alignment = Alignment(horizontal='center', vertical='center')
    subtitle_map = {
        'basic': 'Basic Edition — Journal, Risk & Drawdown Essentials',
        'pro': 'Pro Edition — Full Trading Analytics Suite',
        'elite': 'Elite Edition — The Complete All-in-One System',
    }
    ws.merge_cells('A7:K7')
    s = ws['A7']
    s.value = subtitle_map[tier]
    s.font = Font(size=13, bold=True, color=COLORS['text_light'])
    s.alignment = Alignment(horizontal='center')
    ws.merge_cells('A9:K11')
    d = ws['A9']
    d.value = 'Know Your Edge. Trade With Proof.\nVersion 1.0'
    d.font = Font(size=10, color=COLORS['text_light'])
    d.alignment = Alignment(horizontal='center', wrap_text=True)

    row = 14
    ws.cell(row=row, column=1, value='NAVIGATION').font = Font(size=11, bold=True, color=COLORS['accent_gold'])
    row += 1
    for sheet_name, label in sheet_links:
        cell = ws.cell(row=row, column=1)
        cell.value = f'=HYPERLINK("#{sheet_name}!A1","{label}")'
        cell.font = Font(size=10, color=COLORS['profit_green'], underline='single')
        row += 1
    ws.column_dimensions['A'].width = 34


def build_how_to_use_sheet(wb):
    ws = wb.create_sheet('01_How_To_Use')
    set_sheet_appearance(ws)
    add_title_band(ws, 'How to Use Veriscope Edge', '', 1, 8)
    lines = [
        'QUICK START',
        '1. Open SETTINGS and set your starting capital and risk % per trade.',
        '2. Log every trade in TRADING JOURNAL as you take it.',
        '3. Watch DASHBOARD update live — no manual work needed.',
        '4. Use RISK MANAGER before entering any trade to size your position.',
        '5. Review DRAWDOWN TRACKER weekly to know your worst-case exposure.',
        '',
        'GLOSSARY',
        'R-Multiple: your profit/loss expressed as a multiple of what you risked.',
        'Drawdown: decline from peak balance to lowest point since.',
        'Win Rate: % of trades that were profitable.',
        'Expectancy: average R gained per trade — the heart of your edge.',
    ]
    for i, line in enumerate(lines, start=4):
        c = ws.cell(row=i, column=1, value=line)
        c.font = Font(bold=line.isupper(), size=10)
    ws.column_dimensions['A'].width = 90
    protect(ws)


def build_dashboard_sheet(wb, tier):
    ws = wb.create_sheet('03_Dashboard')
    set_sheet_appearance(ws)
    add_title_band(ws, 'Trading Dashboard', 'Live overview — updates automatically from the Journal', 1, 10)
    home_link(ws)

    kpis = [
        ('Current Balance', '=IFERROR(AccountBalance+SUM(Tbl_Journal[Net P&L]),AccountBalance)', COLORS['accent_gold']),
        ('Net P&L', '=IFERROR(SUM(Tbl_Journal[Net P&L]),0)', COLORS['profit_green']),
        ('Win Rate %', '=IFERROR(COUNTIF(Tbl_Journal[Net P&L],">0")/COUNTA(Tbl_Journal[Trade #])*100,0)', COLORS['profit_green']),
        ('Profit Factor', '=IFERROR(SUMIF(Tbl_Journal[Net P&L],">0")/ABS(SUMIF(Tbl_Journal[Net P&L],"<0")),0)', COLORS['accent_gold']),
        ('Expectancy (R)', '=IFERROR(AVERAGE(Tbl_Journal[R-Multiple]),0)', COLORS['accent_gold']),
        ('Total Trades', '=COUNTA(Tbl_Journal[Trade #])', COLORS['accent_gold']),
        ('Best Trade', '=IFERROR(MAX(Tbl_Journal[Net P&L]),0)', COLORS['profit_green']),
        ('Worst Trade', '=IFERROR(MIN(Tbl_Journal[Net P&L]),0)', COLORS['loss_red']),
    ]
    row, col = 4, 1
    for label, formula, color in kpis:
        ws.cell(row=row, column=col, value=label)
        apply_style(ws.cell(row=row, column=col), StylePresets.kpi_label())
        v = ws.cell(row=row + 1, column=col, value=formula)
        apply_style(v, StylePresets.kpi_value(color))
        ws.row_dimensions[row + 1].height = 30
        col += 1
        if col > 4:
            col = 1
            row += 3

    if tier != 'basic':
        chart_row = row + 1
        ws.cell(row=chart_row, column=1, value='Equity Curve').font = Font(bold=True, color=COLORS['accent_gold'])
        chart = LineChart()
        chart.title = 'Equity Curve'
        chart.style = 2
        data = Reference(wb['04_Trading_Journal'], min_col=20, min_row=4, max_row=JOURNAL_LAST_ROW)
        chart.add_data(data, titles_from_data=True)
        chart.height, chart.width = 8, 18
        ws.add_chart(chart, f'A{chart_row + 1}')

    for c in 'ABCD':
        ws.column_dimensions[c].width = 18
    protect(ws)


def build_trading_journal_sheet(wb, lists):
    ws = wb.create_sheet('04_Trading_Journal')
    set_sheet_appearance(ws)
    add_title_band(ws, 'Trading Journal', 'The engine — log every trade here', 1, 26)
    home_link(ws)

    groups = [('A3:H3', 'PRE-TRADE SETUP', '1E3A5F'), ('I3:Q3', 'POST-TRADE RESULTS', '1F5E4F'), ('R3:Z3', 'ANALYSIS & NOTES', '4A3728')]
    for rng, label, color in groups:
        ws.merge_cells(rng)
        c = ws[rng.split(':')[0]]
        c.value = label
        c.fill = PatternFill('solid', fgColor=color)
        c.font = Font(bold=True, color='FFFFFF')
        c.alignment = Alignment(horizontal='center')

    columns = ['Trade #', 'Date', 'Time', 'Symbol', 'Direction', 'Setup', 'Market Condition', 'Confidence',
               'Entry Price', 'Stop Loss', 'Position Size', 'Risk $ (1R)', 'Risk %',
               'Exit Price', 'Exit Date', 'Gross P&L', 'Commission', 'Net P&L', 'R-Multiple', 'Running Balance',
               'Rule Compliant?', 'Rule Broken', 'Emotion', 'Tags', 'Screenshot Link', 'Notes']
    for i, name in enumerate(columns, start=1):
        c = ws.cell(row=4, column=i, value=name)
        apply_style(c, StylePresets.table_header())
    ws.row_dimensions[4].height = 26

    input_cols = ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'O', 'Q', 'U', 'V', 'W', 'X', 'Y', 'Z']
    for r in range(JOURNAL_FIRST_ROW, JOURNAL_LAST_ROW + 1):
        ws[f'A{r}'] = f'=IF(D{r}="","",COUNTA($D${JOURNAL_FIRST_ROW}:D{r}))'
        for col in input_cols:
            apply_style(ws[f'{col}{r}'], StylePresets.input_cell())
        ws[f'L{r}'] = f'=IF(OR(I{r}="",J{r}="",K{r}=""),"",ABS(I{r}-J{r})*K{r})'
        ws[f'M{r}'] = f'=IF(L{r}="","",IFERROR(L{r}/AccountBalance,""))'
        ws[f'P{r}'] = f'=IF(OR(E{r}="",N{r}=""),"",IF(E{r}="Long",(N{r}-I{r})*K{r},(I{r}-N{r})*K{r}))'
        ws[f'R{r}'] = f'=IF(P{r}="","",P{r}-IF(Q{r}="",0,Q{r}))'
        ws[f'S{r}'] = f'=IFERROR(IF(OR(R{r}="",L{r}=0,L{r}=""),"",R{r}/L{r}),"")'
        ws[f'T{r}'] = f'=IF(R{r}="","",AccountBalance+SUM($R${JOURNAL_FIRST_ROW}:R{r}))'
        for col in ['A', 'L', 'M', 'P', 'R', 'S', 'T']:
            apply_style(ws[f'{col}{r}'], StylePresets.locked_cell())

    widths = [8, 11, 9, 10, 9, 14, 15, 10, 11, 11, 11, 11, 9, 11, 11, 11, 11, 11, 11, 13, 12, 14, 11, 14, 14, 20]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    tab = Table(displayName='Tbl_Journal', ref=f'A4:Z{JOURNAL_LAST_ROW}')
    tab.tableStyleInfo = TableStyleInfo(name='TableStyleMedium2', showRowStripes=True)
    ws.add_table(tab)

    dv_defs = [
        (f'D{JOURNAL_FIRST_ROW}:D{JOURNAL_LAST_ROW}', f"='99_Lists'!$B$2:$B$9"),
        (f'E{JOURNAL_FIRST_ROW}:E{JOURNAL_LAST_ROW}', '"Long,Short"'),
        (f'F{JOURNAL_FIRST_ROW}:F{JOURNAL_LAST_ROW}', f"='99_Lists'!$A$2:$A$9"),
        (f'G{JOURNAL_FIRST_ROW}:G{JOURNAL_LAST_ROW}', '"Trending,Ranging,Choppy,News-Driven"'),
        (f'H{JOURNAL_FIRST_ROW}:H{JOURNAL_LAST_ROW}', '"1,2,3,4,5"'),
        (f'U{JOURNAL_FIRST_ROW}:U{JOURNAL_LAST_ROW}', '"Yes,No"'),
        (f'V{JOURNAL_FIRST_ROW}:V{JOURNAL_LAST_ROW}', f"='99_Lists'!$D$2:$D$8"),
        (f'W{JOURNAL_FIRST_ROW}:W{JOURNAL_LAST_ROW}', f"='99_Lists'!$C$2:$C$8"),
    ]
    for rng, formula in dv_defs:
        dv = DataValidation(type='list', formula1=formula, allow_blank=True)
        ws.add_data_validation(dv)
        dv.add(rng)

    ws.freeze_panes = f'A{JOURNAL_FIRST_ROW}'
    protect(ws)


def build_risk_manager_sheet(wb):
    """FIX: Broker Name lives in its own disconnected info box — never touched by a formula."""
    ws = wb.create_sheet('05_Risk_Manager')
    set_sheet_appearance(ws)
    add_title_band(ws, 'Risk Manager', 'Position sizing calculator', 1, 8)
    home_link(ws)

    ws.merge_cells('A4:B4')
    ws['A4'] = 'POSITION SIZING CALCULATOR'
    apply_style(ws['A4'], StylePresets.section_header())

    ws['A6'] = 'Account Size'
    ws['B6'] = '=AccountBalance'
    apply_style(ws['B6'], StylePresets.locked_cell())
    ws['A7'] = 'Risk % per Trade'
    ws['B7'] = '=RiskPercent'
    apply_style(ws['B7'], StylePresets.input_cell())
    ws['A8'] = 'Entry Price'
    apply_style(ws['B8'], StylePresets.input_cell())
    ws['A9'] = 'Stop Loss Price'
    apply_style(ws['B9'], StylePresets.input_cell())

    ws['A11'] = 'Risk in $'
    ws['B11'] = '=IFERROR(IF(OR(B6="",B7=""),"",B6*B7/100),"")'
    apply_style(ws['B11'], StylePresets.locked_cell())
    ws['A12'] = 'Stop Distance (Pips/Units)'
    ws['B12'] = '=IFERROR(IF(OR(B8="",B9=""),"",ABS(B8-B9)),"")'
    apply_style(ws['B12'], StylePresets.locked_cell())
    ws['A13'] = 'Position Size (Units)'
    ws['B13'] = '=IFERROR(IF(OR(B11="",B12="",B12=0),"",B11/B12),"")'
    apply_style(ws['B13'], StylePresets.kpi_value(COLORS['accent_gold']))
    ws.row_dimensions[13].height = 30

    # Broker info — completely separate, informational only, no formula ever reads this.
    ws.merge_cells('D4:E4')
    ws['D4'] = 'BROKER INFO (reference only)'
    apply_style(ws['D4'], StylePresets.section_header())
    ws['D6'] = 'Broker Name'
    apply_style(ws['E6'], StylePresets.input_cell())
    ws['D7'] = 'Account Currency'
    apply_style(ws['E7'], StylePresets.input_cell())
    ws['D8'] = 'Leverage'
    apply_style(ws['E8'], StylePresets.input_cell())

    ws.column_dimensions['A'].width = 22
    ws.column_dimensions['B'].width = 16
    ws.column_dimensions['D'].width = 20
    ws.column_dimensions['E'].width = 16
    protect(ws)


def build_drawdown_tracker_sheet(wb):
    """FIX: build local helper columns (Balance/RunningMax/Drawdown%/Underwater) via INDEX
    against Tbl_Journal[Running Balance] — never reference a nonexistent table column."""
    ws = wb.create_sheet('06_Drawdown_Tracker')
    set_sheet_appearance(ws)
    add_title_band(ws, 'Drawdown Tracker', 'Peak-to-trough analysis', 1, 8)
    home_link(ws)

    hdr_row = 9
    headers = ['#', 'Balance', 'Running Max', 'Drawdown %', 'Underwater']
    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=hdr_row, column=i, value=h)
        apply_style(c, StylePresets.table_header())

    first, last = hdr_row + 1, hdr_row + N_JOURNAL_ROWS
    for i, r in enumerate(range(first, last + 1), start=1):
        ws[f'A{r}'] = i
        ws[f'B{r}'] = f'=IFERROR(INDEX(Tbl_Journal[Running Balance],{i}),"")'
        ws[f'C{r}'] = f'=IF(B{r}="","",MAX($B${first}:B{r}))'
        ws[f'D{r}'] = f'=IF(OR(B{r}="",C{r}=0),"",(B{r}-C{r})/C{r})'
        ws[f'E{r}'] = f'=IF(D{r}="","",IF(D{r}<0,D{r},0))'
        for col in 'ABCDE':
            apply_style(ws[f'{col}{r}'], StylePresets.locked_cell())
        ws[f'D{r}'].number_format = '0.00%'
        ws[f'E{r}'].number_format = '0.00%'

    ws['A4'] = 'Current Drawdown %'
    apply_style(ws['A4'], StylePresets.kpi_label())
    ws['A5'] = f'=IFERROR(INDEX($D${first}:D{last},COUNT($B${first}:B{last})),0)'
    apply_style(ws['A5'], StylePresets.kpi_value(COLORS['loss_red']))
    ws['A5'].number_format = '0.00%'

    ws['B4'] = 'Max Drawdown %'
    apply_style(ws['B4'], StylePresets.kpi_label())
    ws['B5'] = f'=IFERROR(MIN($D${first}:D{last}),0)'
    apply_style(ws['B5'], StylePresets.kpi_value(COLORS['loss_red']))
    ws['B5'].number_format = '0.00%'

    ws['C4'] = 'Days in Drawdown'
    apply_style(ws['C4'], StylePresets.kpi_label())
    ws['C5'] = f'=COUNTIF($E${first}:E{last},"<0")'
    apply_style(ws['C5'], StylePresets.kpi_value(COLORS['accent_gold']))

    chart = LineChart()
    chart.title = 'Equity vs Underwater'
    data = Reference(ws, min_col=2, min_row=hdr_row, max_row=last)
    chart.add_data(data, titles_from_data=True)
    chart.height, chart.width = 8, 18
    ws.add_chart(chart, 'G4')

    for c in 'ABCDE':
        ws.column_dimensions[c].width = 13
    protect(ws)


def build_profit_dashboard_sheet(wb):
    ws = wb.create_sheet('07_Profit_Dashboard')
    set_sheet_appearance(ws)
    add_title_band(ws, 'Profit Dashboard', 'Full KPI suite', 1, 10)
    home_link(ws)
    kpis = [
        ('Win Rate %', '=IFERROR(COUNTIF(Tbl_Journal[Net P&L],">0")/COUNTA(Tbl_Journal[Trade #])*100,0)', COLORS['profit_green']),
        ('Profit Factor', '=IFERROR(SUMIF(Tbl_Journal[Net P&L],">0")/ABS(SUMIF(Tbl_Journal[Net P&L],"<0")),0)', COLORS['accent_gold']),
        ('Avg Win $', '=IFERROR(AVERAGEIF(Tbl_Journal[Net P&L],">0"),0)', COLORS['profit_green']),
        ('Avg Loss $', '=IFERROR(AVERAGEIF(Tbl_Journal[Net P&L],"<0"),0)', COLORS['loss_red']),
        ('Payoff Ratio', '=IFERROR(ABS(AVERAGEIF(Tbl_Journal[Net P&L],">0")/AVERAGEIF(Tbl_Journal[Net P&L],"<0")),0)', COLORS['accent_gold']),
        ('Expectancy (R)', '=IFERROR(AVERAGE(Tbl_Journal[R-Multiple]),0)', COLORS['accent_gold']),
        ('Net P&L Total', '=IFERROR(SUM(Tbl_Journal[Net P&L]),0)', COLORS['profit_green']),
        ('Largest Win', '=IFERROR(MAX(Tbl_Journal[Net P&L]),0)', COLORS['profit_green']),
        ('Largest Loss', '=IFERROR(MIN(Tbl_Journal[Net P&L]),0)', COLORS['loss_red']),
        ('Total Trades', '=COUNTA(Tbl_Journal[Trade #])', COLORS['accent_gold']),
    ]
    row, col = 4, 1
    for label, formula, color in kpis:
        ws.cell(row=row, column=col, value=label)
        apply_style(ws.cell(row=row, column=col), StylePresets.kpi_label())
        v = ws.cell(row=row + 1, column=col, value=formula)
        apply_style(v, StylePresets.kpi_value(color))
        col += 1
        if col > 5:
            col = 1
            row += 3
    for c in 'ABCDE':
        ws.column_dimensions[c].width = 16
    protect(ws)


def build_r_multiple_tracker_sheet(wb):
    ws = wb.create_sheet('08_R_Multiple_Tracker')
    set_sheet_appearance(ws)
    add_title_band(ws, 'R-Multiple Tracker', 'Distribution and edge-by-setup', 1, 8)
    home_link(ws)
    stats = [
        ('Mean R', '=IFERROR(AVERAGE(Tbl_Journal[R-Multiple]),0)'),
        ('% Positive R', '=IFERROR(COUNTIF(Tbl_Journal[R-Multiple],">0")/COUNTA(Tbl_Journal[Trade #])*100,0)'),
        ('Trades >= 2R', '=COUNTIF(Tbl_Journal[R-Multiple],">=2")'),
        ('Trades <= -1R', '=COUNTIF(Tbl_Journal[R-Multiple],"<=-1")'),
    ]
    for i, (label, formula) in enumerate(stats, start=4):
        ws[f'A{i}'] = label
        ws[f'B{i}'] = formula
        apply_style(ws[f'B{i}'], StylePresets.locked_cell())
    ws.column_dimensions['A'].width = 20
    ws.column_dimensions['B'].width = 15
    protect(ws)


def build_trade_statistics_sheet(wb):
    ws = wb.create_sheet('09_Trade_Statistics')
    set_sheet_appearance(ws)
    add_title_band(ws, 'Trade Statistics', 'Advanced risk-adjusted metrics', 1, 10)
    home_link(ws)
    stats = [
        ('Sharpe Ratio', '=IFERROR(AVERAGE(Tbl_Journal[Net P&L])/STDEV(Tbl_Journal[Net P&L]),0)', 'Above 1.0 = good, above 2.0 = excellent'),
        ('SQN', '=IFERROR(SQRT(COUNTA(Tbl_Journal[Trade #]))*AVERAGE(Tbl_Journal[R-Multiple])/STDEV(Tbl_Journal[R-Multiple]),0)', 'Above 2.0 = solid edge'),
    ]
    r = 4
    for label, formula, note in stats:
        ws.cell(row=r, column=1, value=label).font = Font(bold=True)
        c = ws.cell(row=r, column=2, value=formula)
        apply_style(c, StylePresets.locked_cell())
        ws.cell(row=r, column=3, value=note).font = Font(italic=True, size=9)
        r += 2
    ws.column_dimensions['A'].width = 18
    ws.column_dimensions['B'].width = 14
    ws.column_dimensions['C'].width = 34
    protect(ws)


def build_compounding_calculator_sheet(wb):
    ws = wb.create_sheet('14_Compounding_Calculator')
    set_sheet_appearance(ws)
    add_title_band(ws, 'Compounding Calculator', 'Growth projections', 1, 8)
    home_link(ws)
    ws['A4'] = 'Starting Capital'
    ws['B4'] = '=AccountBalance'
    apply_style(ws['B4'], StylePresets.locked_cell())
    ws['A5'] = 'Expected Monthly Return %'
    apply_style(ws['B5'], StylePresets.input_cell())
    ws['A6'] = 'Monthly Contribution'
    apply_style(ws['B6'], StylePresets.input_cell())
    for i, h in enumerate(['Month', 'Base Case'], start=1):
        c = ws.cell(row=8, column=i, value=h)
        apply_style(c, StylePresets.table_header())
    for r in range(9, 33):
        m = r - 8
        ws[f'A{r}'] = m
        prev = 'B4' if m == 1 else f'B{r-1}'
        ws[f'B{r}'] = f'=IFERROR({prev}*(1+$B$5/100)+$B$6,"")'
        apply_style(ws[f'A{r}'], StylePresets.locked_cell())
        apply_style(ws[f'B{r}'], StylePresets.locked_cell())
    ws.column_dimensions['A'].width = 16
    ws.column_dimensions['B'].width = 16
    protect(ws)


def build_expectancy_calculator_sheet(wb):
    ws = wb.create_sheet('15_Expectancy_Calculator')
    set_sheet_appearance(ws)
    add_title_band(ws, 'Expectancy Calculator', 'Standalone edge validator', 1, 8)
    home_link(ws)
    ws['A4'] = 'Win Rate %'
    apply_style(ws['B4'], StylePresets.input_cell())
    ws['A5'] = 'Average Win (R)'
    apply_style(ws['B5'], StylePresets.input_cell())
    ws['A6'] = 'Average Loss (R)'
    apply_style(ws['B6'], StylePresets.input_cell())
    ws['A7'] = 'Sample Size'
    apply_style(ws['B7'], StylePresets.input_cell())
    ws['A9'] = 'Expectancy (R)'
    ws['B9'] = '=IFERROR(IF(OR(B4="",B5="",B6=""),"",B4/100*B5-(1-B4/100)*ABS(B6)),"")'
    apply_style(ws['B9'], StylePresets.kpi_value(COLORS['accent_gold']))
    ws['A10'] = 'Sample Size Check'
    ws['B10'] = '=IF(B7="","",IF(B7>=30,"OK (30+ trades)","Insufficient (need 30+)"))'
    apply_style(ws['B10'], StylePresets.locked_cell())
    ws.column_dimensions['A'].width = 22
    ws.column_dimensions['B'].width = 22
    protect(ws)


def build_equity_curve_sheet(wb):
    """FIX: Max Drawdown here is now computed from a local helper column, not a
    nonexistent Tbl_Journal[Drawdown %]."""
    ws = wb.create_sheet('17_Equity_Curve')
    set_sheet_appearance(ws)
    add_title_band(ws, 'Equity Curve', 'Full showcase view', 1, 10)
    home_link(ws)

    hdr_row = 9
    ws.cell(row=hdr_row, column=1, value='#')
    ws.cell(row=hdr_row, column=2, value='Balance')
    ws.cell(row=hdr_row, column=3, value='Running Max')
    ws.cell(row=hdr_row, column=4, value='Drawdown %')
    for col in range(1, 5):
        apply_style(ws.cell(row=hdr_row, column=col), StylePresets.table_header())

    first, last = hdr_row + 1, hdr_row + N_JOURNAL_ROWS
    for i, r in enumerate(range(first, last + 1), start=1):
        ws[f'A{r}'] = i
        ws[f'B{r}'] = f'=IFERROR(INDEX(Tbl_Journal[Running Balance],{i}),"")'
        ws[f'C{r}'] = f'=IF(B{r}="","",MAX($B${first}:B{r}))'
        ws[f'D{r}'] = f'=IF(OR(B{r}="",C{r}=0),"",(B{r}-C{r})/C{r})'
        for col in 'ABCD':
            apply_style(ws[f'{col}{r}'], StylePresets.locked_cell())
        ws[f'D{r}'].number_format = '0.00%'

    ws['A4'] = 'Total Return %'
    apply_style(ws['A4'], StylePresets.kpi_label())
    ws['A5'] = '=IFERROR((MAX(Tbl_Journal[Running Balance])-AccountBalance)/AccountBalance*100,0)'
    apply_style(ws['A5'], StylePresets.kpi_value(COLORS['profit_green']))

    ws['B4'] = 'Max Drawdown %'
    apply_style(ws['B4'], StylePresets.kpi_label())
    ws['B5'] = f'=IFERROR(MIN($D${first}:D{last}),0)'
    apply_style(ws['B5'], StylePresets.kpi_value(COLORS['loss_red']))
    ws['B5'].number_format = '0.00%'

    ws['C4'] = 'Recovery Factor'
    apply_style(ws['C4'], StylePresets.kpi_label())
    ws['C5'] = '=IFERROR(A5/ABS(B5*100),0)'
    apply_style(ws['C5'], StylePresets.kpi_value(COLORS['accent_gold']))

    chart = LineChart()
    chart.title = 'Equity Curve'
    data = Reference(ws, min_col=2, min_row=hdr_row, max_row=last)
    chart.add_data(data, titles_from_data=True)
    chart.height, chart.width = 9, 20
    ws.add_chart(chart, 'F4')

    for c in 'ABCD':
        ws.column_dimensions[c].width = 14
    protect(ws)


def build_economic_calendar_sheet(wb):
    ws = wb.create_sheet('10_Economic_Calendar')
    set_sheet_appearance(ws)
    add_title_band(ws, 'Economic Calendar', 'Manual weekly entry — copy from a free source', 1, 10)
    home_link(ws)
    headers = ['Date', 'Time (UTC)', 'Currency', 'Event', 'Impact', 'Forecast', 'Previous', 'Actual', 'Avoid Trading?', 'Notes']
    for i, h in enumerate(headers, start=1):
        apply_style(ws.cell(row=5, column=i, value=h), StylePresets.table_header())
    dv = DataValidation(type='list', formula1='"High,Medium,Low"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add('E6:E55')
    for r in range(6, 56):
        for col in range(1, 9):
            apply_style(ws.cell(row=r, column=col), StylePresets.input_cell())
        ws[f'I{r}'] = f'=IF(E{r}="High","AVOID",IF(E{r}="","","OK"))'
        apply_style(ws[f'I{r}'], StylePresets.locked_cell())
        apply_style(ws[f'J{r}'], StylePresets.input_cell())
    for i in range(1, 11):
        ws.column_dimensions[get_column_letter(i)].width = 14
    protect(ws)


def build_financial_control_sheet(wb):
    ws = wb.create_sheet('11_Financial_Control')
    set_sheet_appearance(ws)
    add_title_band(ws, 'Financial Control', 'Deposits & withdrawals log', 1, 8)
    home_link(ws)
    headers = ['Date', 'Type', 'Amount', 'Note', 'Running Net Deposits']
    for i, h in enumerate(headers, start=1):
        apply_style(ws.cell(row=5, column=i, value=h), StylePresets.table_header())
    dv = DataValidation(type='list', formula1='"Deposit,Withdrawal"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add('B6:B55')
    for r in range(6, 56):
        for col in range(1, 4):
            apply_style(ws.cell(row=r, column=col), StylePresets.input_cell())
        ws[f'E{r}'] = f'=IF(C{r}="","",SUM($C$6:C{r}))'
        apply_style(ws[f'E{r}'], StylePresets.locked_cell())
    for c, w in zip('ABCDE', [12, 12, 12, 22, 18]):
        ws.column_dimensions[c].width = w
    protect(ws)


def build_lot_size_calculator_sheet(wb):
    ws = wb.create_sheet('12_Lot_Size_Calculator')
    set_sheet_appearance(ws)
    add_title_band(ws, 'Lot Size Calculator', 'Position sizing by asset class', 1, 8)
    home_link(ws)
    ws['A4'] = 'Account Size'
    ws['B4'] = '=AccountBalance'
    apply_style(ws['B4'], StylePresets.locked_cell())
    ws['A5'] = 'Risk %'
    apply_style(ws['B5'], StylePresets.input_cell())
    ws['A6'] = 'Entry Price'
    apply_style(ws['B6'], StylePresets.input_cell())
    ws['A7'] = 'Stop Price'
    apply_style(ws['B7'], StylePresets.input_cell())
    ws['A9'] = 'Position Size'
    apply_style(ws['A9'], StylePresets.kpi_label())
    ws['A10'] = '=IFERROR(IF(OR(B5="",B6="",B7=""),"",ROUND(B4*(B5/100)/ABS(B6-B7),2)),"")'
    apply_style(ws['A10'], StylePresets.kpi_value(COLORS['accent_gold']))
    ws.column_dimensions['A'].width = 18
    ws.column_dimensions['B'].width = 15
    protect(ws)


def build_tax_calculator_sheet(wb):
    ws = wb.create_sheet('13_Tax_Calculator')
    set_sheet_appearance(ws)
    add_title_band(ws, 'Tax Calculator', 'Brazil + generic — simplified estimate only', 1, 8)
    home_link(ws)
    ws['A4'] = 'Day Trade Monthly Net Profit'
    apply_style(ws['C4'], StylePresets.input_cell())
    ws['A5'] = 'Day Trade Tax Due (20%)'
    ws['C5'] = '=IFERROR(MAX(0,C4*0.20),"")'
    apply_style(ws['C5'], StylePresets.locked_cell())
    ws['A7'] = 'Swing Trade Monthly Net Profit'
    apply_style(ws['C7'], StylePresets.input_cell())
    ws['A8'] = 'Monthly Sales (Swing)'
    apply_style(ws['C8'], StylePresets.input_cell())
    ws['A9'] = 'Swing Trade Tax Due (15%, exempt <=20k sales)'
    ws['C9'] = '=IFERROR(IF(C8<=20000,0,MAX(0,C7*0.15)),"")'
    apply_style(ws['C9'], StylePresets.locked_cell())
    ws.merge_cells('A11:D12')
    ws['A11'] = 'This is a simplified estimate for personal organization only. Not tax advice. Consult an accountant.'
    apply_style(ws['A11'], StylePresets.warning())
    ws.column_dimensions['A'].width = 34
    ws.column_dimensions['C'].width = 16
    protect(ws)


def build_heatmaps_sheet(wb):
    ws = wb.create_sheet('16_Heatmaps')
    set_sheet_appearance(ws)
    add_title_band(ws, 'Heatmaps', 'Performance by symbol', 1, 8)
    home_link(ws)
    headers = ['Symbol', 'Trades', 'Win Rate %', 'Expectancy (R)', 'Net P&L']
    for i, h in enumerate(headers, start=1):
        apply_style(ws.cell(row=5, column=i, value=h), StylePresets.table_header())
    symbols = ['EURUSD', 'GBPUSD', 'GOLD', 'BTCUSD', 'Other']
    for r, sym in enumerate(symbols, start=6):
        ws[f'A{r}'] = sym
        ws[f'B{r}'] = f'=COUNTIF(Tbl_Journal[Symbol],A{r})'
        ws[f'C{r}'] = f'=IF(B{r}=0,"",COUNTIFS(Tbl_Journal[Symbol],A{r},Tbl_Journal[Net P&L],">0")/B{r}*100)'
        ws[f'D{r}'] = f'=IF(B{r}=0,"",AVERAGEIFS(Tbl_Journal[R-Multiple],Tbl_Journal[Symbol],A{r}))'
        ws[f'E{r}'] = f'=SUMIF(Tbl_Journal[Symbol],A{r},Tbl_Journal[Net P&L])'
        for col in 'ABCDE':
            apply_style(ws[f'{col}{r}'], StylePresets.locked_cell())
    for c in 'ABCDE':
        ws.column_dimensions[c].width = 15
    protect(ws)


# ---------------------------------------------------------------------------
# WORKBOOK ASSEMBLY PER TIER
# ---------------------------------------------------------------------------

TAB_COLORS = {
    '00_Cover': 'D4AF37', '01_How_To_Use': 'D4AF37', '02_Settings': 'D4AF37',
    '03_Dashboard': '22D3A6', '04_Trading_Journal': 'D4AF37',
    '05_Risk_Manager': 'E5484D', '06_Drawdown_Tracker': 'E5484D',
    '07_Profit_Dashboard': '22D3A6', '08_R_Multiple_Tracker': '22D3A6',
    '09_Trade_Statistics': '22D3A6', '10_Economic_Calendar': 'D4AF37',
    '11_Financial_Control': '22D3A6', '12_Lot_Size_Calculator': 'D4AF37',
    '13_Tax_Calculator': 'D4AF37', '14_Compounding_Calculator': 'D4AF37',
    '15_Expectancy_Calculator': 'D4AF37', '16_Heatmaps': '22D3A6', '17_Equity_Curve': '22D3A6',
}

BASIC_LINKS = [('01_How_To_Use', 'How to Use'), ('02_Settings', 'Settings'), ('03_Dashboard', 'Dashboard'),
               ('04_Trading_Journal', 'Trading Journal'), ('05_Risk_Manager', 'Risk Manager'),
               ('06_Drawdown_Tracker', 'Drawdown Tracker')]
PRO_EXTRA_LINKS = [('07_Profit_Dashboard', 'Profit Dashboard'), ('08_R_Multiple_Tracker', 'R-Multiple Tracker'),
                    ('09_Trade_Statistics', 'Trade Statistics'), ('14_Compounding_Calculator', 'Compounding Calculator'),
                    ('15_Expectancy_Calculator', 'Expectancy Calculator'), ('17_Equity_Curve', 'Equity Curve')]
ELITE_EXTRA_LINKS = [('10_Economic_Calendar', 'Economic Calendar'), ('11_Financial_Control', 'Financial Control'),
                      ('12_Lot_Size_Calculator', 'Lot Size Calculator'), ('13_Tax_Calculator', 'Tax Calculator'),
                      ('16_Heatmaps', 'Heatmaps')]


def build_workbook(tier: str):
    wb = Workbook()
    wb.remove(wb.active)

    links = list(BASIC_LINKS)
    if tier in ('pro', 'elite'):
        links += PRO_EXTRA_LINKS
    if tier == 'elite':
        links += ELITE_EXTRA_LINKS

    build_cover_sheet(wb, tier, links)
    build_how_to_use_sheet(wb)
    build_settings_sheet(wb)
    lists = build_lists_sheet(wb)
    build_trading_journal_sheet(wb, lists)
    build_dashboard_sheet(wb, tier)
    build_risk_manager_sheet(wb)
    build_drawdown_tracker_sheet(wb)

    if tier in ('pro', 'elite'):
        build_profit_dashboard_sheet(wb)
        build_r_multiple_tracker_sheet(wb)
        build_trade_statistics_sheet(wb)
        build_compounding_calculator_sheet(wb)
        build_expectancy_calculator_sheet(wb)
        build_equity_curve_sheet(wb)

    if tier == 'elite':
        build_economic_calendar_sheet(wb)
        build_financial_control_sheet(wb)
        build_lot_size_calculator_sheet(wb)
        build_tax_calculator_sheet(wb)
        build_heatmaps_sheet(wb)

    for name, color in TAB_COLORS.items():
        if name in wb.sheetnames:
            wb[name].sheet_properties.tabColor = color

    desired_order = ['00_Cover', '01_How_To_Use', '02_Settings', '03_Dashboard', '04_Trading_Journal',
                      '05_Risk_Manager', '06_Drawdown_Tracker', '07_Profit_Dashboard', '08_R_Multiple_Tracker',
                      '09_Trade_Statistics', '10_Economic_Calendar', '11_Financial_Control',
                      '12_Lot_Size_Calculator', '13_Tax_Calculator', '14_Compounding_Calculator',
                      '15_Expectancy_Calculator', '16_Heatmaps', '17_Equity_Curve', '99_Lists']
    wb._sheets = [wb[name] for name in desired_order if name in wb.sheetnames]
    return wb


def readme_text(tier):
    return f"""VERISCOPE EDGE - {tier.upper()} EDITION
Know Your Edge. Trade With Proof.
Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

QUICK START
1. Open 02_Settings and set your starting capital + risk % per trade.
2. Log every trade in 04_Trading_Journal.
3. 03_Dashboard updates live automatically.
4. Use 05_Risk_Manager before entering any trade.
5. Review 06_Drawdown_Tracker weekly.

Sheets are password-protected with: veriscope2026
Input cells = light blue. Formula cells = white/locked.

DISCLAIMER: Personal tracking tool only. Not tax or investment advice.
"""


def main():
    os.makedirs('output', exist_ok=True)
    for tier in ('basic', 'pro', 'elite'):
        wb = build_workbook(tier)
        path = f'output/Veriscope_Edge_{tier.capitalize()}.xlsx'
        wb.save(path)
        with open(f'output/README_HowToUse_{tier.capitalize()}.txt', 'w', encoding='utf-8') as f:
            f.write(readme_text(tier))
        print(f'Saved {path}')


if __name__ == '__main__':
    main()
