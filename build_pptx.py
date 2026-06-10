from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.chart.data import CategoryChartData

BG = RGBColor(0x0A, 0x0A, 0x0F)
ORANGE = RGBColor(0xF7, 0x93, 0x1A)
RED = RGBColor(0xFF, 0x3B, 0x30)
GREEN = RGBColor(0x30, 0xD1, 0x58)
YELLOW = RGBColor(0xFF, 0xD6, 0x0A)
PURPLE = RGBColor(0xBF, 0x5A, 0xF2)
BLUE = RGBColor(0x0A, 0x84, 0xFF)
WHITE = RGBColor(0xE0, 0xE0, 0xE0)
GRAY = RGBColor(0x88, 0x88, 0x88)
DARK = RGBColor(0x16, 0x16, 0x20)
DGRAY = RGBColor(0x44, 0x44, 0x44)

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

def set_bg(slide, color=BG):
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = color

def add_text(slide, left, top, width, height, text, size=18, color=WHITE, bold=False, align=PP_ALIGN.LEFT, font_name='Calibri'):
    txBox = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(size)
    p.font.color.rgb = color
    p.font.bold = bold
    p.font.name = font_name
    p.alignment = align
    return tf

def add_para(tf, text, size=18, color=WHITE, bold=False, space_before=Pt(6)):
    p = tf.add_paragraph()
    p.text = text
    p.font.size = Pt(size)
    p.font.color.rgb = color
    p.font.bold = bold
    p.font.name = 'Calibri'
    p.space_before = space_before
    return p

def add_rect(slide, left, top, width, height, fill_color=DARK, border_color=None):
    shape = slide.shapes.add_shape(1, Inches(left), Inches(top), Inches(width), Inches(height))
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    if border_color:
        shape.line.color.rgb = border_color
        shape.line.width = Pt(1)
    else:
        shape.line.fill.background()
    return shape

def metric_card(slide, left, top, label, value, color=WHITE, sub=''):
    add_rect(slide, left, top, 2.3, 1.4, DARK, RGBColor(0x2A, 0x2A, 0x2A))
    add_text(slide, left+0.1, top+0.1, 2.1, 0.3, label, size=9, color=GRAY, bold=False, align=PP_ALIGN.CENTER)
    add_text(slide, left+0.1, top+0.4, 2.1, 0.6, value, size=26, color=color, bold=True, align=PP_ALIGN.CENTER)
    if sub:
        add_text(slide, left+0.1, top+1.05, 2.1, 0.3, sub, size=8, color=DGRAY, align=PP_ALIGN.CENTER)

# ===================================================================
# SLIDE 1 — TITLE
# ===================================================================
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide)
add_rect(slide, 0, 0, 13.333, 7.5, RGBColor(0x0D, 0x0A, 0x00))
add_text(slide, 1, 1.5, 11.3, 1.2, 'BTC SUPPLY SHOCK', size=52, color=ORANGE, bold=True, align=PP_ALIGN.CENTER)
add_text(slide, 1, 2.7, 11.3, 0.8, 'DASHBOARD & FORECAST', size=44, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
add_text(slide, 1, 4.0, 11.3, 0.5, 'Exchange Balance Monitor  ·  STRC  ·  SATA  ·  Nuclear Supply Shock Projections', size=16, color=GRAY, align=PP_ALIGN.CENTER)
add_text(slide, 1, 5.0, 11.3, 0.4, 'Data: Glassnode  ·  STRC.live  ·  Strategy.com  ·  Strive IR  ·  SEC 8-K Filings', size=12, color=DGRAY, align=PP_ALIGN.CENTER)
add_text(slide, 1, 5.8, 11.3, 0.4, 'May 2026', size=14, color=DGRAY, align=PP_ALIGN.CENTER)

# ===================================================================
# SLIDE 2 — KEY METRICS
# ===================================================================
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide)
add_text(slide, 0.6, 0.3, 12, 0.6, 'CURRENT STATE OF BTC SUPPLY', size=28, color=ORANGE, bold=True)
add_text(slide, 0.6, 0.85, 10, 0.4, 'As of May 31, 2026 — sourced from Glassnode, SEC filings, STRC.live', size=12, color=GRAY)

metric_card(slide, 0.6, 1.6, 'BTC PRICE', '$73,751', GREEN, 'May 31, 2026')
metric_card(slide, 3.1, 1.6, 'ON EXCHANGES', '~2.97M', RED, 'Glassnode All Exchanges')
metric_card(slide, 5.6, 1.6, '% ON EXCHANGES', '14.85%', RED, 'of 20.0M mined')
metric_card(slide, 8.1, 1.6, 'STRATEGY (MSTR)', '843,738', ORANGE, '4.22% of total supply')
metric_card(slide, 10.6, 1.6, 'STRIVE (ASST)', '16,500', PURPLE, '0.08% of total supply')

add_rect(slide, 0.6, 3.4, 12.1, 0.08, ORANGE)

tf = add_text(slide, 0.6, 3.8, 5.8, 3.2, '', size=14, color=WHITE)
tf.paragraphs[0].text = 'Supply Snapshot'
tf.paragraphs[0].font.size = Pt(20)
tf.paragraphs[0].font.bold = True
tf.paragraphs[0].font.color.rgb = WHITE
add_para(tf, '', size=8)
add_para(tf, '  Total Mined Supply:          20,000,000 BTC', size=13, color=WHITE)
add_para(tf, '  Lost / Inaccessible:          ~3,500,000 BTC (17.5%)', size=13, color=DGRAY)
add_para(tf, '  Long-term Holders (1yr+):  ~8,500,000 BTC (42.5%)', size=13, color=BLUE)
add_para(tf, '  ETFs & Institutions:          ~1,100,000 BTC (5.5%)', size=13, color=YELLOW)
add_para(tf, '  Strategy (MSTR):                843,738 BTC (4.22%)', size=13, color=ORANGE)
add_para(tf, '  Strive (ASST):                      16,500 BTC (0.08%)', size=13, color=PURPLE)
add_para(tf, '  On Exchanges (liquid):     ~2,970,000 BTC (14.85%)', size=13, color=RED)
add_para(tf, '  Other Circulating:             ~3,070,000 BTC (15.35%)', size=13, color=GREEN)

tf2 = add_text(slide, 6.8, 3.8, 5.9, 3.2, '', size=14, color=WHITE)
tf2.paragraphs[0].text = 'Key Dynamics'
tf2.paragraphs[0].font.size = Pt(20)
tf2.paragraphs[0].font.bold = True
tf2.paragraphs[0].font.color.rgb = WHITE
add_para(tf2, '', size=8)
add_para(tf2, '  New miner issuance: 450 BTC/day (3,150/week)', size=13, color=WHITE)
add_para(tf2, '  Next halving: April 2028 → cuts to ~1,575/week', size=13, color=WHITE)
add_para(tf2, '  Strategy buying 3-10x miner output per week', size=13, color=ORANGE, bold=True)
add_para(tf2, '  Strive adding ~350 BTC/week and accelerating', size=13, color=PURPLE)
add_para(tf2, '  Exchange reserves at multi-year structural lows', size=13, color=RED, bold=True)
add_para(tf2, '  STRC now primary BTC acquisition vehicle', size=13, color=PURPLE)
add_para(tf2, '  Corporate demand structurally > new supply', size=13, color=YELLOW, bold=True)

# ===================================================================
# SLIDE 3 — STRATEGY WEEKLY PURCHASES
# ===================================================================
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide)
add_text(slide, 0.6, 0.3, 12, 0.6, 'STRATEGY WEEKLY BTC PURCHASES — 2026 YTD', size=28, color=ORANGE, bold=True)
add_text(slide, 0.6, 0.85, 10, 0.4, 'Breakdown: STRC Preferred Stock funding vs MSTR Equity / Other', size=12, color=GRAY)

chart_data = CategoryChartData()
chart_data.categories = ['Jan 5','Jan 12','Jan 19','Jan 26','Feb 2','Feb 9','Feb 16','Feb 23',
                          'Mar 2','Mar 9','Mar 16','Mar 23','Mar 30',
                          'Apr 6','Apr 13','Apr 20','Apr 27',
                          'May 4','May 11','May 18','May 25']

strc_vals = [0,800,4500,4200,3200,4100,1500,9800,1200,8500,10767,0,0,2100,7500,18000,1000,0,200,20000,0]
other_vals = [1468,1730,6500,5907,4433,2955,1500,10556,1815,9494,11570,0,0,2771,6427,16164,2273,0,335,4869,0]

chart_data.add_series('STRC Preferred', strc_vals)
chart_data.add_series('MSTR Equity / Other', other_vals)

chart_frame = slide.shapes.add_chart(
    XL_CHART_TYPE.COLUMN_STACKED, Inches(0.6), Inches(1.4), Inches(12.1), Inches(4.2), chart_data)
chart = chart_frame.chart
chart.has_legend = True
chart.legend.position = XL_LEGEND_POSITION.TOP
chart.legend.include_in_layout = False
chart.legend.font.size = Pt(10)
chart.legend.font.color.rgb = GRAY

plot = chart.plots[0]
plot.gap_width = 80
s0 = plot.series[0]
s0.format.fill.solid()
s0.format.fill.fore_color.rgb = PURPLE
s1 = plot.series[1]
s1.format.fill.solid()
s1.format.fill.fore_color.rgb = ORANGE

chart.chart_style = 2
cat_axis = chart.category_axis
cat_axis.tick_labels.font.size = Pt(8)
cat_axis.tick_labels.font.color.rgb = GRAY
val_axis = chart.value_axis
val_axis.tick_labels.font.size = Pt(9)
val_axis.tick_labels.font.color.rgb = GRAY
val_axis.major_gridlines.format.line.color.rgb = RGBColor(0x22, 0x22, 0x22)

add_rect(slide, 0.6, 5.85, 12.1, 1.2, RGBColor(0x18, 0x0E, 0x22), PURPLE)
tf = add_text(slide, 0.9, 5.95, 11.5, 1.0, '', size=13, color=WHITE)
tf.paragraphs[0].text = 'STRC is now the primary BTC acquisition engine.'
tf.paragraphs[0].font.bold = True
tf.paragraphs[0].font.size = Pt(14)
tf.paragraphs[0].font.color.rgb = PURPLE
add_para(tf, 'In the week of May 11-17, STRC ATM inflows hit $2.0B — funding nearly 100% of Strategy\'s 24,869 BTC purchase.', size=12, color=WHITE)
add_para(tf, 'Daily STRC ATM spiked to 14,439 BTC-equivalent on May 14.  STRC market cap: $10.5B — largest preferred stock globally.', size=12, color=GRAY)

# ===================================================================
# SLIDE 4 — STRC vs SATA
# ===================================================================
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide)
add_text(slide, 0.6, 0.3, 12, 0.6, 'STRC vs SATA — PREFERRED STOCK COMPARISON', size=28, color=ORANGE, bold=True)
add_text(slide, 0.6, 0.85, 10, 0.4, 'Two preferred equity instruments driving BTC treasury accumulation', size=12, color=GRAY)

headers = ['METRIC', 'STRC (STRATEGY)', 'SATA (STRIVE)']
rows = [
    ['Issuer', 'Strategy Inc (MSTR)', 'Strive Inc (ASST)'],
    ['Type', 'Variable Rate Perpetual Preferred', 'Variable Rate Perpetual Preferred'],
    ['Dividend Rate', '11.50% annual', '13.00% annual'],
    ['Dividend Frequency', 'Monthly', 'Daily (from Jun 16, 2026)'],
    ['Par Value', '$100', '$100'],
    ['Market Cap', '~$10.5B', '~$845M'],
    ['Notional Outstanding', '$15.5B', '—'],
    ['BTC Backing', '843,738 BTC', '16,500 BTC'],
    ['Debt Status', '$8B+ convertible notes', '$0 — debt free'],
    ['BTC Encumbered', 'Partially', '0% — fully unencumbered'],
]

col_widths = [3.5, 4.3, 4.3]
col_starts = [0.6, 4.1, 8.4]
row_h = 0.42
y_start = 1.5

for ci, h in enumerate(headers):
    add_rect(slide, col_starts[ci], y_start, col_widths[ci], row_h, RGBColor(0x1A, 0x1A, 0x28))
    add_text(slide, col_starts[ci]+0.15, y_start+0.05, col_widths[ci]-0.3, row_h, h, size=10, color=GRAY, bold=True)

for ri, row in enumerate(rows):
    y = y_start + (ri + 1) * row_h
    bg_c = DARK if ri % 2 == 0 else RGBColor(0x12, 0x12, 0x1A)
    for ci, cell in enumerate(row):
        add_rect(slide, col_starts[ci], y, col_widths[ci], row_h, bg_c)
        c = WHITE
        if ci == 1 and ('11.50' in cell or '10.5B' in cell): c = ORANGE
        if ci == 2 and ('13.00' in cell or '845M' in cell): c = PURPLE
        if 'debt free' in cell or 'unencumbered' in cell: c = GREEN
        add_text(slide, col_starts[ci]+0.15, y+0.05, col_widths[ci]-0.3, row_h, cell, size=11, color=c)

add_rect(slide, 0.6, 6.0, 5.8, 1.1, RGBColor(0x12, 0x0A, 0x00), ORANGE)
tf = add_text(slide, 0.9, 6.1, 5.4, 0.9, '', size=12, color=WHITE)
tf.paragraphs[0].text = 'STRC: The Bitcoin Machine'
tf.paragraphs[0].font.bold = True
tf.paragraphs[0].font.color.rgb = ORANGE
add_para(tf, 'Scaled to $10.5B mkt cap in 10 months. Now the', size=11, color=WHITE)
add_para(tf, 'largest preferred stock by market cap globally.', size=11, color=WHITE)
add_para(tf, 'Generates $1-2B/week in BTC acquisition capital.', size=11, color=ORANGE)

add_rect(slide, 6.8, 6.0, 5.9, 1.1, RGBColor(0x10, 0x08, 0x18), PURPLE)
tf = add_text(slide, 7.1, 6.1, 5.5, 0.9, '', size=12, color=WHITE)
tf.paragraphs[0].text = 'SATA: Clean Balance Sheet Play'
tf.paragraphs[0].font.bold = True
tf.paragraphs[0].font.color.rgb = PURPLE
add_para(tf, 'Zero debt, 100% unencumbered BTC.', size=11, color=GREEN)
add_para(tf, 'First U.S. listed security to pay daily dividends.', size=11, color=WHITE)
add_para(tf, '13% annual yield — highest in BTC preferred space.', size=11, color=PURPLE)

# ===================================================================
# SLIDE 5 — EXCHANGE BALANCE & RATE OF CHANGE
# ===================================================================
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide)
add_text(slide, 0.6, 0.3, 12, 0.6, 'EXCHANGE BALANCE — STRUCTURAL DECLINE', size=28, color=ORANGE, bold=True)
add_text(slide, 0.6, 0.85, 10, 0.4, 'Glassnode aggregate exchange balance — 2020 to present', size=12, color=GRAY)

chart_data = CategoryChartData()
dates = ['2020-01','2020-06','2020-12','2021-03','2021-06','2021-09','2021-12',
         '2022-03','2022-06','2022-09','2022-12','2023-03','2023-06','2023-09','2023-12',
         '2024-03','2024-06','2024-09','2024-12','2025-03','2025-06','2025-09','2025-12',
         '2026-01','2026-02','2026-03','2026-05']
bals = [3.10,3.05,2.90,2.80,2.65,2.55,2.60,2.55,2.50,2.45,2.40,2.38,2.35,2.30,2.35,
        2.32,2.80,2.90,3.05,3.10,3.08,3.05,3.02,3.00,2.98,2.97,2.97]
chart_data.categories = dates
chart_data.add_series('BTC on Exchanges (M)', bals)

chart_frame = slide.shapes.add_chart(
    XL_CHART_TYPE.LINE, Inches(0.6), Inches(1.4), Inches(7.5), Inches(4.5), chart_data)
chart = chart_frame.chart
chart.has_legend = False
chart.chart_style = 2
s = chart.series[0]
s.format.line.color.rgb = RED
s.format.line.width = Pt(2.5)
s.smooth = True
val_axis = chart.value_axis
val_axis.minimum_scale = 1.0
val_axis.maximum_scale = 3.5
val_axis.tick_labels.font.size = Pt(9)
val_axis.tick_labels.font.color.rgb = GRAY
val_axis.major_gridlines.format.line.color.rgb = RGBColor(0x22, 0x22, 0x22)
cat_axis = chart.category_axis
cat_axis.tick_labels.font.size = Pt(8)
cat_axis.tick_labels.font.color.rgb = GRAY

# Threshold annotations (as text boxes)
add_text(slide, 0.7, 4.9, 2.5, 0.3, '── NUCLEAR (1.0M)', size=9, color=RED)
add_text(slide, 0.7, 4.55, 2.5, 0.3, '── SEVERE (1.5M)', size=9, color=RGBColor(0xFF,0x95,0x00))
add_text(slide, 0.7, 4.15, 2.5, 0.3, '── HIGH (2.0M)', size=9, color=YELLOW)

# Right side — key points
tf = add_text(slide, 8.5, 1.4, 4.3, 4.8, '', size=14, color=WHITE)
tf.paragraphs[0].text = 'Why Exchange Balance Matters'
tf.paragraphs[0].font.size = Pt(20)
tf.paragraphs[0].font.bold = True
tf.paragraphs[0].font.color.rgb = WHITE
add_para(tf, '', size=8)
add_para(tf, 'Exchange BTC is the "float" — the only supply available for immediate purchase by new buyers.', size=13, color=WHITE)
add_para(tf, '', size=8)
add_para(tf, 'When buyers want more than sellers offer at current prices, price must rise to incentivize holders to sell.', size=13, color=WHITE)
add_para(tf, '', size=8)
add_para(tf, 'Every major BTC rally (2013, 2017, 2021) was preceded by declining exchange reserves.', size=13, color=YELLOW, bold=True)
add_para(tf, '', size=8)
add_para(tf, 'Currently at ~2.97M BTC (14.85% of supply). Trend is resuming downward as corporate treasury buying accelerates.', size=13, color=RED)
add_para(tf, '', size=8)
add_para(tf, 'Post-halving issuance (450 BTC/day) cannot replace what\'s being withdrawn.', size=13, color=GRAY)

# ===================================================================
# SLIDE 6 — DEMAND VS SUPPLY IMBALANCE
# ===================================================================
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide)
add_text(slide, 0.6, 0.3, 12, 0.6, 'THE STRUCTURAL IMBALANCE', size=28, color=ORANGE, bold=True)
add_text(slide, 0.6, 0.85, 10, 0.4, 'Corporate BTC demand vs miner new supply — multiples of weekly issuance', size=12, color=GRAY)

chart_data = CategoryChartData()
weeks = ['Mar 23','Mar 30','Apr 6','Apr 13','Apr 20','Apr 27','May 4','May 11','May 18','May 25']
strat_buy = [0,0,4871,13927,34164,3273,0,535,24869,0]
strive_buy = [350,350,350,350,350,350,350,350,350,350]
issuance = [3150]*10

chart_data.categories = weeks
chart_data.add_series('Strategy Buy', strat_buy)
chart_data.add_series('Strive Est.', strive_buy)
chart_data.add_series('New Issuance (miners)', issuance)

chart_frame = slide.shapes.add_chart(
    XL_CHART_TYPE.COLUMN_STACKED, Inches(0.6), Inches(1.4), Inches(7.5), Inches(4.0), chart_data)
chart = chart_frame.chart
chart.has_legend = True
chart.legend.position = XL_LEGEND_POSITION.TOP
chart.legend.include_in_layout = False
chart.legend.font.size = Pt(10)
chart.legend.font.color.rgb = GRAY
chart.chart_style = 2

plot = chart.plots[0]
plot.gap_width = 100
plot.series[0].format.fill.solid()
plot.series[0].format.fill.fore_color.rgb = ORANGE
plot.series[1].format.fill.solid()
plot.series[1].format.fill.fore_color.rgb = PURPLE

# Make issuance line instead
# Can't mix chart types in python-pptx easily, so add annotation
s2 = plot.series[2]
s2.format.fill.solid()
s2.format.fill.fore_color.rgb = GREEN

val_axis = chart.value_axis
val_axis.tick_labels.font.size = Pt(9)
val_axis.tick_labels.font.color.rgb = GRAY
val_axis.major_gridlines.format.line.color.rgb = RGBColor(0x22, 0x22, 0x22)
cat_axis = chart.category_axis
cat_axis.tick_labels.font.size = Pt(9)
cat_axis.tick_labels.font.color.rgb = GRAY

# Right callout
add_rect(slide, 8.5, 1.4, 4.3, 4.0, DARK, RGBColor(0x2A, 0x2A, 0x2A))
tf = add_text(slide, 8.8, 1.6, 3.7, 3.6, '', size=14, color=WHITE)
tf.paragraphs[0].text = 'The Math'
tf.paragraphs[0].font.size = Pt(20)
tf.paragraphs[0].font.bold = True
tf.paragraphs[0].font.color.rgb = ORANGE
add_para(tf, '', size=8)
add_para(tf, 'Miner Output:', size=14, color=GRAY, bold=True)
add_para(tf, '3,150 BTC / week', size=16, color=GREEN, bold=True)
add_para(tf, '', size=10)
add_para(tf, 'Strategy Peak Week:', size=14, color=GRAY, bold=True)
add_para(tf, '34,164 BTC (10.8x issuance)', size=16, color=ORANGE, bold=True)
add_para(tf, '', size=10)
add_para(tf, 'Strategy Avg Active Week:', size=14, color=GRAY, bold=True)
add_para(tf, '~14,000 BTC (4.4x issuance)', size=16, color=ORANGE, bold=True)
add_para(tf, '', size=10)
add_para(tf, 'This is structurally', size=14, color=WHITE)
add_para(tf, 'unprecedented.', size=14, color=RED, bold=True)

add_rect(slide, 0.6, 5.7, 12.1, 1.3, RGBColor(0x1A, 0x0A, 0x0A), RED)
tf = add_text(slide, 0.9, 5.8, 11.5, 1.1, '', size=13, color=WHITE)
tf.paragraphs[0].text = 'Key Insight'
tf.paragraphs[0].font.bold = True
tf.paragraphs[0].font.color.rgb = RED
tf.paragraphs[0].font.size = Pt(16)
add_para(tf, 'Strategy alone absorbs 3-10x more BTC per week than miners produce. Add Strive, ETF inflows, other corporate treasuries,', size=13, color=WHITE)
add_para(tf, 'and organic self-custody outflows — the net drain on exchange-available supply is accelerating faster than at any point in Bitcoin\'s history.', size=13, color=WHITE)

# ===================================================================
# SLIDE 7 — SUPPLY SHOCK THRESHOLDS
# ===================================================================
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide)
add_text(slide, 0.6, 0.3, 12, 0.6, 'SUPPLY SHOCK SEVERITY LEVELS', size=28, color=ORANGE, bold=True)
add_text(slide, 0.6, 0.85, 10, 0.4, 'What happens when exchange supply drops below critical thresholds', size=12, color=GRAY)

levels = [
    ('ELEVATED', '< 2.5M BTC (12.5%)', YELLOW, 'Reduced liquidity. Larger orders start causing slippage. Price becomes more sensitive to demand spikes. Early warning.'),
    ('HIGH', '< 2.0M BTC (10.0%)', RGBColor(0xFF,0x95,0x00), 'Significant supply constraint. OTC desks struggle to fill large orders. Bid-ask spreads widen. Institutional buyers compete for shrinking float.'),
    ('SEVERE', '< 1.5M BTC (7.5%)', RGBColor(0xFF,0x55,0x30), 'Critical supply shortage. Exchange order books thin dramatically. Small purchases create outsized price moves. Resembles 2017 pre-parabolic conditions.'),
    ('NUCLEAR', '< 1.0M BTC (5.0%)', RED, 'Unprecedented territory. Functional supply crisis. Sellers must be incentivized at dramatically higher prices. Potential for parabolic, reflexive price discovery as supply vanishes.'),
]

for i, (label, threshold, color, desc) in enumerate(levels):
    y = 1.6 + i * 1.4
    add_rect(slide, 0.6, y, 12.1, 1.2, DARK, color)
    add_text(slide, 0.9, y + 0.1, 2.0, 0.4, label, size=20, color=color, bold=True)
    add_text(slide, 0.9, y + 0.5, 2.5, 0.3, threshold, size=12, color=GRAY)
    add_text(slide, 3.5, y + 0.15, 8.9, 0.9, desc, size=13, color=WHITE)

add_text(slide, 0.6, 7.0, 12, 0.4, 'Current: 2.97M BTC on exchanges (14.85%) — Score: 25/100 — just entered the descent path toward "Elevated" territory', size=12, color=YELLOW, bold=True, align=PP_ALIGN.CENTER)

# ===================================================================
# SLIDE 8 — FORECAST SCENARIOS
# ===================================================================
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide)
add_text(slide, 0.6, 0.3, 12, 0.6, 'NUCLEAR SUPPLY SHOCK FORECAST', size=28, color=ORANGE, bold=True)
add_text(slide, 0.6, 0.85, 10, 0.4, 'Four scenario projections — when does exchange balance hit critical levels?', size=12, color=GRAY)

# Table
scenarios = [
    ('Conservative', '5,600', '2,290', '115 wks', 'Aug 2028', '< 2.0M', 'HIGH', YELLOW),
    ('Base Case', '14,000', '13,200', '86 wks', 'Jan 2028', '< 1.0M', 'NUCLEAR', RED),
    ('Bull Case', '21,000', '21,825', '63 wks', 'Aug 2027', '< 1.0M', 'NUCLEAR', RED),
    ('Hyper Bull', '35,000', '39,475', '41 wks', 'Mar 2027', '< 1.0M', 'NUCLEAR', RED),
]

sh = ['SCENARIO', 'STRATEGY BUY/WK', 'NET DRAIN/WK', 'TIME TO SHOCK', 'PROJECTED DATE', 'EXCHANGE BAL', 'SEVERITY']
col_w = [2.0, 1.8, 1.7, 1.5, 1.7, 1.6, 1.6]
cx = [0.6]
for w in col_w[:-1]:
    cx.append(cx[-1] + w)

y0 = 1.5
for ci, h in enumerate(sh):
    add_rect(slide, cx[ci], y0, col_w[ci], 0.4, RGBColor(0x1A, 0x1A, 0x28))
    add_text(slide, cx[ci]+0.1, y0+0.05, col_w[ci]-0.2, 0.3, h, size=8, color=GRAY, bold=True)

for ri, (name, buy, drain, weeks, date, bal, sev, sev_c) in enumerate(scenarios):
    y = y0 + (ri + 1) * 0.55
    bg_c = DARK if ri % 2 == 0 else RGBColor(0x12, 0x12, 0x1A)
    vals = [name, buy + ' BTC', drain + ' BTC', weeks, date, bal, sev]
    for ci, v in enumerate(vals):
        add_rect(slide, cx[ci], y, col_w[ci], 0.5, bg_c)
        c = WHITE
        if ci == 0: c = WHITE; bold = True
        elif ci == 6: c = sev_c; bold = True
        else: c = WHITE; bold = False
        add_text(slide, cx[ci]+0.1, y+0.1, col_w[ci]-0.2, 0.35, v, size=12, color=c, bold=(ci==0 or ci==6))

# Assumptions below
tf = add_text(slide, 0.6, 4.2, 12, 3.0, '', size=11, color=GRAY)
tf.paragraphs[0].text = 'Model Assumptions'
tf.paragraphs[0].font.size = Pt(16)
tf.paragraphs[0].font.color.rgb = DGRAY
tf.paragraphs[0].font.bold = True
add_para(tf, '', size=6)
add_para(tf, '  Starting exchange balance: 2,970,000 BTC  •  New issuance: 3,150 BTC/week  •  Organic outflow: 4,000 BTC/week  •  Inflow: 2,500 BTC/week', size=10, color=DGRAY)
add_para(tf, '  Strategy buy acceleration: 5%/month compounding  •  Strive: 350 BTC/week  •  New corporate entrants: 500 BTC/week', size=10, color=DGRAY)
add_para(tf, '  "Nuclear" threshold: exchange balance < 1,000,000 BTC (5% of supply)  •  Does NOT account for: ETF flows, sovereign purchases, demand destruction from price', size=10, color=DGRAY)
add_para(tf, '', size=10)
add_para(tf, '  Conservative scenario: Strategy buys at 40% of current rate, organic outflows halved', size=10, color=GREEN)
add_para(tf, '  Base Case: Current rates sustained with 5%/mo acceleration', size=10, color=YELLOW)
add_para(tf, '  Bull Case: 1.5x current rates, 1.3x organic outflows', size=10, color=RGBColor(0xFF,0x95,0x00))
add_para(tf, '  Hyper Bull: 2.5x current rates, 2x organic outflows — Strategy hits 1M BTC target + new entrants', size=10, color=RED)

add_rect(slide, 0.6, 6.7, 12.1, 0.5, RGBColor(0x1A, 0x0A, 0x0A))
add_text(slide, 0.9, 6.75, 11.5, 0.4, 'Even the conservative case projects exchange balance hitting "High" shock territory by mid-2028. The base case reaches nuclear by early 2028.', size=13, color=RED, bold=True)

# ===================================================================
# SLIDE 9 — TIMELINE VISUALIZATION
# ===================================================================
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide)
add_text(slide, 0.6, 0.3, 12, 0.6, 'PROJECTED TIMELINE TO NUCLEAR SUPPLY SHOCK', size=28, color=ORANGE, bold=True)
add_text(slide, 0.6, 0.85, 10, 0.4, 'When does exchange balance cross the 1M BTC nuclear threshold under each scenario?', size=12, color=GRAY)

# Timeline arrow
add_rect(slide, 0.6, 3.3, 12.1, 0.06, GRAY)

markers = [
    ('NOW\nJun 2026', 0.6, WHITE, '2.97M BTC'),
    ('Hyper Bull\nMar 2027', 3.5, RED, '41 weeks'),
    ('Bull Case\nAug 2027', 5.5, RGBColor(0xFF,0x95,0x00), '63 weeks'),
    ('Base Case\nJan 2028', 7.9, YELLOW, '86 weeks'),
    ('Conservative\nAug 2028', 10.5, GREEN, '115 weeks'),
]

for label, x, color, sub in markers:
    add_rect(slide, x, 2.9, 0.08, 0.9, color)
    lines = label.split('\n')
    add_text(slide, x-0.8, 2.0, 1.7, 0.8, label, size=12, color=color, bold=True, align=PP_ALIGN.CENTER)
    add_text(slide, x-0.8, 3.55, 1.7, 0.3, sub, size=10, color=DGRAY, align=PP_ALIGN.CENTER)

# Halving marker
add_rect(slide, 9.2, 2.9, 0.06, 0.9, BLUE)
add_text(slide, 8.4, 3.55, 1.7, 0.5, 'HALVING\nApr 2028', size=10, color=BLUE, align=PP_ALIGN.CENTER)

# Bottom narrative
add_rect(slide, 0.6, 4.8, 5.8, 2.2, DARK, RGBColor(0x2A, 0x2A, 0x2A))
tf = add_text(slide, 0.9, 4.95, 5.2, 2.0, '', size=13, color=WHITE)
tf.paragraphs[0].text = 'Accelerating Factors'
tf.paragraphs[0].font.size = Pt(18)
tf.paragraphs[0].font.bold = True
tf.paragraphs[0].font.color.rgb = RED
add_para(tf, '', size=6)
add_para(tf, '  Strategy targeting 1M BTC (~6,158/wk needed)', size=12, color=ORANGE)
add_para(tf, '  STRC ATM scaling — $2B/week and growing', size=12, color=PURPLE)
add_para(tf, '  Strive SATA launching daily dividends Jun 16', size=12, color=PURPLE)
add_para(tf, '  Other corporate treasuries entering the space', size=12, color=WHITE)
add_para(tf, '  ETF inflows resume in risk-on environment', size=12, color=YELLOW)
add_para(tf, '  April 2028 halving cuts issuance by 50%', size=12, color=BLUE)

add_rect(slide, 6.8, 4.8, 5.9, 2.2, DARK, RGBColor(0x2A, 0x2A, 0x2A))
tf = add_text(slide, 7.1, 4.95, 5.3, 2.0, '', size=13, color=WHITE)
tf.paragraphs[0].text = 'Decelerating Factors'
tf.paragraphs[0].font.size = Pt(18)
tf.paragraphs[0].font.bold = True
tf.paragraphs[0].font.color.rgb = GREEN
add_para(tf, '', size=6)
add_para(tf, '  Rising BTC price increases acquisition cost', size=12, color=WHITE)
add_para(tf, '  Strategy debt load ($8B+) limits borrowing', size=12, color=WHITE)
add_para(tf, '  Potential regulatory intervention', size=12, color=WHITE)
add_para(tf, '  Market sell-offs / risk-off environments', size=12, color=WHITE)
add_para(tf, '  STRC/SATA dividend obligations ($1B+/yr)', size=12, color=WHITE)
add_para(tf, '  Profit-taking at higher price levels', size=12, color=WHITE)

# ===================================================================
# SLIDE 10 — CONCLUSION
# ===================================================================
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide)
add_rect(slide, 0, 0, 13.333, 7.5, RGBColor(0x0D, 0x0A, 0x00))
add_text(slide, 1, 0.8, 11.3, 0.8, 'CONCLUSION', size=36, color=ORANGE, bold=True, align=PP_ALIGN.CENTER)

points = [
    ('Exchange supply is in structural decline', 'Only 14.85% of BTC remains on exchanges — down from 16%+ in 2022. The trend is resuming downward.', RED),
    ('Corporate demand exceeds new supply by multiples', 'Strategy alone buys 3-10x what miners produce weekly. STRC preferred stock has industrialized BTC acquisition.', ORANGE),
    ('Two preferred stock instruments are accelerating the drain', 'STRC ($10.5B mkt cap) and SATA (13% daily dividend) represent a new financial primitive — equity-to-BTC conversion engines.', PURPLE),
    ('Nuclear supply shock projected within 12-24 months', 'Base case: Jan 2028. Bull case: Aug 2027. Even conservative estimates show critical levels by mid-2028.', YELLOW),
    ('The April 2028 halving is a force multiplier', 'Issuance drops from 3,150 to ~1,575 BTC/week — making the supply deficit even more acute.', BLUE),
]

for i, (title, body, color) in enumerate(points):
    y = 2.0 + i * 1.05
    add_rect(slide, 1.0, y, 11.3, 0.9, DARK, RGBColor(0x22, 0x22, 0x22))
    add_text(slide, 1.3, y + 0.08, 10.7, 0.35, title, size=16, color=color, bold=True)
    add_text(slide, 1.3, y + 0.45, 10.7, 0.4, body, size=11, color=GRAY)

add_text(slide, 1, 7.0, 11.3, 0.3, 'Sources: Glassnode  •  STRC.live  •  Strategy.com  •  Strive IR  •  SEC 8-K Filings  •  CoinDesk  •  Bitbo', size=9, color=DGRAY, align=PP_ALIGN.CENTER)

# ===================================================================
# SAVE
# ===================================================================
out = '/home/user/btc-supply-dashboard/BTC_Supply_Shock_Dashboard.pptx'
prs.save(out)
print(f'Saved to {out}')
