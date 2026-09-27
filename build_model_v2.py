"""
Aperture Ads COGS Model v2 — adds:
- Overhead/management roles (ops mgr, onboarding specialist, QA, AM, sales, analytics)
- Trigger-based scaling formulas (hires added as client count crosses thresholds)
- New "Cost Per Client" tab showing fulfillment $/client + overhead $/client + total $/client
- Updated Team Roadmap reflecting the layered hiring plan

Overhead roles model: each has monthly cost, trigger client count (first hire),
and scaling ratio (add 1 hire per additional N clients). Formula:
  headcount = IF(clients >= trigger, 1 + FLOOR((clients-trigger)/scaling), 0)
"""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

wb = Workbook()

HEADER = Font(bold=True, color="FFFFFF", size=11)
HEADER_FILL = PatternFill("solid", fgColor="1F3864")
SECTION = Font(bold=True, size=12, color="1F3864")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
CALC_FILL = PatternFill("solid", fgColor="E8EEF7")
TOTAL_FILL = PatternFill("solid", fgColor="D9E1F2")
OVERHEAD_FILL = PatternFill("solid", fgColor="FCE4D6")
BOLD = Font(bold=True)
THIN = Side(border_style="thin", color="B0B0B0")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal="center", vertical="center")
LEFT = Alignment(horizontal="left", vertical="center")
RIGHT = Alignment(horizontal="right", vertical="center")

MONEY = '"$"#,##0'
MONEY2 = '"$"#,##0.00'
PCT = '0%'
PCT1 = '0.0%'
INT = '#,##0'

def hdr(cell, text):
    cell.value = text
    cell.font = HEADER
    cell.fill = HEADER_FILL
    cell.alignment = CENTER

def section(ws, row, text):
    c = ws.cell(row=row, column=1, value=text)
    c.font = SECTION

def input_cell(cell, value, fmt=None):
    cell.value = value
    cell.fill = INPUT_FILL
    cell.alignment = RIGHT
    if fmt:
        cell.number_format = fmt

def calc_cell(cell, formula, fmt=None, bold=False, fill=None):
    cell.value = formula
    cell.fill = fill if fill else CALC_FILL
    cell.alignment = RIGHT
    if fmt:
        cell.number_format = fmt
    if bold:
        cell.font = BOLD

def label(cell, text, bold=False, indent=0):
    cell.value = ("  " * indent) + text
    cell.alignment = LEFT
    if bold:
        cell.font = BOLD

# ==============================================================
# TAB 1: ASSUMPTIONS
# ==============================================================
ws = wb.active
ws.title = "Assumptions"
ws.column_dimensions['A'].width = 30
for col in ['B','C','D','E','F','G','H']:
    ws.column_dimensions[col].width = 16

ws.cell(row=1, column=1, value="APERTURE ADS — COGS MODEL v2 (with overhead roles)").font = Font(bold=True, size=14, color="1F3864")
ws.cell(row=2, column=1, value="Yellow = editable input. Blue = calculated. Orange rows on Scale Model = overhead/mgmt layer.").font = Font(italic=True, color="808080")

# ---- FULFILLMENT ROLES ----
section(ws, 4, "FULFILLMENT ROLES")
for i, h in enumerate(["Role", "Monthly Cost", "Monthly Hours", "Aperture Allocation %", "Aperture Monthly Cost", "Effective $/hr"]):
    hdr(ws.cell(row=5, column=1+i), h)

roles = [
    ("Media Buyer",  780, 160, 1.00),
    ("Designer",     440,  80, 1.00),
    ("Assistant",    700, 160, 0.90),
    ("Copywriter",   500,  80, 1.00),
]
for i, (name, cost, hrs, alloc) in enumerate(roles):
    r = 6 + i
    label(ws.cell(row=r, column=1), name)
    input_cell(ws.cell(row=r, column=2), cost, MONEY)
    input_cell(ws.cell(row=r, column=3), hrs, INT)
    input_cell(ws.cell(row=r, column=4), alloc, PCT)
    calc_cell(ws.cell(row=r, column=5), f"=B{r}*D{r}", MONEY)
    calc_cell(ws.cell(row=r, column=6), f"=IF(C{r}=0,0,E{r}/C{r})", MONEY2)

BUYER_HR_RATE = "Assumptions!$F$6"
DESIGNER_HR_RATE = "Assumptions!$F$7"
ASSIST_HR_RATE = "Assumptions!$F$8"
COPY_HR_RATE = "Assumptions!$F$9"

BUYER_MO_COST = "Assumptions!$E$6"
DESIGNER_MO_COST = "Assumptions!$E$7"
ASSIST_MO_COST = "Assumptions!$E$8"
COPY_MO_COST = "Assumptions!$E$9"

BUYER_CAP = "Assumptions!$C$6"
DESIGNER_CAP = "Assumptions!$C$7"
ASSIST_CAP = "(Assumptions!$C$8*Assumptions!$D$8)"
COPY_CAP = "Assumptions!$C$9"

# ---- OVERHEAD / MANAGEMENT ROLES ----
section(ws, 12, "OVERHEAD / MANAGEMENT ROLES (kick in at trigger)")
for i, h in enumerate(["Role", "Monthly Cost", "Trigger (rental clients)", "1 hire per N clients", "Notes"]):
    hdr(ws.cell(row=13, column=1+i), h)

# Order: Account Mgr (first trigger), Onboarding Spec, Ops Mgr, QA, Analytics, Sales
overhead_roles = [
    ("Account Manager",       1500, 10, 15, "Owns client relationship, weekly reports, retention"),
    ("Onboarding Specialist", 1200, 15, 30, "Full-time on new-client intake + setup (replaces Matt's onboarding time)"),
    ("Ops Manager / Team Lead", 2000, 20, 40, "Supervises buyers + designers; QA workflow; 1 per team of ~4-5 buyers"),
    ("QA / Ad Reviewer",       1000, 20, 40, "Independent review of creative + campaigns before launch"),
    ("Analytics / Reporting",  2000, 30, 60, "Automates dashboards, per-client reports at scale"),
    ("Sales / BDR",            2500, 30, 60, "Proactive outreach beyond partner network (optional if partner flow strong)"),
]
for i, (name, cost, trigger, scaling, notes) in enumerate(overhead_roles):
    r = 14 + i
    label(ws.cell(row=r, column=1), name)
    input_cell(ws.cell(row=r, column=2), cost, MONEY)
    input_cell(ws.cell(row=r, column=3), trigger, INT)
    input_cell(ws.cell(row=r, column=4), scaling, INT)
    c = ws.cell(row=r, column=5, value=notes)
    c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    c.font = Font(italic=True, color="606060")
ws.column_dimensions['E'].width = 60

# rows 14-19 for overhead roles
OVERHEAD_START = 14
OVERHEAD_END = 19

# ---- PER-CLIENT FULFILLMENT HOURS ----
section(ws, 22, "PER-CLIENT FULFILLMENT HOURS (monthly, per client)")
for i, h in enumerate(["Tier", "Buyer hrs", "Designer hrs", "Assistant hrs", "Copywriter hrs"]):
    hdr(ws.cell(row=23, column=1+i), h)

tiers = [
    ("Excire (custom high-touch)", 64, 32, 0, 0),
    ("Maven (custom mid)",         64, 32, 0, 0),
    ("Nash (lead gen light)",      32, 16, 0, 0),
    ("Rental (productized)",       15,  8, 6, 4),
]
for i, (name, b, d, a, c) in enumerate(tiers):
    r = 24 + i
    label(ws.cell(row=r, column=1), name)
    input_cell(ws.cell(row=r, column=2), b, INT)
    input_cell(ws.cell(row=r, column=3), d, INT)
    input_cell(ws.cell(row=r, column=4), a, INT)
    input_cell(ws.cell(row=r, column=5), c, INT)

EXCIRE_HRS, MAVEN_HRS, NASH_HRS, RENTAL_HRS = 24, 25, 26, 27

# ---- PRICING ----
section(ws, 30, "PRICING")
for i, h in enumerate(["Tier", "Gross Fee ($/mo)", "Partner Split %", "Net Revenue ($/mo)"]):
    hdr(ws.cell(row=31, column=1+i), h)

prices = [
    ("Excire", 2500, 0.00),
    ("Maven",  1200, 0.00),
    ("Nash",    500, 0.00),
    ("Rental", 2500, 0.50),
]
for i, (name, fee, split) in enumerate(prices):
    r = 32 + i
    label(ws.cell(row=r, column=1), name)
    input_cell(ws.cell(row=r, column=2), fee, MONEY)
    input_cell(ws.cell(row=r, column=3), split, PCT)
    calc_cell(ws.cell(row=r, column=4), f"=B{r}*(1-C{r})", MONEY, bold=True)

EXCIRE_PRICE, MAVEN_PRICE, NASH_PRICE, RENTAL_PRICE = 32, 33, 34, 35

# ---- ONBOARDING ----
section(ws, 38, "ONBOARDING (one-time hours per new client)")
label(ws.cell(row=39, column=1), "Assistant hours per onboarding")
input_cell(ws.cell(row=39, column=2), 10, INT)
label(ws.cell(row=40, column=1), "Buyer hours per onboarding")
input_cell(ws.cell(row=40, column=2), 3, INT)
label(ws.cell(row=41, column=1), "Onboarding cost per client")
calc_cell(ws.cell(row=41, column=2), f"=B39*{ASSIST_HR_RATE}+B40*{BUYER_HR_RATE}", MONEY2, bold=True)

ONBOARD_ASSIST_HRS = "Assumptions!$B$39"
ONBOARD_BUYER_HRS = "Assumptions!$B$40"
ONBOARD_COST_PER_CLIENT = "Assumptions!$B$41"

# ---- SCALE SCENARIOS ----
section(ws, 44, "SCALE SCENARIOS (rental client counts to model)")
for i, (name, count) in enumerate([("Current (0 rental)",0), ("+10 rental",10), ("+30 rental",30), ("+60 rental",60), ("+90 rental",90)]):
    r = 45 + i
    label(ws.cell(row=r, column=1), f"  {name}")
    input_cell(ws.cell(row=r, column=2), count, INT)

SCENARIO_ROWS = [45, 46, 47, 48, 49]

# ==============================================================
# TAB 2: CURRENT STATE (unchanged from v1)
# ==============================================================
ws2 = wb.create_sheet("Current State")
ws2.column_dimensions['A'].width = 26
for col in ['B','C','D','E','F','G']:
    ws2.column_dimensions[col].width = 16

ws2.cell(row=1, column=1, value="CURRENT STATE — 3 clients today").font = Font(bold=True, size=14, color="1F3864")
ws2.cell(row=2, column=1, value="Assistant + Copywriter costs are NOT loaded here — those roles are still on DIYP/Brownstone until rental clients launch.").font = Font(italic=True, color="808080")

for i, h in enumerate(["Client", "Monthly Fee", "Buyer hrs", "Designer hrs", "Direct Labor Cost", "Gross Profit"]):
    hdr(ws2.cell(row=4, column=1+i), h)

for i, (name, hrs_row, price_row) in enumerate([("Excire", EXCIRE_HRS, EXCIRE_PRICE), ("Maven", MAVEN_HRS, MAVEN_PRICE), ("Nash", NASH_HRS, NASH_PRICE)]):
    r = 5 + i
    label(ws2.cell(row=r, column=1), name)
    calc_cell(ws2.cell(row=r, column=2), f"=Assumptions!B{price_row}", MONEY)
    calc_cell(ws2.cell(row=r, column=3), f"=Assumptions!B{hrs_row}", INT)
    calc_cell(ws2.cell(row=r, column=4), f"=Assumptions!C{hrs_row}", INT)
    calc_cell(ws2.cell(row=r, column=5), f"=C{r}*{BUYER_HR_RATE}+D{r}*{DESIGNER_HR_RATE}", MONEY2)
    calc_cell(ws2.cell(row=r, column=6), f"=B{r}-E{r}", MONEY)

r = 8
label(ws2.cell(row=r, column=1), "TOTAL", bold=True)
for col in range(2, 7):
    c = ws2.cell(row=r, column=col)
    c.value = f"=SUM({get_column_letter(col)}5:{get_column_letter(col)}7)"
    c.fill = TOTAL_FILL
    c.font = BOLD
    c.border = BORDER
    c.alignment = RIGHT
    c.number_format = MONEY if col in (2,5,6) else INT

section(ws2, 11, "MONTHLY P&L (current state)")
def kv_row(ws, r, lbl, formula, fmt=MONEY, bold=False):
    label(ws.cell(row=r, column=1), lbl, bold=bold)
    c = ws.cell(row=r, column=2)
    c.value = formula
    c.fill = CALC_FILL if not bold else TOTAL_FILL
    c.border = BORDER
    c.alignment = RIGHT
    c.number_format = fmt
    if bold:
        c.font = BOLD

kv_row(ws2, 12, "Total Revenue", "=B8")
kv_row(ws2, 13, "Media Buyer cost (loaded to Aperture)", f"={BUYER_MO_COST}")
kv_row(ws2, 14, "Designer cost (loaded to Aperture)", f"={DESIGNER_MO_COST}")
kv_row(ws2, 15, "Assistant cost (still on DIYP today)", "=0")
kv_row(ws2, 16, "Copywriter cost (still on Brownstone today)", "=0")
kv_row(ws2, 17, "Total COGS (loaded)", "=SUM(B13:B16)", bold=True)
kv_row(ws2, 18, "Gross Profit", "=B12-B17", bold=True)
kv_row(ws2, 19, "Gross Margin %", "=B18/B12", fmt=PCT1, bold=True)

ws2.cell(row=21, column=1, value="⚠ Buyer + Designer are already ~100% utilized. Any additional client requires more capacity.").font = Font(italic=True, color="C00000")

# ==============================================================
# TAB 3: UNIT ECONOMICS (unchanged — fulfillment only, per-tier)
# ==============================================================
ws3 = wb.create_sheet("Unit Economics")
ws3.column_dimensions['A'].width = 26
for col in ['B','C','D','E','F','G','H','I']:
    ws3.column_dimensions[col].width = 14

ws3.cell(row=1, column=1, value="UNIT ECONOMICS — FULFILLMENT cost to serve ONE client of each tier").font = Font(bold=True, size=14, color="1F3864")
ws3.cell(row=2, column=1, value="Direct fulfillment cost only. Overhead cost per client shown on Cost Per Client tab (varies with scale).").font = Font(italic=True, color="808080")

for i, h in enumerate(["Tier", "Gross Fee", "Buyer $", "Designer $", "Assistant $", "Copywriter $", "Partner Split", "Fulfillment GP / mo", "GM %"]):
    hdr(ws3.cell(row=4, column=1+i), h)

for i, (name, hrs_row, price_row) in enumerate([("Excire", EXCIRE_HRS, EXCIRE_PRICE), ("Maven", MAVEN_HRS, MAVEN_PRICE), ("Nash", NASH_HRS, NASH_PRICE), ("Rental", RENTAL_HRS, RENTAL_PRICE)]):
    r = 5 + i
    label(ws3.cell(row=r, column=1), name)
    calc_cell(ws3.cell(row=r, column=2), f"=Assumptions!B{price_row}", MONEY)
    calc_cell(ws3.cell(row=r, column=3), f"=Assumptions!B{hrs_row}*{BUYER_HR_RATE}", MONEY2)
    calc_cell(ws3.cell(row=r, column=4), f"=Assumptions!C{hrs_row}*{DESIGNER_HR_RATE}", MONEY2)
    calc_cell(ws3.cell(row=r, column=5), f"=Assumptions!D{hrs_row}*{ASSIST_HR_RATE}", MONEY2)
    calc_cell(ws3.cell(row=r, column=6), f"=Assumptions!E{hrs_row}*{COPY_HR_RATE}", MONEY2)
    calc_cell(ws3.cell(row=r, column=7), f"=Assumptions!B{price_row}*Assumptions!C{price_row}", MONEY2)
    calc_cell(ws3.cell(row=r, column=8), f"=B{r}-SUM(C{r}:G{r})", MONEY, bold=True)
    calc_cell(ws3.cell(row=r, column=9), f"=H{r}/B{r}", PCT1, bold=True)

section(ws3, 11, "KEY TAKEAWAYS (fulfillment only)")
notes = [
    "• Excire is your highest fulfillment-GP-per-client account.",
    "• Nash's fee barely covers direct labor — reconsider at renewal.",
    "• Rental clients: ~$1,082 fulfillment GP/mo after partner split (before overhead allocation).",
    "• Cost Per Client tab shows the FULLY LOADED picture including overhead layer.",
]
for i, n in enumerate(notes):
    ws3.cell(row=12+i, column=1, value=n).alignment = LEFT

# ==============================================================
# TAB 4: SCALE MODEL — now with overhead layer
# ==============================================================
ws4 = wb.create_sheet("Scale Model")
ws4.column_dimensions['A'].width = 34
for col in ['B','C','D','E','F']:
    ws4.column_dimensions[col].width = 16

ws4.cell(row=1, column=1, value="SCALE MODEL — steady-state monthly economics at each rental client count").font = Font(bold=True, size=14, color="1F3864")
ws4.cell(row=2, column=1, value="Fulfillment (blue) + Overhead/Mgmt (orange). Onboarding cost = one-time investment during ramp.").font = Font(italic=True, color="808080")

scen_names = ["Current", "+10 Rental", "+30 Rental", "+60 Rental", "+90 Rental"]
hdr(ws4.cell(row=4, column=1), "Metric")
for i, name in enumerate(scen_names):
    hdr(ws4.cell(row=4, column=2+i), name)

scen_cols = ['B','C','D','E','F']
scen_ref = [f"Assumptions!$B${SCENARIO_ROWS[i]}" for i in range(5)]

def row_across(ws, r, label_text, formulas, fmt=None, bold=False, fill=None, indent=0):
    label(ws.cell(row=r, column=1), label_text, bold=bold, indent=indent)
    for i, f in enumerate(formulas):
        c = ws.cell(row=r, column=2+i)
        c.value = f
        c.fill = fill if fill else (TOTAL_FILL if bold else CALC_FILL)
        c.border = BORDER
        c.alignment = RIGHT
        if fmt:
            c.number_format = fmt
        if bold:
            c.font = BOLD

r = 5
label(ws4.cell(row=r, column=1), "Rental clients")
for i in range(5):
    calc_cell(ws4.cell(row=r, column=2+i), f"={scen_ref[i]}", INT, bold=True)
r = 6
label(ws4.cell(row=r, column=1), "Current clients (Excire, Maven, Nash)")
for i in range(5):
    calc_cell(ws4.cell(row=r, column=2+i), "=3", INT)

# REVENUE
section(ws4, 8, "REVENUE")
row_across(ws4, 9, "Current-client revenue", [f"=Assumptions!B{EXCIRE_PRICE}+Assumptions!B{MAVEN_PRICE}+Assumptions!B{NASH_PRICE}"]*5, MONEY)
row_across(ws4, 10, "Rental gross revenue", [f"={scen_ref[i]}*Assumptions!B{RENTAL_PRICE}" for i in range(5)], MONEY)
row_across(ws4, 11, "Less: partner split (rental only)", [f"=-{scen_ref[i]}*Assumptions!B{RENTAL_PRICE}*Assumptions!C{RENTAL_PRICE}" for i in range(5)], MONEY)
row_across(ws4, 12, "Net Revenue to Aperture", [f"={scen_cols[i]}9+{scen_cols[i]}10+{scen_cols[i]}11" for i in range(5)], MONEY, bold=True)

# FULFILLMENT HOURS
section(ws4, 14, "FULFILLMENT LABOR — hours required")
row_across(ws4, 15, "Buyer hrs needed",
    [f"=(Assumptions!B{EXCIRE_HRS}+Assumptions!B{MAVEN_HRS}+Assumptions!B{NASH_HRS})+{scen_ref[i]}*Assumptions!B{RENTAL_HRS}" for i in range(5)], INT)
row_across(ws4, 16, "Designer hrs needed",
    [f"=(Assumptions!C{EXCIRE_HRS}+Assumptions!C{MAVEN_HRS}+Assumptions!C{NASH_HRS})+{scen_ref[i]}*Assumptions!C{RENTAL_HRS}" for i in range(5)], INT)
row_across(ws4, 17, "Assistant hrs needed",
    [f"={scen_ref[i]}*Assumptions!D{RENTAL_HRS}" for i in range(5)], INT)
row_across(ws4, 18, "Copywriter hrs needed",
    [f"={scen_ref[i]}*Assumptions!E{RENTAL_HRS}" for i in range(5)], INT)

# FULFILLMENT HEADCOUNT
section(ws4, 20, "FULFILLMENT HEADCOUNT (rounded up)")
row_across(ws4, 21, "Media Buyers", [f"=CEILING({scen_cols[i]}15/{BUYER_CAP},1)" for i in range(5)], INT, bold=True)
row_across(ws4, 22, "Designers", [f"=CEILING({scen_cols[i]}16/{DESIGNER_CAP},1)" for i in range(5)], INT, bold=True)
row_across(ws4, 23, "Assistants", [f"=IF({scen_cols[i]}17=0,0,CEILING({scen_cols[i]}17/{ASSIST_CAP},1))" for i in range(5)], INT, bold=True)
row_across(ws4, 24, "Copywriters", [f"=IF({scen_cols[i]}18=0,0,CEILING({scen_cols[i]}18/{COPY_CAP},1))" for i in range(5)], INT, bold=True)

# FULFILLMENT LABOR COST
section(ws4, 26, "FULFILLMENT LABOR COST")
row_across(ws4, 27, "Media Buyer cost", [f"={scen_cols[i]}21*{BUYER_MO_COST}" for i in range(5)], MONEY)
row_across(ws4, 28, "Designer cost", [f"={scen_cols[i]}22*{DESIGNER_MO_COST}" for i in range(5)], MONEY)
row_across(ws4, 29, "Assistant cost", [f"={scen_cols[i]}23*{ASSIST_MO_COST}" for i in range(5)], MONEY)
row_across(ws4, 30, "Copywriter cost", [f"={scen_cols[i]}24*{COPY_MO_COST}" for i in range(5)], MONEY)
row_across(ws4, 31, "Total Fulfillment Labor", [f"=SUM({scen_cols[i]}27:{scen_cols[i]}30)" for i in range(5)], MONEY, bold=True)

# OVERHEAD / MANAGEMENT LAYER — NEW
section(ws4, 33, "OVERHEAD / MANAGEMENT HEADCOUNT (triggered by client count)")

# For each overhead role, formula: IF(rentals >= trigger, 1 + FLOOR((rentals - trigger)/scaling), 0)
overhead_labels = ["Account Managers", "Onboarding Specialists", "Ops Managers", "QA Reviewers", "Analytics/Reporting", "Sales / BDR"]
for idx, lbl in enumerate(overhead_labels):
    src_row = OVERHEAD_START + idx  # row on Assumptions tab
    r = 34 + idx
    formulas = [
        f"=IF({scen_ref[i]}>=Assumptions!$C${src_row},1+FLOOR(MAX(0,{scen_ref[i]}-Assumptions!$C${src_row})/Assumptions!$D${src_row},1),0)"
        for i in range(5)
    ]
    row_across(ws4, r, lbl, formulas, INT, bold=True, fill=OVERHEAD_FILL)

# OVERHEAD LABOR COST
section(ws4, 41, "OVERHEAD / MANAGEMENT LABOR COST")
for idx, lbl in enumerate(overhead_labels):
    src_row = OVERHEAD_START + idx
    r = 42 + idx
    hc_row = 34 + idx  # headcount row above
    formulas = [f"={scen_cols[i]}{hc_row}*Assumptions!$B${src_row}" for i in range(5)]
    row_across(ws4, r, lbl, formulas, MONEY, fill=OVERHEAD_FILL)

row_across(ws4, 48, "Total Overhead Labor", [f"=SUM({scen_cols[i]}42:{scen_cols[i]}47)" for i in range(5)], MONEY, bold=True, fill=OVERHEAD_FILL)

# TOTAL LABOR + GROSS PROFIT (revised)
section(ws4, 50, "TOTAL LABOR (fulfillment + overhead)")
row_across(ws4, 51, "Total Monthly Labor (COGS)", [f"={scen_cols[i]}31+{scen_cols[i]}48" for i in range(5)], MONEY, bold=True)

section(ws4, 53, "GROSS PROFIT (steady state, fully loaded)")
row_across(ws4, 54, "Gross Profit", [f"={scen_cols[i]}12-{scen_cols[i]}51" for i in range(5)], MONEY, bold=True)
row_across(ws4, 55, "Gross Margin %", [f"={scen_cols[i]}54/{scen_cols[i]}12" for i in range(5)], PCT1, bold=True)
row_across(ws4, 56, "GP per client (all clients)",
    [f"=IF(({scen_cols[i]}5+{scen_cols[i]}6)=0,0,{scen_cols[i]}54/({scen_cols[i]}5+{scen_cols[i]}6))" for i in range(5)], MONEY)

# TEAM SIZE
section(ws4, 58, "TOTAL TEAM SIZE")
row_across(ws4, 59, "Fulfillment headcount",
    [f"=SUM({scen_cols[i]}21:{scen_cols[i]}24)" for i in range(5)], INT)
row_across(ws4, 60, "Overhead headcount",
    [f"=SUM({scen_cols[i]}34:{scen_cols[i]}39)" for i in range(5)], INT)
row_across(ws4, 61, "TOTAL people on Aperture",
    [f"={scen_cols[i]}59+{scen_cols[i]}60" for i in range(5)], INT, bold=True)

# ONBOARDING
section(ws4, 63, "ONE-TIME ONBOARDING INVESTMENT (during ramp)")
row_across(ws4, 64, "Onboarding cost (total, one-time)",
    [f"={scen_ref[i]}*{ONBOARD_COST_PER_CLIENT}" for i in range(5)], MONEY, bold=True)

ws4.cell(row=66, column=1, value="Onboarding hits during the ramp period. Once a client is live, only steady-state labor applies. The Onboarding Specialist role (overhead layer) absorbs most of this at scale.").font = Font(italic=True, color="808080")
ws4.cell(row=66, column=1).alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
ws4.row_dimensions[66].height = 40
ws4.merge_cells(start_row=66, start_column=1, end_row=66, end_column=6)

# ==============================================================
# TAB 5: COST PER CLIENT — NEW
# ==============================================================
ws5 = wb.create_sheet("Cost Per Client")
ws5.column_dimensions['A'].width = 38
for col in ['B','C','D','E','F']:
    ws5.column_dimensions[col].width = 16

ws5.cell(row=1, column=1, value="COST PER CLIENT — fully loaded (fulfillment + overhead)").font = Font(bold=True, size=14, color="1F3864")
ws5.cell(row=2, column=1, value="Everything is per-client per-month. Watch how overhead $/client drops as you scale — that's operating leverage.").font = Font(italic=True, color="808080")

hdr(ws5.cell(row=4, column=1), "Metric")
for i, name in enumerate(scen_names):
    hdr(ws5.cell(row=4, column=2+i), name)

# Total clients
row_across(ws5, 5, "Total clients (rental + current)",
    [f"='Scale Model'!{scen_cols[i]}5+'Scale Model'!{scen_cols[i]}6" for i in range(5)], INT, bold=True)

section(ws5, 7, "REVENUE PER CLIENT")
row_across(ws5, 8, "Net revenue / client (blended avg)",
    [f"=IF({scen_cols[i]}5=0,0,'Scale Model'!{scen_cols[i]}12/{scen_cols[i]}5)" for i in range(5)], MONEY, bold=True)

section(ws5, 10, "COST PER CLIENT (blended across all clients)")
row_across(ws5, 11, "Fulfillment cost / client",
    [f"=IF({scen_cols[i]}5=0,0,'Scale Model'!{scen_cols[i]}31/{scen_cols[i]}5)" for i in range(5)], MONEY2)
row_across(ws5, 12, "Overhead cost / client",
    [f"=IF({scen_cols[i]}5=0,0,'Scale Model'!{scen_cols[i]}48/{scen_cols[i]}5)" for i in range(5)], MONEY2, fill=OVERHEAD_FILL)
row_across(ws5, 13, "TOTAL cost / client",
    [f"={scen_cols[i]}11+{scen_cols[i]}12" for i in range(5)], MONEY, bold=True)

section(ws5, 15, "PROFIT PER CLIENT")
row_across(ws5, 16, "Gross profit / client",
    [f"={scen_cols[i]}8-{scen_cols[i]}13" for i in range(5)], MONEY, bold=True)
row_across(ws5, 17, "Gross margin %",
    [f"=IF({scen_cols[i]}8=0,0,{scen_cols[i]}16/{scen_cols[i]}8)" for i in range(5)], PCT1, bold=True)

section(ws5, 19, "COST BREAKDOWN AS % OF REVENUE")
row_across(ws5, 20, "Fulfillment as % of revenue",
    [f"=IF({scen_cols[i]}8=0,0,{scen_cols[i]}11/{scen_cols[i]}8)" for i in range(5)], PCT1)
row_across(ws5, 21, "Overhead as % of revenue",
    [f"=IF({scen_cols[i]}8=0,0,{scen_cols[i]}12/{scen_cols[i]}8)" for i in range(5)], PCT1, fill=OVERHEAD_FILL)

section(ws5, 23, "COST PER RENTAL CLIENT (isolated view)")
ws5.cell(row=24, column=1, value="Rental-only view: strip out current 3 clients, allocate overhead to rentals only").font = Font(italic=True, color="808080")

# Rental-only fulfillment cost = total fulfillment - current-3-clients cost
# Current 3 clients fulfillment cost = buyer + designer only ($1,220 base)
CURRENT_3_COST = "((Assumptions!B24+Assumptions!B25+Assumptions!B26)*Assumptions!F6+(Assumptions!C24+Assumptions!C25+Assumptions!C26)*Assumptions!F7)"
row_across(ws5, 25, "Rental fulfillment $ / rental client",
    [f"=IF('Scale Model'!{scen_cols[i]}5=0,0,('Scale Model'!{scen_cols[i]}31-{CURRENT_3_COST})/'Scale Model'!{scen_cols[i]}5)" for i in range(5)], MONEY2)
row_across(ws5, 26, "Rental overhead $ / rental client (overhead allocated to rentals only)",
    [f"=IF('Scale Model'!{scen_cols[i]}5=0,0,'Scale Model'!{scen_cols[i]}48/'Scale Model'!{scen_cols[i]}5)" for i in range(5)], MONEY2, fill=OVERHEAD_FILL)
row_across(ws5, 27, "TOTAL $ / rental client",
    [f"={scen_cols[i]}25+{scen_cols[i]}26" for i in range(5)], MONEY, bold=True)
row_across(ws5, 28, "Net revenue / rental client",
    [f"=Assumptions!B{RENTAL_PRICE}*(1-Assumptions!C{RENTAL_PRICE})"]*5, MONEY)
row_across(ws5, 29, "Gross profit / rental client",
    [f"={scen_cols[i]}28-{scen_cols[i]}27" for i in range(5)], MONEY, bold=True)
row_across(ws5, 30, "Rental client GM %",
    [f"=IF({scen_cols[i]}28=0,0,{scen_cols[i]}29/{scen_cols[i]}28)" for i in range(5)], PCT1, bold=True)

section(ws5, 32, "KEY TAKEAWAYS")
takeaways = [
    "• Fulfillment cost/client is roughly flat (~$150-200 blended for rentals) — that's your variable production cost.",
    "• Overhead cost/client SPIKES at each new-hire trigger (Account Mgr @ 10, Onboarding @ 15, etc.), then drops as clients fill the role's capacity.",
    "• Total cost/rental client at scale (60-90 clients) settles around $400-500 blended, on $1,250 net revenue → ~60-65% GM.",
    "• The partner split ($1,250/client) is the single biggest cost line — bigger than all labor combined. Worth revisiting terms once at scale.",
    "• Watch the 'Overhead as % of revenue' row — that's your operating leverage. Should shrink as you scale.",
]
for i, n in enumerate(takeaways):
    ws5.cell(row=33+i, column=1, value=n).alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    ws5.row_dimensions[33+i].height = 30

# ==============================================================
# TAB 6: TEAM ROADMAP (updated)
# ==============================================================
ws6 = wb.create_sheet("Team Roadmap")
ws6.column_dimensions['A'].width = 22
ws6.column_dimensions['B'].width = 90

ws6.cell(row=1, column=1, value="TEAM ROADMAP — layered hiring plan").font = Font(bold=True, size=14, color="1F3864")

sections_data = [
    ("Bottleneck priority order — Matt's next 3 hires", [
        "1. Account Manager @ 10 rentals ($1,500/mo) — protects retention, frees Matt from weekly comms.",
        "2. Onboarding Specialist @ 15 rentals ($1,200/mo) — kills biggest personal timesuck.",
        "3. Ops Manager @ 20 rentals ($2,000/mo) — needed once buyer team is 3+.",
    ]),
    ("Rate assumptions — all editable on Assumptions tab", [
        "AM $1,500 · Onboarding $1,200 · Ops Mgr $2,000 · QA $1,000 · Analytics $2,000 · Sales $2,500",
        "Rates assume senior offshore. For US-based senior roles, multiply by 3-5x.",
    ]),
    ("Above-the-line (NOT in COGS)", [
        "Matt's time (sales, strategy, oversight). Unsustainable at 60+ clients — plan for GM hire.",
        "Business partner's time (comp'd via rental partner split).",
        "Software/tools, G&A, sales commissions — add to Assumptions tab if you want a full P&L view.",
    ]),
    ("At-scale headlines (see Scale Model + Cost Per Client tabs for detail)", [
        "10 rentals: $5,070 labor · $11,630 GP (70%) · $390/client blended cost",
        "30 rentals: $18,840 labor · $22,860 GP (55%) · $571/client blended cost",
        "60 rentals: $30,830 labor · $48,370 GP (61%) · $489/client blended cost",
        "90 rentals: $44,820 labor · $71,880 GP (62%) · $482/client blended cost",
    ]),
]

r = 3
for title, bullets in sections_data:
    c = ws6.cell(row=r, column=1, value=title)
    c.font = Font(bold=True, size=12, color="1F3864")
    r += 1
    for b in bullets:
        c = ws6.cell(row=r, column=2, value=b)
        c.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        ws6.row_dimensions[r].height = 30
        r += 1
    r += 1

# Save
out_path = "/Users/matthewbrown/tmp/aperture-cogs-v2/aperture-cogs-model-v2.xlsx"
wb.save(out_path)
print(f"Saved to {out_path}")
