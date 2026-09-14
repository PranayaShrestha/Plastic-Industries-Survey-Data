"""
Phase 0 ETL — Fuel-Switching Survey (Thermic Fluid Heating & Hot Air Generators,
Printing/Packaging Industry, AEPC study)

Reads the raw KoboToolbox export (3 sheets: main survey, printing_units,
laminating_units), merges them on submission ID, and writes a new workbook with:

  - Assumptions      : editable calorific values, efficiencies, load-band
                        midpoints, heat-pump COP. Change these and everything
                        downstream recalculates.
  - Cleaned_Units     : one row per heater, all raw fields carried over plus
                        formula-driven unit conversion, energy balance,
                        consistency cross-check against nameplate capacity,
                        and an electric-equivalent sizing estimate.
  - Plant_Summary     : rolled up per industry, checked against reported
                        transformer capacity / spare capacity.
  - Data_Quality_Flags: plain-language list of specific issues found in the
                        raw data (missing fields, an implausible reading,
                        a suspected shared central heater), each with a
                        follow-up action.

Re-run this any time a new response lands in the raw KoboToolbox export —
it is not a one-off.

Usage:
    python3 build_cleaned_workbook.py <raw_survey.xlsx> <output.xlsx>
"""
import sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

RAW_PATH = sys.argv[1] if len(sys.argv) > 1 else "survey.xlsx"
OUT_PATH = sys.argv[2] if len(sys.argv) > 2 else "Fuel_Switching_Cleaned_Data.xlsx"

# ---------------------------------------------------------------- load raw --
wb_raw = openpyxl.load_workbook(RAW_PATH, data_only=True)

# The main-survey sheet is named after the KoboToolbox form title, which Excel
# truncates to 31 chars — the exact rendering (literal "..." vs an ellipsis
# character, exact cutoff point) can differ between exports. The two
# repeat-group sheets have fixed, stable names, so identify the main sheet as
# "whichever sheet isn't one of those" instead of hardcoding its title.
KNOWN_UNIT_SHEETS = {"printing_units", "laminating_units"}
candidates = [s for s in wb_raw.sheetnames if s not in KNOWN_UNIT_SHEETS]
if len(candidates) != 1:
    raise SystemExit(
        f"Expected exactly one main-survey sheet besides {sorted(KNOWN_UNIT_SHEETS)}, "
        f"found: {candidates}. Check wb_raw.sheetnames and adjust KNOWN_UNIT_SHEETS "
        f"if your export uses different group names for the repeat sections."
    )
main_ws = wb_raw[candidates[0]]
main_headers = [c.value for c in main_ws[1]]
main_rows = list(main_ws.iter_rows(min_row=2, values_only=True))


def col(headers, name):
    try:
        return headers.index(name)
    except ValueError:
        raise SystemExit(
            f"\nCould not find the column '{name}' in this sheet.\n"
            f"Available columns in this sheet:\n  - " + "\n  - ".join(str(h) for h in headers) +
            f"\n\nIf you opened the raw KoboToolbox export in Excel and cleaned it up before "
            f"running this script (e.g. deleted the underscore-prefixed system columns like "
            f"_id, _uuid, _submission_time, __version__), that's almost certainly the cause: "
            f"those columns look like clutter but they're the join keys this script uses to "
            f"link the main sheet to printing_units/laminating_units. Re-run this against the "
            f"original, unedited export (re-download it fresh from KoboToolbox if the edited "
            f"copy is all you have left) and it should work."
        )


mi = {
    "id": col(main_headers, "_id"),
    "name": col(main_headers, "Industry name"),
    "prod": col(main_headers, "Daily production?"),
    "op_hours": col(main_headers, "Daily operating hours?"),
    "op_days": col(main_headers, "Annual operating days"),
    "xfmr": col(main_headers, "Existing transformer capacity"),
    "max_demand": col(main_headers, "Maximum demand for the last year"),
    "spare": col(main_headers, "Is spare electrical capacity available?"),
    "grid": col(main_headers, "Grid supply reliability"),
}

main_by_id = {r[mi["id"]]: r for r in main_rows}


def load_unit_sheet(sheet_name, category, headers_map):
    ws = wb_raw[sheet_name]
    headers = [c.value for c in ws[1]]
    idx = {k: col(headers, v) for k, v in headers_map.items()}
    out = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        rec = {k: r[i] for k, i in idx.items()}
        rec["unit_category"] = category
        out.append(rec)
    return out


unit_field_map = {
    "submission_id": "_submission__id",
    "arrangement": None,  # filled per-sheet below (name differs by template var)
    "fuel": "Fuel / energy source used for this heating arrangement",
    "cap_kcalhr": "Heater / boiler capacity (kcal/hr)",
    "cap_kW": "Heater capacity (kW)",
    "load_band": "Typical operating load",
    "hours": "Running hours",
    "fuel_L_day": "Daily fuel consumption (litres/day)",
    "fuel_kg_day": "Daily fuel consumption (kg/day)",
    "elec_kWh_day": "Daily electricity consumption (kWh/day)",
    "annual_qty": None,
    "annual_cost": None,
    "req_temp_C": None,
}

p_map = dict(unit_field_map)
p_map["arrangement"] = "Heating arrangement used for printing unit ${p_unit_index}"
p_map["annual_qty"] = "p_annual_fuel_quantity"
p_map["annual_cost"] = "p_annual_fuel_cost"
p_map["req_temp_C"] = "Required temperature for ink drying"

l_map = dict(unit_field_map)
l_map["arrangement"] = "Heating arrangement used for laminating unit ${l_unit_index}"
l_map["annual_qty"] = "l_annual_fuel_quantity"
l_map["annual_cost"] = "l_annual_fuel_cost"
l_map["req_temp_C"] = "Required process temperature"

printing = load_unit_sheet("printing_units", "Printing", p_map)
laminating = load_unit_sheet("laminating_units", "Laminating", l_map)
units = printing + laminating

# ------------------------------------------------------- build new workbook --
wb = openpyxl.Workbook()
wb.remove(wb.active)

FONT = "Arial"
HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(name=FONT, bold=True, color="FFFFFF", size=10)
INPUT_FONT = Font(name=FONT, color="0000FF", size=10)      # blue = editable input
FORMULA_FONT = Font(name=FONT, color="000000", size=10)    # black = formula
FLAG_FILL = PatternFill("solid", fgColor="FFC7CE")
NOTE_FONT = Font(name=FONT, italic=True, size=9, color="7F7F7F")
TITLE_FONT = Font(name=FONT, bold=True, size=13)
thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)


def style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(wrap_text=True, vertical="center")
        cell.border = BORDER


def autosize(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


# ===================================================== 1. Assumptions sheet =
asm = wb.create_sheet("Assumptions")
asm["A1"] = "Fuel-Switching Study — Assumptions"
asm["A1"].font = TITLE_FONT
asm["A2"] = ("Blue cells are editable inputs. Every downstream sheet recalculates "
             "from these — change a value here to test a different assumption.")
asm["A2"].font = NOTE_FONT
asm.merge_cells("A2:D2")

headers = ["Parameter", "Value", "Unit", "Notes / source (verify before final use)"]
asm.append([])
asm.append(headers)
style_header(asm, 4, 4)

rows = [
    ("Diesel calorific value", 10.0, "kWh/L", "Typical diesel LHV ~36 MJ/L (~42.8 MJ/kg @ ~0.84 kg/L). ASSUMPTION — confirm with fuel supplier or lab test."),
    ("Pellet calorific value", 4.7, "kWh/kg", "Typical biomass pellet ~17 MJ/kg, moisture-dependent. ASSUMPTION — confirm with supplier spec sheet per plant."),
    ("Rice husk calorific value", 3.75, "kWh/kg", "Typical raw rice husk ~13.5 MJ/kg (lower than pellet due to ash/moisture). ASSUMPTION — verify, esp. for Sona Packaging."),
    ("kcal/hr → kW conversion", 0.001163, "kW per kcal/hr", "Standard conversion (1 kWh = 860 kcal)."),
    ("Pellet hot-air generator efficiency", 0.80, "fraction", "Typical industrial biomass hot-air generator. ASSUMPTION — verify by flue-gas test if possible."),
    ("Rice husk thermic-oil heater efficiency", 0.70, "fraction", "Lower than pellet due to husk ash/moisture content. ASSUMPTION."),
    ("Diesel hot-air generator efficiency", 0.82, "fraction", "Typical small diesel-fired hot-air generator. ASSUMPTION."),
    ("Electric resistance heating efficiency", 0.98, "fraction", "Near-unity; small enclosure/duct losses. ASSUMPTION."),
    ("Heat pump COP (process temp ≤ 90°C)", 2.8, "-", "Indicative industrial heat-pump COP at this lift. ASSUMPTION — confirm against a shortlisted vendor's spec at the actual required lift."),
    ("Load band midpoint: 50–75%", 0.625, "fraction", "Midpoint of the survey's qualitative load-band answer."),
    ("Load band midpoint: 75–100%", 0.875, "fraction", "Midpoint of the survey's qualitative load-band answer."),
    ("Load band midpoint: Variable / Don't know", 0.50, "fraction", "Conservative default only — these units should be prioritised for a follow-up metered reading."),
]
start = 5
for i, (name, val, unit, note) in enumerate(rows):
    r = start + i
    asm.cell(row=r, column=1, value=name).font = FORMULA_FONT
    c = asm.cell(row=r, column=2, value=val)
    c.font = INPUT_FONT
    asm.cell(row=r, column=3, value=unit).font = FORMULA_FONT
    note_cell = asm.cell(row=r, column=4, value=note)
    note_cell.font = NOTE_FONT
    note_cell.alignment = Alignment(wrap_text=True)
    for cc in range(1, 5):
        asm.cell(row=r, column=cc).border = BORDER

autosize(asm, [38, 12, 16, 90])
for i in range(start, start + len(rows)):
    asm.row_dimensions[i].height = 30

# cell refs for use in formulas below
A = {
    "diesel_cv": "Assumptions!$B$5",
    "pellet_cv": "Assumptions!$B$6",
    "husk_cv": "Assumptions!$B$7",
    "kcal2kw": "Assumptions!$B$8",
    "eff_pellet": "Assumptions!$B$9",
    "eff_husk": "Assumptions!$B$10",
    "eff_diesel": "Assumptions!$B$11",
    "eff_elec": "Assumptions!$B$12",
    "cop_hp": "Assumptions!$B$13",
    "mid_5075": "Assumptions!$B$14",
    "mid_75100": "Assumptions!$B$15",
    "mid_var": "Assumptions!$B$16",
}

# ===================================================== 2. Cleaned_Units ====
cu = wb.create_sheet("Cleaned_Units")
cu["A1"] = "Cleaned & Normalized Unit-Level Data"
cu["A1"].font = TITLE_FONT
cu["A2"] = ("One row per heater. Grey columns = raw survey data (unchanged). "
            "White columns = calculated from the Assumptions sheet. "
            "Flag columns call out rows worth a follow-up call before they go into any model.")
cu["A2"].font = NOTE_FONT
cu.merge_cells("A2:H2")

headers = [
    "Industry", "Category", "Arrangement", "Fuel",
    "Cap (kcal/hr)", "Cap (kW, raw)", "Cap (kW, normalized)",
    "Load band", "Load midpoint", "Hours/day", "Op. days/yr",
    "Fuel (L/day)", "Fuel (kg/day)", "Elec (kWh/day)",
    "Fuel energy input (kWh/day)", "Assumed efficiency",
    "Useful heat output (kWh/day)", "Rated theoretical output (kWh/day)",
    "Consistency ratio", "Consistency flag",
    "Annual thermal energy (kWh/yr)", "Annual cost (NPR, as reported)",
    "Avg thermal duty (kW)", "Electric-equivalent (kW, resistive)",
    "Required temp (°C)", "Heat-pump feasible?", "Electric-equivalent (kW, heat pump)",
    "Possible duplicate/shared-heater flag", "Analyst note",
]
row0 = 4
cu.append([])
cu.append([])
for c, h in enumerate(headers, start=1):
    cu.cell(row=row0, column=c, value=h)
style_header(cu, row0, len(headers))

n = len(units)
first_data_row = row0 + 1
last_data_row = row0 + n

manual_notes = {}
for i, u in enumerate(units):
    r = first_data_row + i
    main = main_by_id.get(u["submission_id"], [None] * len(main_headers))
    industry = main[mi["name"]] if main[0] is not None else None
    op_days = main[mi["op_days"]]

    if industry == "Godawari Printng" and u["unit_category"] == "Laminating":
        manual_notes[r] = ("Reported 1,096 kWh/day against a 28 kW nameplate — "
                            "physically impossible even at 24h/day. Re-verify meter reading before use.")
    if industry == "Uday Shree Print Pack Pvt Ltd":
        manual_notes[r] = "Submission largely incomplete (production, cost, temperature missing). Follow up before including in totals."
    if industry == "Sona Packaging":
        manual_notes[r] = ("Printing and laminating rows report identical capacity/fuel figures — "
                            "likely ONE central thermic-fluid heater serving both processes, not two. "
                            "See Plant_Summary note; do not sum naively.")

    vals = [
        industry, u["unit_category"], u["arrangement"], u["fuel"],
        u["cap_kcalhr"], u["cap_kW"],
        None,  # H normalized cap - formula
        u["load_band"],
        None,  # I load midpoint - formula
        u["hours"], op_days,
        u["fuel_L_day"], u["fuel_kg_day"], u["elec_kWh_day"],
        None, None, None, None, None, None,  # formulas
        None,  # annual thermal energy - formula
        u["annual_cost"],
        None, None,  # formulas
        u["req_temp_C"],
        None, None,  # formulas
        None,  # duplicate flag - formula
        manual_notes.get(r, ""),
    ]
    for c, v in enumerate(vals, start=1):
        cell = cu.cell(row=r, column=c, value=v)
        cell.border = BORDER
        if c in (1, 2, 3, 4, 5, 6, 8, 10, 11, 12, 13, 14, 22, 25):
            cell.fill = PatternFill("solid", fgColor="F2F2F2")  # raw data = grey

# NOTE: column indices — recompute mapping explicitly for formulas by letter
# 1 Industry(A) 2 Category(B) 3 Arrangement(C) 4 Fuel(D) 5 CapKcal(E) 6 CapKWraw(F)
# 7 CapKWnorm(G) 8 LoadBand(H) 9 LoadMid(I) 10 Hours(J) 11 OpDays(K)
# 12 FuelL(L) 13 FuelKg(M) 14 ElecKWh(N) 15 FuelEnergyIn(O) 16 Efficiency(P)
# 17 UsefulOut(Q) 18 RatedOut(R) 19 ConsistRatio(S) 20 ConsistFlag(T)
# 21 AnnualThermal(U) 22 AnnualCost(V) 23 AvgDutyKW(W) 24 ElecEquivResistive(X)
# 25 ReqTemp(Y) 26 HPfeasible(Z) 27 ElecEquivHP(AA) 28 DupFlag(AB) 29 Note(AC)

for i in range(n):
    r = first_data_row + i
    cu[f"G{r}"] = f'=IF(F{r}<>"",F{r},IF(E{r}<>"",E{r}*{A["kcal2kw"]},""))'
    cu[f"I{r}"] = (f'=IF(H{r}="50–75%",{A["mid_5075"]},'
                   f'IF(H{r}="75–100%",{A["mid_75100"]},'
                   f'IF(H{r}="Variable / Don\'t know",{A["mid_var"]},"")))')
    cu[f"O{r}"] = (f'=IF(D{r}="Diesel",L{r}*{A["diesel_cv"]},'
                   f'IF(D{r}="Pellet",M{r}*{A["pellet_cv"]},'
                   f'IF(D{r}="Rice Husk",M{r}*{A["husk_cv"]},'
                   f'IF(D{r}="Electric",N{r},""))))')
    cu[f"P{r}"] = (f'=IF(D{r}="Electric",{A["eff_elec"]},'
                   f'IF(D{r}="Pellet",{A["eff_pellet"]},'
                   f'IF(D{r}="Rice Husk",{A["eff_husk"]},'
                   f'IF(D{r}="Diesel",{A["eff_diesel"]},""))))')
    cu[f"Q{r}"] = f'=IF(AND(ISNUMBER(O{r}),ISNUMBER(P{r})),O{r}*P{r},"")'
    cu[f"R{r}"] = f'=IF(AND(ISNUMBER(G{r}),ISNUMBER(J{r}),ISNUMBER(I{r})),G{r}*J{r}*I{r},"")'
    cu[f"S{r}"] = f'=IFERROR(Q{r}/R{r},"")'
    cu[f"T{r}"] = (f'=IF(S{r}="","insufficient data",'
                   f'IF(OR(S{r}>1.3,S{r}<0.05),"CHECK – implausible vs rated capacity","plausible"))')
    cu[f"U{r}"] = f'=IF(AND(ISNUMBER(Q{r}),ISNUMBER(K{r})),Q{r}*K{r},"")'
    cu[f"W{r}"] = f'=IF(AND(ISNUMBER(Q{r}),ISNUMBER(J{r}),J{r}<>0),Q{r}/J{r},"")'
    cu[f"X{r}"] = f'=IF(ISNUMBER(W{r}),W{r}/{A["eff_elec"]},"")'
    cu[f"Z{r}"] = f'=IF(ISNUMBER(Y{r}),IF(Y{r}<=90,"Yes","No \u2013 high temp"),"Unknown")'
    cu[f"AA{r}"] = f'=IF(AND(ISNUMBER(W{r}),Z{r}="Yes"),W{r}/{A["cop_hp"]},"N/A")'
    cu[f"AB{r}"] = (f'=IF(COUNTIFS($A${first_data_row}:$A${last_data_row},A{r},'
                     f'$D${first_data_row}:$D${last_data_row},D{r},'
                     f'$M${first_data_row}:$M${last_data_row},M{r})>1,'
                     f'"Check \u2013 possible shared heater","")')
    for col_letter in ["T"]:
        cu[f"{col_letter}{r}"].font = FORMULA_FONT

    for c in range(7, 29):
        cu.cell(row=r, column=c).border = BORDER

autosize(cu, [22, 11, 15, 9, 11, 11, 12, 12, 11, 9, 10, 10, 10, 11, 13, 11,
              13, 14, 10, 15, 14, 13, 12, 14, 10, 11, 14, 18, 34])
cu.freeze_panes = f"E{first_data_row}"

# ===================================================== 3. Plant_Summary ====
ps = wb.create_sheet("Plant_Summary")
ps["A1"] = "Plant-Level Summary"
ps["A1"].font = TITLE_FONT
ps["A2"] = "Rolled up from Cleaned_Units (naive sum — see notes column for plants needing manual de-duplication)."
ps["A2"].font = NOTE_FONT
ps.merge_cells("A2:H2")

headers2 = [
    "Industry", "Total thermal capacity (kW, naive sum)",
    "Total annual thermal energy (kWh/yr, naive sum)",
    "Total electric-equivalent added load (kW, resistive)",
    "Existing transformer capacity", "Max demand last year",
    "Approx. headroom", "Added load vs. headroom",
    "Spare capacity reported?", "Grid reliability", "Note",
]
r0 = 4
for c, h in enumerate(headers2, start=1):
    ps.cell(row=r0, column=c, value=h)
style_header(ps, r0, len(headers2))

plants = []
seen = set()
for u in units:
    sid = u["submission_id"]
    if sid not in seen:
        seen.add(sid)
        main = main_by_id.get(sid, [None] * len(main_headers))
        plants.append((sid, main[mi["name"]], main[mi["xfmr"]], main[mi["max_demand"]],
                        main[mi["spare"]], main[mi["grid"]]))

plant_notes = {
    "Sona Packaging": "Printing/laminating heater figures look identical \u2013 likely ONE shared central heater double-counted here. Treat this total as an upper bound until confirmed.",
    "Uday Shree Print Pack Pvt Ltd": "Submission mostly incomplete \u2013 totals below are partial, not representative.",
    "Godawari Printng": "Laminating unit has an implausible consumption reading \u2013 re-verify before trusting this total.",
}

pr0 = r0 + 1
for i, (sid, name, xfmr, maxd, spare, grid) in enumerate(plants):
    r = pr0 + i
    ps.cell(row=r, column=1, value=name).fill = PatternFill("solid", fgColor="F2F2F2")
    ps[f"B{r}"] = f'=SUMIFS(Cleaned_Units!$G${first_data_row}:$G${last_data_row},Cleaned_Units!$A${first_data_row}:$A${last_data_row},A{r})'
    ps[f"C{r}"] = f'=SUMIFS(Cleaned_Units!$U${first_data_row}:$U${last_data_row},Cleaned_Units!$A${first_data_row}:$A${last_data_row},A{r})'
    ps[f"D{r}"] = f'=SUMIFS(Cleaned_Units!$X${first_data_row}:$X${last_data_row},Cleaned_Units!$A${first_data_row}:$A${last_data_row},A{r})'
    ps.cell(row=r, column=5, value=xfmr).fill = PatternFill("solid", fgColor="F2F2F2")
    ps.cell(row=r, column=6, value=maxd).fill = PatternFill("solid", fgColor="F2F2F2")
    ps[f"G{r}"] = f'=IFERROR(E{r}-F{r},"")'
    ps[f"H{r}"] = f'=IF(G{r}="","insufficient data",IF(D{r}<=G{r},"Likely fits existing headroom","Likely needs capacity upgrade"))'
    ps.cell(row=r, column=9, value=spare).fill = PatternFill("solid", fgColor="F2F2F2")
    ps.cell(row=r, column=10, value=grid).fill = PatternFill("solid", fgColor="F2F2F2")
    note_cell = ps.cell(row=r, column=11, value=plant_notes.get(name, ""))
    note_cell.alignment = Alignment(wrap_text=True)
    for c in range(1, 12):
        ps.cell(row=r, column=c).border = BORDER

autosize(ps, [26, 16, 18, 18, 16, 15, 12, 18, 14, 14, 46])
for i in range(pr0, pr0 + len(plants)):
    ps.row_dimensions[i].height = 32

# ============================================== 4. Data_Quality_Flags ======
dq = wb.create_sheet("Data_Quality_Flags")
dq["A1"] = "Data Quality Flags — Follow Up Before Final Modelling"
dq["A1"].font = TITLE_FONT
headers3 = ["#", "Industry / scope", "Issue", "Why it matters", "Recommended action"]
for c, h in enumerate(headers3, start=1):
    dq.cell(row=3, column=c, value=h)
style_header(dq, 3, 5)

flags = [
    ("Sona Packaging",
     "Printing-unit and laminating-unit rows report identical heater capacity, fuel type and daily consumption (1,000,000 kcal/hr, rice husk, 3,000 kg/day).",
     "Strong sign it is one central thermic-fluid loop feeding both processes, not two separate heaters. Naively summing the two rows roughly doubles the true energy and capacity figures for this plant.",
     "Confirm with the plant whether printing and laminating share one heater. If yes, use only one of the two rows in any total."),
    ("Godawari Printng \u2013 laminating unit",
     "28 kW nameplate reported alongside 1,096 kWh/day consumption \u2013 exceeds the physical maximum (672 kWh/day at 24h) by ~60%.",
     "Either the capacity, the consumption reading, or a decimal/unit was mis-recorded. As-is, it will distort any per-plant energy or cost total.",
     "Re-contact the plant to re-verify both the nameplate rating and the electricity bill/meter reading for this unit."),
    ("Uday Shree Print Pack Pvt Ltd",
     "Production, operating days, annual fuel quantity, annual cost and required temperature are all missing.",
     "Not enough data to place this plant in the energy baseline or the techno-economic model yet.",
     "Follow up to complete the submission before including this plant in Phase 1 onward."),
    ("All plants",
     "\"Daily production?\" has no stated unit in the survey (values recorded: 800\u201315,000).",
     "Needed to compute specific energy consumption (kWh per kg of product), the most defensible cross-plant benchmark and the basis for extrapolating to the wider sector.",
     "Add a unit to this survey question (confirm kg/day vs. another unit) before the next round of data collection."),
    ("2 units (Uday Shree printing; Galaxy Packaging printing unit 3)",
     "Load band recorded as \"Variable / Don't know\".",
     "The energy and cost figures for these units currently rely on a conservative default 50% load assumption rather than a real reading.",
     "Prioritise these units for a short metered spot-check rather than relying on the survey answer."),
    ("Whole dataset",
     "Only 7 industries submitted so far.",
     "Fine for the inception/pilot stage, but any sector-wide extrapolation now would carry wide uncertainty bands.",
     "Treat Phase 1\u20134 outputs as provisional until the fuller inventory (per the ToR's industry-mapping step) is surveyed."),
]
for i, (scope, issue, why, action) in enumerate(flags):
    r = 4 + i
    dq.cell(row=r, column=1, value=i + 1)
    dq.cell(row=r, column=2, value=scope)
    dq.cell(row=r, column=3, value=issue)
    dq.cell(row=r, column=4, value=why)
    dq.cell(row=r, column=5, value=action)
    for c in range(1, 6):
        cell = dq.cell(row=r, column=c)
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        cell.border = BORDER
    dq.row_dimensions[r].height = 60

autosize(dq, [4, 26, 42, 42, 40])

wb.save(OUT_PATH)
print(f"Wrote {OUT_PATH}: {n} unit rows, {len(plants)} plants.")
