"""
Fuel-switching survey data foundation
=====================================
Portable, local Python script for the AEPC printing/laminating survey.

Inputs
------
One Excel workbook containing:
  1) company-level survey sheet
  2) printing_units
  3) laminating_units

Outputs
-------
output_dir/
  clean_per_unit.csv
  anomalies.csv
  clean_per_unit.xlsx
  company_summary.csv
  plots/
    01_capacity_vs_daily_energy.png
    02_implied_average_power_vs_capacity.png
    03_transformer_vs_max_demand.png
    04_daily_energy_by_sector_and_fuel.png
    05_operating_hours.png
    06_capacity_utilization_screen.png

Dependencies
------------
pip install pandas numpy matplotlib seaborn openpyxl

Important engineering note
--------------------------
This script creates a COMMON INPUT-ENERGY BASIS. It does not claim that all
reported input energy becomes useful process heat. Useful heat should be
calculated later from measured/verified system efficiencies and process duty.

Conversion assumptions are explicit and editable below.
"""

from __future__ import annotations

from pathlib import Path
import re
import warnings

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# =============================================================================
# USER SETTINGS
# =============================================================================

# Apply Seaborn theme for modernized, professional visualizations
sns.set_theme(style="whitegrid", context="talk", font_scale=0.8, palette="deep")

# By default, work relative to the folder containing this script.
# The workbook name may change when it is downloaded/renamed, so the script
# automatically searches that folder when the configured file is not found.
SCRIPT_DIR = Path(__file__).resolve().parent
INPUT_FILE = SCRIPT_DIR / (
    "Thermic_fluid_Heating_and_hot_air_generator_Survey_of_Printing_and_Packaging_Industry_for_Fuel_Switching_-_all_versions_-_labels_-_2026-09-13-13-09-37.xlsx"
)
OUTPUT_DIR = SCRIPT_DIR / "fuel_switching_outputs"
PLOTS_DIR = OUTPUT_DIR / "plots"

# Screening heating values / energy factors.
# Replace these with laboratory/supplier values once verified during field audit.
KCAL_PER_HR_TO_KW = 0.001163  # 1 kcal/hr = 0.001163 kW
DIESEL_KWH_PER_L = 10.66  # approx. LHV/input energy basis
PELLET_MJ_PER_KG = 16.0
RICE_HUSK_MJ_PER_KG = 14.5
MJ_PER_KWH = 3.6

# Optional benchmark efficiencies ONLY for contextual screening, not final sizing.
BENCHMARK_EFFICIENCY = {
    "diesel": 0.80,
    "pellet": 0.75,
    "rice husk": 0.70,
    "electric": 0.98,
}

# Anomaly thresholds. These are intentionally conservative.
# A hard anomaly means the reported daily energy exceeds rated-capacity × hours.
HARD_OVERENERGY_FACTOR = 1.20
LOW_UTILIZATION_FACTOR = 0.15
HIGH_IMPLIED_LOAD_FACTOR = 1.20


# =============================================================================
# HELPERS
# =============================================================================

def norm_text(x) -> str:
    if pd.isna(x):
        return ""
    return str(x).strip()


def canonical_fuel(x) -> str:
    s = norm_text(x).lower()
    if not s:
        return "unknown"
    if "electric" in s:
        return "electric"
    if "diesel" in s:
        return "diesel"
    if "pellet" in s:
        return "pellet"
    if "rice" in s and "husk" in s:
        return "rice husk"
    return s


def canonical_arrangement(x) -> str:
    s = norm_text(x).lower()
    if "thermic" in s:
        return "thermic oil"
    if "hot air" in s:
        return "hot air generator"
    return norm_text(x)


def to_number(value) -> float:
    """Convert numbers and common numeric strings like '10', '10–12', '50,000'."""
    if pd.isna(value):
        return np.nan
    if isinstance(value, (int, float, np.integer, np.floating)):
        return float(value)
    s = str(value).strip().replace(",", "")
    if not s:
        return np.nan
    # direct conversion first
    try:
        return float(s)
    except ValueError:
        # use first numeric token when a field contains extra text
        m = re.search(r"[-+]?\d+(?:\.\d+)?", s)
        return float(m.group()) if m else np.nan


def parse_load_range(value):
    """Return low/high/representative decimal load fractions from survey labels."""
    s = norm_text(value).lower().replace("%", "")
    if not s:
        return np.nan, np.nan, np.nan
    nums = re.findall(r"\d+(?:\.\d+)?", s)
    if len(nums) >= 2:
        lo, hi = float(nums[0]) / 100, float(nums[1]) / 100
        return lo, hi, (lo + hi) / 2
    if len(nums) == 1:
        v = float(nums[0]) / 100
        return v, v, v
    if "variable" in s or "don't know" in s or "dont know" in s:
        return np.nan, np.nan, np.nan
    return np.nan, np.nan, np.nan


def first_existing(df: pd.DataFrame, names: list[str]) -> str | None:
    for name in names:
        if name in df.columns:
            return name
    return None


def safe_div(a, b):
    a = pd.to_numeric(a, errors="coerce")
    b = pd.to_numeric(b, errors="coerce")
    out = pd.Series(np.nan, index=a.index, dtype=float)
    mask = b.notna() & (b != 0)
    out.loc[mask] = a.loc[mask] / b.loc[mask]
    return out


# =============================================================================
# LOAD WORKBOOK
# =============================================================================

def resolve_input_workbook(path: Path) -> Path:
    """Resolve the input workbook robustly for local Windows/macOS/Linux use."""
    path = Path(path)

    # 1. Exact configured path.
    if path.exists() and path.is_file():
        return path

    # 2. Search beside the script for likely survey workbooks.
    folder = path.parent if path.parent != Path(".") else SCRIPT_DIR
    candidates = sorted(
        [p for p in folder.glob("*.xlsx") if not p.name.startswith("~$")],
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )

    if not candidates:
        raise FileNotFoundError(
            "No Excel workbook (*.xlsx) was found in the script folder:\n"
            f"  {folder}\n\n"
            "Put the survey workbook in the same folder as this script. "
            "The script will detect it automatically."
        )

    # Prefer filenames that look like the Thermic fluid / printing survey.
    keywords = ("thermic", "fluid", "printing", "packaging", "fuel", "switch")
    likely = [p for p in candidates if any(k in p.name.lower() for k in keywords)]
    chosen = likely[0] if likely else candidates[0]

    print("Configured workbook was not found.")
    print(f"Automatically selected: {chosen.name}")
    return chosen


def load_workbook(path: Path):
    path = resolve_input_workbook(path)
    print(f"Using input workbook: {path.resolve()}")

    xls = pd.ExcelFile(path, engine="openpyxl")
    print("Workbook sheets:")
    for i, s in enumerate(xls.sheet_names, 1):
        print(f"  {i}. {s}")

    # Match by actual sheet names instead of assuming exact company-sheet title.
    unit_sheet_map = {}
    for s in xls.sheet_names:
        sl = s.lower()
        if "printing" in sl and "unit" in sl:
            unit_sheet_map["printing"] = s
        if "laminating" in sl and "unit" in sl:
            unit_sheet_map["laminating"] = s

    if "printing" not in unit_sheet_map or "laminating" not in unit_sheet_map:
        raise ValueError("Could not identify both printing_units and laminating_units sheets.")

    unit_sheets = {k: pd.read_excel(path, sheet_name=v, engine="openpyxl")
                   for k, v in unit_sheet_map.items()}

    company_candidates = [s for s in xls.sheet_names if s not in unit_sheet_map.values()]
    if not company_candidates:
        raise ValueError("Could not identify the company-level survey sheet.")
    company_sheet = company_candidates[0]
    company = pd.read_excel(path, sheet_name=company_sheet, engine="openpyxl")

    return company, unit_sheets["printing"], unit_sheets["laminating"]


# =============================================================================
# COMPANY-LEVEL FIELDS
# =============================================================================

def prepare_company(company: pd.DataFrame) -> pd.DataFrame:
    if "_id" not in company.columns:
        raise ValueError("Company sheet is missing '_id', which is the parent/submission identifier.")

    out = pd.DataFrame()
    out["submission_id"] = company["_id"]

    field_map = {
        "industry_name": ["Industry name"],
        "location": ["Address / location"],
        "transformer_capacity_kva": ["Existing transformer capacity"],
        "maximum_demand_kw": ["Maximum demand for the last year"],
        "spare_electrical_capacity": ["Is spare electrical capacity available?"],
        "grid_reliability": ["Grid supply reliability"],
        "space_available": ["Is sufficient space available for an alternative heating system?"],
        "daily_operating_hours_company": ["Daily operating hours?"],
        "annual_operating_days": ["Annual operating days"],
        "daily_production": ["Daily production?"],
        "willing_to_switch": ["Are you willing to switch the existing fuel/heating system?"],
    }

    for dest, candidates in field_map.items():
        src = first_existing(company, candidates)
        out[dest] = company[src] if src else np.nan

    for c in [
        "transformer_capacity_kva",
        "maximum_demand_kw",
        "daily_operating_hours_company",
        "annual_operating_days",
    ]:
        out[c] = out[c].apply(to_number)

    return out


# =============================================================================
# UNIT TABLE PREPARATION
# =============================================================================

def prepare_units(df: pd.DataFrame, sector: str, company: pd.DataFrame) -> pd.DataFrame:
    if "_submission__id" not in df.columns:
        raise ValueError(f"{sector} sheet is missing '_submission__id'.")

    arrangement_col = first_existing(
        df,
        [
            "Heating arrangement used for printing unit ${p_unit_index}",
            "Heating arrangement used for laminating unit ${l_unit_index}",
        ],
    )
    if not arrangement_col:
        raise ValueError(f"Could not identify heating-arrangement column in {sector} sheet.")

    # Preserve all original survey fields, but use normalized names for analysis.
    out = pd.DataFrame(index=df.index)
    out["submission_id"] = df["_submission__id"]
    out["sector"] = sector

    unit_index_col = "p_unit_index" if sector == "printing" else "l_unit_index"
    out["unit_index"] = df[unit_index_col] if unit_index_col in df else np.arange(1, len(df) + 1)
    out["unit_id"] = (
            out["submission_id"].astype(str)
            + "-"
            + sector
            + "-"
            + out["unit_index"].astype(str)
    )

    out["heating_arrangement"] = df[arrangement_col].apply(canonical_arrangement)
    out["fuel"] = df["Fuel / energy source used for this heating arrangement"].apply(canonical_fuel)

    cap_kcal_col = "Heater / boiler capacity (kcal/hr)"
    cap_kw_col = "Heater capacity (kW)"
    out["capacity_kcal_hr_reported"] = df[cap_kcal_col].apply(to_number) if cap_kcal_col in df else np.nan
    out["capacity_kw_reported"] = df[cap_kw_col].apply(to_number) if cap_kw_col in df else np.nan

    # Prefer explicit kW. Otherwise convert explicit kcal/hr field.
    out["capacity_kw_thermal"] = out["capacity_kw_reported"]
    mask = out["capacity_kw_thermal"].isna() & out["capacity_kcal_hr_reported"].notna()
    out.loc[mask, "capacity_kw_thermal"] = out.loc[mask, "capacity_kcal_hr_reported"] * KCAL_PER_HR_TO_KW
    out["capacity_source"] = np.where(
        out["capacity_kw_reported"].notna(),
        "reported_kW",
        np.where(mask, "converted_kcal_per_hr", "missing"),
    )

    load_col = "Typical operating load"
    out["operating_load_raw"] = df[load_col] if load_col in df else np.nan
    parsed = out["operating_load_raw"].apply(parse_load_range)
    out["load_fraction_low"] = parsed.str[0]
    out["load_fraction_high"] = parsed.str[1]
    out["load_fraction_mid"] = parsed.str[2]

    out["running_hours_day"] = df["Running hours"].apply(to_number) if "Running hours" in df else np.nan
    out["diesel_l_day"] = df["Daily fuel consumption (litres/day)"].apply(
        to_number) if "Daily fuel consumption (litres/day)" in df else np.nan
    out["fuel_kg_day"] = df["Daily fuel consumption (kg/day)"].apply(
        to_number) if "Daily fuel consumption (kg/day)" in df else np.nan
    out["electric_kwh_day"] = df["Daily electricity consumption (kWh/day)"].apply(
        to_number) if "Daily electricity consumption (kWh/day)" in df else np.nan
    out["daily_consumption_raw"] = df["Daily consumption"] if "Daily consumption" in df else np.nan
    out["process_temperature_c"] = df["Required process temperature"].apply(
        to_number) if "Required process temperature" in df else np.nan

    # Input energy/day on common kWh basis.
    out["input_energy_kwh_day"] = np.nan
    elec = out["fuel"].eq("electric")
    diesel = out["fuel"].eq("diesel")
    pellet = out["fuel"].eq("pellet")
    rice = out["fuel"].eq("rice husk")

    out.loc[elec, "input_energy_kwh_day"] = out.loc[elec, "electric_kwh_day"]
    out.loc[diesel, "input_energy_kwh_day"] = out.loc[diesel, "diesel_l_day"] * DIESEL_KWH_PER_L
    out.loc[pellet, "input_energy_kwh_day"] = out.loc[pellet, "fuel_kg_day"] * PELLET_MJ_PER_KG / MJ_PER_KWH
    out.loc[rice, "input_energy_kwh_day"] = out.loc[rice, "fuel_kg_day"] * RICE_HUSK_MJ_PER_KG / MJ_PER_KWH

    out["input_energy_mj_day"] = out["input_energy_kwh_day"] * MJ_PER_KWH
    out["annual_input_energy_mwh"] = out["input_energy_kwh_day"] * np.nan

    # Merge company information.
    out = out.merge(company, on="submission_id", how="left", validate="many_to_one")

    # Annual energy uses company operating days when available.
    out["annual_input_energy_mwh"] = (
            out["input_energy_kwh_day"] * out["annual_operating_days"] / 1000.0
    )

    # Rated hours and implied average power.
    out["rated_energy_ceiling_kwh_day"] = out["capacity_kw_thermal"] * out["running_hours_day"]
    out["implied_average_power_kw"] = safe_div(
        out["input_energy_kwh_day"], out["running_hours_day"]
    )
    out["implied_load_fraction"] = safe_div(
        out["implied_average_power_kw"], out["capacity_kw_thermal"]
    )

    # If a load range is reported, put the midpoint/range alongside the implied load.
    out["reported_midpoint_vs_implied_ratio"] = safe_div(
        out["implied_load_fraction"], out["load_fraction_mid"]
    )

    # Contextual benchmark useful-heat estimate, clearly labeled as screening only.
    eff = out["fuel"].map(BENCHMARK_EFFICIENCY)
    out["screening_useful_heat_kwh_day"] = out["input_energy_kwh_day"] * eff
    out["screening_efficiency_assumption"] = out["fuel"].map(BENCHMARK_EFFICIENCY)

    return out


# =============================================================================
# VALIDATION / ANOMALY FLAGS
# =============================================================================

def add_anomaly_flags(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    flags = []

    for _, r in out.iterrows():
        row_flags = []
        cap = r["capacity_kw_thermal"]
        hrs = r["running_hours_day"]
        energy = r["input_energy_kwh_day"]
        implied = r["implied_average_power_kw"]
        load_mid = r["load_fraction_mid"]
        fuel = r["fuel"]

        if pd.isna(r["capacity_kw_thermal"]):
            row_flags.append("missing_capacity")
        if pd.isna(energy):
            row_flags.append("missing_daily_energy")
        if pd.isna(hrs):
            row_flags.append("missing_running_hours")

        # Hard physical consistency check where rated capacity and hours exist.
        if pd.notna(cap) and pd.notna(hrs) and pd.notna(energy) and cap > 0 and hrs > 0:
            ceiling = cap * hrs
            if energy > ceiling * HARD_OVERENERGY_FACTOR:
                row_flags.append("daily_energy_exceeds_rated_ceiling")

        # Implied average power should not exceed nameplate capacity by large margin.
        if pd.notna(cap) and cap > 0 and pd.notna(implied):
            if implied > cap * HIGH_IMPLIED_LOAD_FACTOR:
                row_flags.append("implied_power_exceeds_capacity")

        # If a mid-point operating-load label exists, compare it with the implied load.
        if pd.notna(implied) and pd.notna(cap) and pd.notna(load_mid) and cap > 0:
            if abs(implied / cap - load_mid) > 0.35:
                row_flags.append("reported_load_vs_implied_load_mismatch")

        # Combustion units missing the corresponding fuel quantity deserve review.
        if fuel in {"diesel", "pellet", "rice husk"}:
            if pd.isna(r["input_energy_kwh_day"]):
                row_flags.append("combustion_energy_missing")

        # Zero/negative values if supplied.
        for c in ["capacity_kw_thermal", "running_hours_day", "input_energy_kwh_day"]:
            v = r[c]
            if pd.notna(v) and v < 0:
                row_flags.append(f"negative_{c}")

        flags.append(row_flags)

    out["anomaly_flags"] = ["; ".join(x) if x else "" for x in flags]
    out["anomaly_flag_count"] = out["anomaly_flags"].apply(lambda x: 0 if not x else len(x.split("; ")))
    out["validation_status"] = np.where(
        out["anomaly_flag_count"] == 0,
        "screening_pass",
        np.where(
            out["anomaly_flags"].str.contains("exceeds_rated|implied_power_exceeds", regex=True),
            "needs_verification",
            "data_gap_or_review",
        ),
    )

    return out


# =============================================================================
# COMPANY ELECTRICAL SCREENING
# =============================================================================

def add_electrical_screening(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()

    # Sum reported electric unit consumption where available.
    elec = out["fuel"].eq("electric")
    electric_unit_daily = out["electric_kwh_day"].where(elec)
    company_elec_daily = electric_unit_daily.groupby(out["submission_id"]).transform("sum")
    out["sum_reported_electric_unit_kwh_day"] = company_elec_daily

    # Add new electric thermal load screening: if the whole thermal demand of this
    # unit were met by resistance heating, required average kW can be approximated
    # from input energy/day ÷ operating hours.
    out["screening_new_electric_average_kw"] = safe_div(
        out["screening_useful_heat_kwh_day"], out["running_hours_day"]
    )

    # Company-level available transformer headroom screening.
    out["transformer_headroom_kva_screening"] = (
            out["transformer_capacity_kva"] - out["maximum_demand_kw"]
    )
    out["transformer_headroom_pct_screening"] = safe_div(
        out["transformer_headroom_kva_screening"], out["transformer_capacity_kva"]
    )

    return out


# =============================================================================
# SUMMARY TABLES
# =============================================================================

def make_anomaly_table(df: pd.DataFrame) -> pd.DataFrame:
    cols = [
        "unit_id", "submission_id", "industry_name", "sector", "heating_arrangement", "fuel",
        "capacity_kw_thermal", "capacity_source", "running_hours_day", "input_energy_kwh_day",
        "rated_energy_ceiling_kwh_day", "implied_average_power_kw", "implied_load_fraction",
        "operating_load_raw", "process_temperature_c", "anomaly_flags", "validation_status"
    ]
    return df.loc[df["anomaly_flag_count"] > 0, [c for c in cols if c in df.columns]].copy()


def make_company_summary(df: pd.DataFrame) -> pd.DataFrame:
    records = []
    for sid, g in df.groupby("submission_id", dropna=False):
        first = g.iloc[0]
        records.append({
            "submission_id": sid,
            "industry_name": first["industry_name"],
            "location": first["location"],
            "printing_units": int((g["sector"] == "printing").sum()),
            "laminating_units": int((g["sector"] == "laminating").sum()),
            "total_units": int(len(g)),
            "total_input_energy_kwh_day_reported": g["input_energy_kwh_day"].sum(min_count=1),
            "total_input_energy_mwh_year_reported": g["annual_input_energy_mwh"].sum(min_count=1),
            "transformer_capacity_kva": first["transformer_capacity_kva"],
            "maximum_demand_kw": first["maximum_demand_kw"],
            "screening_transformer_headroom": first["transformer_headroom_kva_screening"],
            "spare_electrical_capacity": first["spare_electrical_capacity"],
            "space_available": first["space_available"],
            "willing_to_switch": first["willing_to_switch"],
        })
    return pd.DataFrame(records)


# =============================================================================
# PLOTS (Modernized with Seaborn)
# =============================================================================

def save_plot(fig, filename: str):
    fig.tight_layout()
    fig.savefig(PLOTS_DIR / filename, dpi=200, bbox_inches="tight")
    plt.close(fig)


def plot_capacity_vs_energy(df):
    d = df.dropna(subset=["capacity_kw_thermal", "input_energy_kwh_day"]).copy()
    fig, ax = plt.subplots(figsize=(9, 6.5))

    if not d.empty:
        sns.scatterplot(
            data=d,
            x="capacity_kw_thermal",
            y="input_energy_kwh_day",
            hue="fuel",
            s=80,
            alpha=0.85,
            edgecolor="w",
            ax=ax
        )
        x = np.linspace(max(d["capacity_kw_thermal"].min() * 0.8, 0.1), d["capacity_kw_thermal"].max() * 1.1, 100)
        # 10 h/day reference line is contextual only.
        ax.plot(x, x * 10, linestyle="--", color="gray", label="10 h/day at 100% rated load")

    ax.set_xlabel("Rated thermal capacity (kW-th)")
    ax.set_ylabel("Reported daily input energy (kWh/day)")
    ax.set_title("Rated Capacity vs Reported Daily Energy", pad=15)
    ax.legend(bbox_to_anchor=(1.05, 1), loc='upper left')
    save_plot(fig, "01_capacity_vs_daily_energy.png")


def plot_implied_power(df):
    d = df.dropna(subset=["capacity_kw_thermal", "implied_average_power_kw"]).copy()
    fig, ax = plt.subplots(figsize=(9, 6.5))

    if not d.empty:
        sns.scatterplot(
            data=d,
            x="capacity_kw_thermal",
            y="implied_average_power_kw",
            hue="fuel",
            s=80,
            alpha=0.85,
            edgecolor="w",
            ax=ax
        )
        xmax = max(d["capacity_kw_thermal"].max() * 1.1, 1)
        x = np.linspace(0, xmax, 100)
        ax.plot(x, x, linestyle="--", color="gray", label="100% nameplate")
        ax.plot(x, 1.2 * x, linestyle=":", color="red", label="120% nameplate")

    ax.set_xlabel("Rated thermal capacity (kW-th)")
    ax.set_ylabel("Implied average input power (kW-eq)")
    ax.set_title("Implied Average Power vs Rated Capacity", pad=15)
    ax.legend(bbox_to_anchor=(1.05, 1), loc='upper left')
    save_plot(fig, "02_implied_average_power_vs_capacity.png")


def plot_transformer(df):
    d = df.drop_duplicates("submission_id").dropna(subset=["transformer_capacity_kva", "maximum_demand_kw"]).copy()
    if d.empty:
        return

    fig, ax = plt.subplots(figsize=(9, 6.5))

    # Use regplot for a statistical trendline (Linear Regression)
    sns.regplot(
        data=d,
        x="transformer_capacity_kva",
        y="maximum_demand_kw",
        ax=ax,
        scatter_kws={"s": 80, "alpha": 0.7, "edgecolor": "w"},
        line_kws={"color": "orange", "ls": "-.", "label": "Demand Trendline"}
    )

    for _, r in d.iterrows():
        name = str(r["industry_name"])[:18]
        ax.annotate(name, (r["transformer_capacity_kva"], r["maximum_demand_kw"]), xytext=(5, 5),
                    textcoords="offset points", fontsize=9)

    xmax = d["transformer_capacity_kva"].max() * 1.1
    x = np.linspace(0, xmax, 100)
    ax.plot(x, x, linestyle="--", color="gray", label="Demand = transformer rating")

    ax.set_xlabel("Existing transformer capacity (kVA)")
    ax.set_ylabel("Maximum demand (reported, kW)")
    ax.set_title("Transformer Capacity vs Maximum Demand", pad=15)
    ax.legend()
    save_plot(fig, "03_transformer_vs_max_demand.png")


def plot_daily_energy_by_sector_fuel(df):
    d = df.dropna(subset=["input_energy_kwh_day"]).copy()
    if d.empty:
        return

    g = d.groupby(["sector", "fuel"], as_index=False)["input_energy_kwh_day"].sum()

    fig, ax = plt.subplots(figsize=(9, 6.5))
    sns.barplot(
        data=g,
        x="fuel",
        y="input_energy_kwh_day",
        hue="sector",
        ax=ax,
        edgecolor="black",
        linewidth=1
    )

    ax.set_ylabel("Reported input energy (kWh/day)")
    ax.set_xlabel("Fuel / energy source")
    ax.set_title("Reported Daily Input Energy by Sector and Fuel", pad=15)
    ax.legend(title="Sector")
    save_plot(fig, "04_daily_energy_by_sector_and_fuel.png")


def plot_operating_hours(df):
    d = df.dropna(subset=["running_hours_day"])
    if d.empty:
        return

    fig, ax = plt.subplots(figsize=(9, 6.5))

    # Histogram with Kernel Density Estimate overlay
    sns.histplot(
        data=d,
        x="running_hours_day",
        bins=8,
        kde=True,
        color=sns.color_palette("deep")[0],
        ax=ax
    )

    ax.set_xlabel("Running hours per day")
    ax.set_ylabel("Number of units")
    ax.set_title("Distribution of Reported Unit Operating Hours", pad=15)
    save_plot(fig, "05_operating_hours.png")


def plot_capacity_utilization(df):
    d = df.dropna(subset=["implied_load_fraction"]).copy()
    d = d[(d["implied_load_fraction"] >= 0) & (d["implied_load_fraction"] <= 2)].reset_index(drop=True)
    if d.empty:
        return

    fig, ax = plt.subplots(figsize=(10, 6.5))
    sns.barplot(
        x=d.index,
        y=d["implied_load_fraction"] * 100,
        color=sns.color_palette("deep")[2],
        ax=ax,
        edgecolor="w"
    )
    ax.axhline(100, linestyle="--", color="red", label="100% nameplate")

    ax.set_ylabel("Implied utilization / load (%)")
    ax.set_xlabel("Unit (ordered as in clean table)")
    ax.set_title("Screening Check: Implied Average Load as % of Rated Capacity", pad=15)
    ax.legend()
    # Hide x-axis labels if there are too many units
    ax.set_xticks([])
    save_plot(fig, "06_capacity_utilization_screen.png")


# =============================================================================
# EXCEL OUTPUT
# =============================================================================

def write_excel(clean: pd.DataFrame, anomalies: pd.DataFrame, company_summary: pd.DataFrame):
    output_file = OUTPUT_DIR / "clean_per_unit.xlsx"
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        clean.to_excel(writer, sheet_name="clean_per_unit", index=False)
        anomalies.to_excel(writer, sheet_name="anomalies", index=False)
        company_summary.to_excel(writer, sheet_name="company_summary", index=False)

        # Assumptions/definitions sheet makes the database auditable.
        assumptions = pd.DataFrame({
            "parameter": [
                "kcal/hr to kW",
                "diesel kWh/L",
                "pellet MJ/kg",
                "rice husk MJ/kg",
                "MJ/kWh",
                "note",
            ],
            "value": [
                KCAL_PER_HR_TO_KW,
                DIESEL_KWH_PER_L,
                PELLET_MJ_PER_KG,
                RICE_HUSK_MJ_PER_KG,
                MJ_PER_KWH,
                "Input-energy basis only; verify heating values and efficiencies during field audit.",
            ],
        })
        assumptions.to_excel(writer, sheet_name="assumptions", index=False)

    return output_file


# =============================================================================
# MAIN
# =============================================================================

def main():
    OUTPUT_DIR.mkdir(exist_ok=True)
    PLOTS_DIR.mkdir(exist_ok=True)

    print("\n=== AEPC Fuel-Switching Survey Data Foundation ===\n")
    print(f"Input configuration folder: {SCRIPT_DIR.resolve()}")

    company_raw, printing_raw, laminating_raw = load_workbook(INPUT_FILE)
    company = prepare_company(company_raw)

    printing = prepare_units(printing_raw, "printing", company)
    laminating = prepare_units(laminating_raw, "laminating", company)

    clean = pd.concat([printing, laminating], ignore_index=True)
    clean = add_anomaly_flags(clean)
    clean = add_electrical_screening(clean)

    # Useful sorting for review.
    clean = clean.sort_values(["industry_name", "sector", "unit_index"], na_position="last").reset_index(drop=True)

    anomalies = make_anomaly_table(clean)
    company_summary = make_company_summary(clean)

    # Write CSV outputs.
    clean_csv = OUTPUT_DIR / "clean_per_unit.csv"
    anomalies_csv = OUTPUT_DIR / "anomalies.csv"
    company_csv = OUTPUT_DIR / "company_summary.csv"
    clean.to_csv(clean_csv, index=False)
    anomalies.to_csv(anomalies_csv, index=False)
    company_summary.to_csv(company_csv, index=False)

    # Workbook output.
    excel_out = write_excel(clean, anomalies, company_summary)

    # Plots.
    plot_capacity_vs_energy(clean)
    plot_implied_power(clean)
    plot_transformer(clean)
    plot_daily_energy_by_sector_fuel(clean)
    plot_operating_hours(clean)
    plot_capacity_utilization(clean)

    # Validation summary for quick human review.
    summary_file = OUTPUT_DIR / "validation_summary.txt"
    with summary_file.open("w", encoding="utf-8") as f:
        f.write("AEPC Fuel-Switching Survey Validation Summary\n")
        f.write("=============================================\n\n")
        f.write(f"Companies: {clean['submission_id'].nunique()}\n")
        f.write(f"Units: {len(clean)}\n")
        f.write(f"Printing units: {(clean['sector'] == 'printing').sum()}\n")
        f.write(f"Laminating units: {(clean['sector'] == 'laminating').sum()}\n")
        f.write(f"Units needing review: {(clean['anomaly_flag_count'] > 0).sum()}\n")
        f.write("\nFuel counts:\n")
        f.write(clean["fuel"].value_counts(dropna=False).to_string())
        f.write("\n\nHard anomaly records:\n")
        hard = clean[
            clean["anomaly_flags"].str.contains("daily_energy_exceeds_rated_ceiling|implied_power_exceeds_capacity",
                                                na=False)]
        if hard.empty:
            f.write("None\n")
        else:
            f.write(hard[
                        ["unit_id", "industry_name", "capacity_kw_thermal", "running_hours_day", "input_energy_kwh_day",
                         "implied_average_power_kw", "anomaly_flags"]].to_string(index=False))

    print("\nCreated:")
    print(f"  {excel_out}")
    print(f"  {clean_csv}")
    print(f"  {anomalies_csv}")
    print(f"  {company_csv}")
    print(f"  {summary_file}")
    print(f"  plots/ ({len(list(PLOTS_DIR.glob('*.png')))} PNG files)")

    print("\n=== Validation Snapshot ===")
    print(f"Companies: {clean['submission_id'].nunique()}")
    print(f"Units: {len(clean)}")
    print(f"Needs verification/review: {(clean['anomaly_flag_count'] > 0).sum()}")
    if not anomalies.empty:
        print("\nFlagged records:")
        print(anomalies[["unit_id", "industry_name", "fuel", "capacity_kw_thermal", "running_hours_day",
                         "input_energy_kwh_day", "anomaly_flags"]].to_string(index=False))


if __name__ == "__main__":
    main()