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

# Set Seaborn theme for professional plots
sns.set_theme(style="whitegrid", context="talk", font_scale=0.8, palette="deepWhat kind of program would you like me to write? Please let me know what you want the program to do and what programming language you prefer (e.g., Python, C++, Java, JavaScript).