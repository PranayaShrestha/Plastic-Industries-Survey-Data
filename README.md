# Plastic-Industries-Survey-Data
This script processes and cleans survey data collected from the printing and packaging industry. It takes raw Excel survey data, standardizes the energy consumption values across different fuel types, flags potential data entry errors or physical impossibilities, and generates analytical visualizations. 
The primary goal is to establish a verified "input-energy baseline" to help evaluate the feasibility of switching these facilities to electric heating.

1. Prerequisites and Setup
To run this script, you will need Python installed on your computer along with a few standard data science libraries.

Open your terminal or command prompt and install the required packages using:
pip install pandas numpy matplotlib seaborn openpyxl

2. How to Use the Script
Place the Files: Put the Python script (e.g., process_survey.py) and the raw Excel survey workbook in the exact same folder.

Run the Script: Execute the script via your terminal or IDE.

The script is designed to automatically detect the Excel workbook in its folder, even if the filename changes slightly upon download (it looks for keywords like "thermic", "printing", or "survey").

View Results: Once finished, the script will create a new folder named fuel_switching_outputs containing all the generated reports and charts.

3. Understanding the Outputs
The script organizes all generated files into a new directory called fuel_switching_outputs. Below is a detailed explanation of everything you will find inside.

Data Tables (Spreadsheets & CSVs)
clean_per_unit.xlsx
This is the master output file containing multiple sheets for easy viewing in Excel:

clean_per_unit: The fully processed dataset. Every printing and laminating unit is listed as its own row. All fuel consumption (liters of diesel, kg of pellets, etc.) has been mathematically converted into a standard kWh/day input energy metric for easy comparison.

anomalies: A filtered list showing only the units that failed validation checks.

company_summary: A high-level overview aggregated by factory/company.

assumptions: A reference sheet detailing the exact conversion factors used (e.g., how much energy is assumed per liter of diesel).

clean_per_unit.csv
A flat, comma-separated version of the complete, cleaned unit data. Ideal for importing into PowerBI, Tableau, or other database software.

company_summary.csv
A factory-level summary showing total energy consumption across all units at a site, existing transformer capacity, maximum electrical demand, and available electrical headroom.

anomalies.csv
A critical file for data quality assurance. If a unit's reported energy consumption physically exceeds what its heater could produce, or if crucial data is missing, it is flagged here with a specific error code (e.g., daily_energy_exceeds_rated_ceiling).

Text Reports
validation_summary.txt
A quick-read text file summarizing the health of the dataset. It tells you exactly how many companies and units were processed, the breakdown of fuel types currently in use, and highlights any "hard anomalies" (extreme data entry errors) that require immediate human review.

Visualizations (plots/ folder)
The script automatically generates modernized, publication-ready charts (PNG format) to help visualize the state of the industry:

01_capacity_vs_daily_energy.png
A scatter plot comparing each unit's rated machine capacity (kW-thermal) against its daily energy use. Helps identify if machines of similar sizes are using wildly different amounts of fuel.

02_implied_average_power_vs_capacity.png
Shows the "implied" average power (total daily energy divided by operating hours) versus the nameplate capacity. Points that fall far above the dashed reference lines indicate likely data reporting errors (a machine cannot average 150% of its maximum capacity all day).

03_transformer_vs_max_demand.png
A factory-level scatter plot comparing the site's reported maximum electrical demand against its total transformer capacity. Includes a statistical trendline. Points near the dashed line have no spare electrical capacity; points far below the line have room to add new electric heaters.

04_daily_energy_by_sector_and_fuel.png
A bar chart showing total energy consumption broken down by the type of fuel (diesel, electric, pellet, rice husk) and the sector (printing vs. laminating).

05_operating_hours.png
A histogram showing the distribution of factory operating hours. Helps determine standard operating shifts (e.g., identifying if most factories run 8-hour, 12-hour, or 24-hour schedules).

06_capacity_utilization_screen.png
A bar chart showing the implied utilization rate of every unit. Any bar spiking above the red 100% line immediately flags a unit where the reported fuel consumption outpaces the physical limits of the stated machine size.

4. Important Engineering Note
This script translates all reported fuels into a common input-energy basis (kWh/day) using standard calorific values.

It does not calculate the final "useful heat" delivered to the process. For example, a diesel boiler might consume 100 kWh of fuel energy, but only deliver 80 kWh of useful heat to the laminator due to exhaust and standby losses. When sizing new electric equipment, engineers must apply appropriate thermal efficiency factors to these baseline numbers.
