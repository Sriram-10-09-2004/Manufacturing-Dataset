                                  ManufactureIQ — Manufacturing Production
# Project Title :
Manufacturing Dataset

## Project Overview :
This project analyzes 1,000 manufacturing records using Microsoft Excel to evaluate production performance, efficiency, quality, downtime, maintenance, costs, revenue, and profit. The project uses Excel formulas and calculated KPIs such as Production Achievement, Availability, Performance, Quality, and OEE, along with PivotTables and PivotCharts to analyze machines, production lines, products, shifts, downtime, maintenance risks, and financial performance. The results are presented through three interactive dashboards with slicers for Month, Production Line, Machine, Shift, and Product Type, helping identify important trends, operational issues, and areas for improvement.

## Tools :
Excel

## Project Workflow :
## Day 1 – Data Preparation and KPI Calculations
Opened `Manufacturing_Data_Analytics_1000_Records.xlsx` in Microsoft Excel.
Selected the complete dataset containing 1,000 manufacturing records.
Pressed `Ctrl + T` to convert the dataset into an Excel Table.
Enabled `My table has headers`.
Opened the `Table Design` tab and renamed the table to `Manufacturing_Data`.
Used `Record_ID` to identify duplicate records and confirmed that there were no duplicate values.
Created the calculated columns from `AA` onwards.
Created `Production_Achievement_%` and entered:
`=K2/J2`
Filled the formula down to all records.
Created `Defect_Rate_%`:
`=M2/K2`
Created `Scrap_Rate_%`:
`=N2/K2`
Created `Good_Rate_%`:
`=L2/K2`
Filled each formula down through the dataset.
Created `Downtime_Hours`:
`=O2/60`
Created `Actual_Production_Rate`:
`=IFERROR(3600/R2,0)`
Created `Production_Time_Hours`:
`=(K2*R2)/3600`
Created `Planned_Production_Time`:
`=(J2*Q2)/3600`
Created `Operating_Time_Hours`:
`=MAX(0,AI2-AE2)`
Created `Availability_%`:
`=IFERROR(AH2/AI2,0)`
Created `Ideal_Production_Rate`:
`=IFERROR(3600/Q2,0)`
Created `Performance_%`:
`=IFERROR(AF2/AK2,0)`
Created `Quality_%`:
`=IFERROR(L2/K2,0)`
Created `OEE_%`:
`=AJ2*AL2*AM2`
Filled all calculated formulas down to the final record.
Selected the percentage columns and applied Percentage formatting with the required decimal places.

## Day 2 – Production, Downtime, OEE and Quality Analysis
Selected a cell inside the `Manufacturing_Data` table.
Went to `Insert → PivotTable`.
Created PivotTables to analyze production performance.
For Production Achievement by Production Line, placed `Production_Line` in the `Rows` area and `Production_Achievement_%` in the `Values` area.
Opened `Value Field Settings` and changed the calculation from `Sum` to `Average`.
Formatted the result as a percentage.
Repeated the same process using `Machine_ID` in Rows to compare production achievement by machine.
Created a production-output PivotTable by placing `Product_Type` in Rows and `Produced_Quantity` in Values.
Kept `Produced_Quantity` as `Sum`.
Created another PivotTable using `Shift` in Rows and `Produced_Quantity` in Values.
Created the machine downtime PivotTable by placing `Machine_ID` in Rows and `Downtime_Hours` in Values.
Kept `Downtime_Hours` as `Sum`.
Created additional downtime PivotTables using `Production_Line`, `Shift`, and `Downtime_Reason` in Rows.
Created a cycle-time PivotTable.
Placed `Machine_ID` in Rows.
Placed `Ideal_Cycle_Time_Sec` and `Actual_Cycle_Time_Sec` in Values.
Changed both fields to `Average`.
Compared actual cycle time against ideal cycle time for each machine.
Created OEE analysis by placing `Production_Line` in Rows and `OEE_%` in Values.
Changed OEE to `Average`.
Formatted OEE as a percentage.
Created quality PivotTables using `Machine_ID`, `Product_Type`, `Production_Line`, and `Shift`.
Used `Defect_Rate_%` and `Quality_%` as Values and changed the calculations to `Average`.
Applied Conditional Formatting to `Defect_Rate_%` using a color scale to make higher defect rates easier to identify.
Applied Conditional Formatting to `Downtime_Hours` using a color scale to identify higher downtime records.

## Day 3 – Maintenance, Risk, Cost, Profit and Trend Analysis
Created the `Maintenance_Risk` calculated column.
Entered:
`=IF((Y2>25)+(O2>25)+(S2>73)+(T2>2)>=3,"High",IF((Y2>25)+(O2>25)+(S2>73)+(T2>2)=2,"Medium","Low"))`
Filled the formula down.
Created a PivotTable for Maintenance Risk.
Placed `Machine_ID` in Rows.
Placed `Maintenance_Risk` in Columns.
Placed `Record_ID` in Values.
Changed `Record_ID` from `Sum` to `Count`.
This produced High, Medium, and Low maintenance-risk record counts for every machine.
Created the `Anomaly_Flag` column.
Entered:
`=IF(AN2<70%,"Critical",IF(AN2<85%,"Warning","Normal"))`
Filled the formula down.
Created PivotTables to analyze maintenance status against downtime, defect rate, and OEE.
Created `Production_Cost` using the available Raw Material Cost field.
Entered:
`=V2`
Created `Cost_Per_Unit`:
`=IFERROR(AP2/K2,0)`
Created `Profit`:
`=Z2-AP2`
Created `Profit_Per_Unit`:
`=IFERROR(AR2/K2,0)`
Created `Scrap_Cost`:
`=IFERROR(N2*AQ2,0)`
Filled all formulas down.
Applied Conditional Formatting to the cost-related columns.
Created PivotTables for Scrap Cost by Machine, Production Line, and Product Type.
Used `Scrap_Cost` in Values and kept the calculation as `Sum`.
Created PivotTables for Profit by Machine, Production Line, and Product Type.
Used `Profit` in Values and kept the calculation as `Sum`.
Created a Cost Per Unit PivotTable by placing `Machine_ID` in Rows and `Cost_Per_Unit` in Values.
Changed Cost Per Unit to `Average`.
Created monthly trend PivotTables using `Production_Date`.
Placed `Production_Date` in Rows and `Produced_Quantity` in Values.
Used Excel's date grouping option to group Production_Date by Months and Years.
Created similar monthly PivotTables for Downtime Hours and OEE.
For OEE, changed the Values calculation to `Average` and formatted it as a percentage.
Used the monthly PivotTables to compare production, downtime, and OEE from January to June 2026.

## Day 4 – Dashboard Creation, Slicers and Finalization
Created three dashboard worksheets:
`Operations Dashboard`
`Quality & Cost Dashboard`
`Management Dashboard`
For the Operations Dashboard, created PivotCharts from the analysis PivotTables.
Created charts for production by Production Line and Product Type, along with the required operational analysis charts.
For the Quality & Cost Dashboard, created KPI cards for:
Total Defect Rate
Total Scrap Rate
Average Quality
Total Scrap Cost
Total Profit
Average Cost Per Unit
Created PivotCharts for:
Defect Rate by Machine
Scrap Cost by Machine
Quality by Product
Profit by Production Line
Cost Per Unit by Machine
Created the Management Dashboard.
Merged `A1:F1` for the dashboard title and entered `MANAGEMENT DASHBOARD`.
Created eight KPI cards:
Total Production — 692,181
Production Achievement — 93.52%
OEE — 86.36%
Defect Rate — 6.49%
Scrap Cost — ₹3,01,043.70
Downtime Hours — 420.95
Revenue — ₹25,038,055.78
Profit — ₹9,355,415.61
Created the seven Management Dashboard charts:
`Production Trend`
`OEE by Line`
`Downtime by Machine`
`Defect Rate by Product`
`Scrap Cost by Line`
`Production by Shift`
`Maintenance Risk`
For each chart, created or used the corresponding PivotTable, selected the required Rows and Values fields, selected the appropriate chart type, and entered the final chart title.
Added five slicers.
Added the grouped `Months (Production_Date)` field as a slicer and renamed it `Month`.
Added `Production_Line` as a slicer and renamed it `Production Line`.
Added `Machine_ID` as a slicer and renamed it `Machine`.
Added `Shift` as a slicer and renamed it `Shift`.
Added `Product_Type` as a slicer and renamed it `Product Type`.
Used `Report Connections` to connect the slicers to the relevant PivotTables.
Formatted the dashboard charts and KPI cards.
Applied comma formatting to production quantities.
Applied percentage formatting to production achievement, OEE, defect rate, quality, and related KPIs.
Applied Indian Rupee formatting to revenue, profit, scrap cost, and other monetary values.
Adjusted PivotTable settings so column widths would not automatically resize after updates.
Created the `AI_Usage_Log` worksheet.
Recorded the AI tool, prompt, purpose, AI output summary, Excel validation, and final decision for the AI-assisted analysis.
Prepared the final business insights, conclusions, and recommendations based on the completed Excel analysis.

## Analysis :
Data cleaning and validation
Production achievement analysis
Production analysis by production line, machine, product type, and shift
Downtime analysis by machine, production line, shift, and downtime reason
Cycle time analysis comparing ideal and actual cycle time
OEE analysis
Availability, performance, and quality analysis
Defect rate, scrap rate, and good production analysis
Quality analysis by machine, product, line, and shift
Maintenance status and maintenance risk analysis
Maintenance days vs downtime, defects, and OEE analysis
Production cost and cost-per-unit analysis
Scrap cost analysis
Revenue and profit analysis
Profit-per-unit analysis
Monthly production, downtime, and OEE trend analysis
Anomaly and high-risk machine identification
Conditional formatting for identifying high-defect, high-downtime, and high-cost records

## Key Insights
Total production reached 692,181 units during January–June 2026.
Production Line 4 achieved the highest production achievement at 98.74%.
Machine M303 recorded the highest downtime at 39.92 hours and the lowest OEE at 73.91%.
Overall OEE was 86.36%, with availability at 95.82%, performance at 96.32%, and quality at 93.51%.
Machine Breakdown was the leading downtime reason, accounting for 106.57 hours.
Line 3 showed the lowest production achievement at 87.28% and the lowest line-level OEE at 82.07%.
The Night shift had the lowest performance at 69.15% and quality at 92.89%.
Overall defect rate was 6.49%, while scrap rate was 1.94%.
Total scrap cost was ₹301,043.70, with Line 3 having the highest scrap cost.
M303 showed unusually high maintenance risk, with 52 high-risk records out of 59.
Monthly OEE remained relatively stable, ranging from 85.82% to 86.70%.
The analysis identified machine reliability, downtime, quality losses, and shift performance as key areas for improvement.

## Author
Sriram Aditya
