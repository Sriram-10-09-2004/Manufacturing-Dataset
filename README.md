                          ManufactureIQ — Manufacturing Production
# Project Title :
Manufacturing Dataset

## Project Overview :
This project analyzes 1,000 manufacturing records using Excel to evaluate production, efficiency, quality, downtime, maintenance, costs, revenue, and profit. It includes data analysis, KPI calculations, PivotTables, charts, slicers, and interactive dashboards to identify operational trends, performance issues, and areas for improvement.

## Tools :
Excel

## Project Workflow :
## Day 1 – Dataset Setup and Excel Table Creation
Used Excel's duplicate-removal/checking functionality on `Record_ID`.
Confirmed that there were no duplicate Record_ID values in the dataset.
Used the column filters to work with the different manufacturing fields.
Confirmed that the numerical columns were available for calculations, including Planned Quantity, Produced Quantity, Good Quantity, Defect Quantity, Scrap Quantity, Downtime, Cycle Time, Temperature, Vibration, Energy Consumption, Raw Material Cost, Labour Hours, Maintenance Days, and Revenue.
Kept the original source columns unchanged and prepared additional columns on the right side of the table for calculated metrics.

## Day 2 – Creating Production and Quality Calculated Columns
Added a new column named `Production_Achievement_%`.
In the first data row, entered:
`=K2/J2`
This divided Produced Quantity by Planned Quantity.
Pressed Enter and filled the formula down through all 1,000 records.
Added `Defect_Rate_%`.
Entered:
`=M2/K2`
This calculated Defect Quantity as a percentage of Produced Quantity.
Filled the formula down.
Added `Scrap_Rate_%`.
Entered:
`=N2/K2`
Filled the formula down.
Added `Good_Rate_%`.
Entered:
`=L2/K2`
Filled the formula down.
Selected the percentage columns and used Excel's Percentage formatting.
Adjusted the percentage display to show the required decimal places.
Added `Downtime_Hours`.
Entered:
`=O2/60`
This converted Downtime Minutes into Downtime Hours.
Filled the formula down through the dataset.

## Day 3 – Production Time, Operating Time and Availability
Added `Production_Time_Hours`.
Entered:
`=(K2*R2)/3600`
This used Produced Quantity and Actual Cycle Time to calculate production time in hours.
Filled the formula down.
Added `Planned_Production_Time`.
Entered:
`=(J2*Q2)/3600`
This used Planned Quantity and Ideal Cycle Time to calculate the planned production time.
Filled the formula down.
Added `Operating_Time_Hours`.
Entered:
`=MAX(0,AI2-AE2)`
This calculated operating time by subtracting downtime hours from planned production time.
The `MAX(0,...)` portion prevented negative operating-time values.
Filled the formula down.
Created the `Availability_%` column.
Entered:
`=IFERROR(AH2/AI2,0)`
This divided Operating Time by Planned Production Time.
Used `IFERROR` so that an error would return 0 instead of displaying an Excel error.
Filled the formula down.
Formatted Availability as a percentage.

## Day 4 – Actual Production Rate, Performance and OEE
Created `Actual_Production_Rate`.
Entered:
`=IFERROR(3600/R2,0)`
This calculated the production rate from Actual Cycle Time.
Filled the formula down.
Created `Ideal_Production_Rate`.
Entered:
`=IFERROR(3600/Q2,0)`
This calculated the ideal production rate from Ideal Cycle Time.
Filled the formula down.
Created `Performance_%`.
Entered:
`=IFERROR(AF2/AK2,0)`
This compared Actual Production Rate with Ideal Production Rate.
Filled the formula down.
Created `Quality_%`.
Entered:
`=IFERROR(L2/K2,0)`
This calculated the proportion of produced units that were good units.
Filled the formula down.
Created `OEE_%`.
Entered:
`=AJ2*AL2*AM2`
This multiplied Availability, Performance, and Quality.
Filled the formula down.
Formatted Availability, Performance, Quality, and OEE as percentages.
Selected the calculated columns and adjusted the decimal formatting so the percentages were displayed consistently.

## Day 5 – Production Analysis Using PivotTables
Selected a cell inside the `Manufacturing_Data` table.
Went to:
`Insert → PivotTable`
Selected the `Manufacturing_Data` table as the source.
Created a new worksheet for the PivotTable.
For Production Achievement by Production Line:
Dragged `Production_Line` into the `Rows` area.
Dragged `Production_Achievement_%` into the `Values` area.
Opened the Values field settings.
Changed the calculation from Sum to Average.
Formatted the resulting values as percentages.
Used the PivotTable results to compare the production achievement of each production line.
Repeated the PivotTable process for Machine performance.
Placed `Machine_ID` in `Rows`.
Placed `Production_Achievement_%` in `Values`.
Changed the Values calculation to Average.
Used the results to compare production achievement between machines.
Created a production-output PivotTable.
Placed `Product_Type` in `Rows`.
Placed `Produced_Quantity` in `Values`.
Kept the calculation as Sum.
Created another production-output PivotTable using `Shift` in Rows and `Produced_Quantity` in Values.
Used these PivotTables to compare production output across products and shifts.

## Day 6 – Downtime and Cycle-Time Analysis
Created a PivotTable for machine downtime.
Placed `Machine_ID` in the `Rows` area.
Placed `Downtime_Hours` in the `Values` area.
Kept the Values calculation as Sum.
Formatted the downtime values to display two decimal places.
Sorted the machine downtime values to identify the machines contributing the most downtime.
Created another PivotTable using `Production_Line` in Rows and `Downtime_Hours` in Values.
Created another using `Shift` in Rows and `Downtime_Hours` in Values.
Created another using `Downtime_Reason` in Rows and `Downtime_Hours` in Values.
Used Sum of Downtime Hours for each of these analyses.
Compared Ideal Cycle Time and Actual Cycle Time by machine.
Created a PivotTable with `Machine_ID` in Rows.
Added `Ideal_Cycle_Time_Sec` to Values.
Added `Actual_Cycle_Time_Sec` to Values.
Changed both Value calculations to Average.
Compared the two cycle-time columns to identify machines where actual cycle time was higher than ideal cycle time.
Applied Conditional Formatting to the `Downtime_Hours` column.
Used a color scale to make higher downtime values visually easier to identify.

## Day 7 – Quality, Maintenance and Risk Analysis
Created PivotTables for quality analysis.
For defect rate by machine:
Placed `Machine_ID` in Rows.
Placed `Defect_Rate_%` in Values.
Changed the calculation to Average.
Formatted the result as a percentage.
For quality by product:
Placed `Product_Type` in Rows.
Placed `Quality_%` in Values.
Changed the calculation to Average.
Formatted the result as a percentage.
Created quality comparisons by Production Line and Shift using the same PivotTable method.
Applied Conditional Formatting to `Defect_Rate_%`.
Used a color scale so higher defect-rate records could be identified visually.
Created the `Maintenance_Risk` column.
Used maintenance age, downtime, machine temperature, and machine vibration as the risk conditions.
Entered the following formula:
`=IF((Y2>25)+(O2>25)+(S2>73)+(T2>2)>=3,"High",IF((Y2>25)+(O2>25)+(S2>73)+(T2>2)=2,"Medium","Low"))`
Filled the formula down through all records.
This classified each record as High, Medium, or Low maintenance risk.
Created a PivotTable for Maintenance Risk.
Placed `Machine_ID` in Rows.
Placed `Maintenance_Risk` in Columns.
Placed `Record_ID` in Values.
Changed Record_ID from Sum to Count.
This produced the number of High, Medium, and Low risk records for each machine.

## Day 8 – Cost, Profit and Anomaly Analysis
Created `Production_Cost`.
Used the available Raw Material Cost because the dataset did not contain a separate labour-cost rate.
Entered:
`=V2`
Filled the formula down.
Created `Cost_Per_Unit`.
Entered:
`=IFERROR(AP2/K2,0)`
Filled the formula down.
Created `Profit`.
Entered:
`=Z2-AP2`
Filled the formula down.
Created `Profit_Per_Unit`.
Entered:
`=IFERROR(AR2/K2,0)`
Filled the formula down.
Created `Scrap_Cost`.
Entered:
`=IFERROR(N2*AQ2,0)`
Filled the formula down.
Applied Conditional Formatting to the cost columns to visually identify higher-cost records.
Created PivotTables for Scrap Cost by Machine, Production Line, and Product Type.
Used `Scrap_Cost` as the Values field and kept the calculation as Sum.
Created PivotTables for Profit by Machine, Production Line, and Product Type.
Used `Profit` as the Values field and kept the calculation as Sum.
Created a Cost Per Unit PivotTable by Machine.
Placed `Machine_ID` in Rows.
Placed `Cost_Per_Unit` in Values.
Changed the calculation to Average.
Created the `Anomaly_Flag` column.
Entered:
`=IF(AN2<70%,"Critical",IF(AN2<85%,"Warning","Normal"))`
Filled the formula down.
This classified OEE records into Critical, Warning, and Normal categories.

## Day 9 – Monthly Trends and Dashboard Creation
Created monthly production analysis using the existing `Production_Date` field.
Created a PivotTable with:
`Production_Date → Rows`
`Produced_Quantity → Values`
Changed the date field to group the records by Months and Years.
Used Sum of Produced Quantity for the monthly production values.
Created a monthly downtime PivotTable using:
`Production_Date → Rows`
`Downtime_Hours → Values`
Grouped the date field by Months and Years.
Created a monthly OEE PivotTable using:
`Production_Date → Rows`
`OEE_% → Values`
Changed OEE to Average.
Formatted the OEE values as percentages.
Created the `Operations Dashboard`.
Used PivotTables and PivotCharts to display production-line, product, downtime, shift, cycle-time, and OEE analysis.
Created the `Quality & Cost Dashboard`.
Added KPI values for:
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

## Day 10 – Management Dashboard, Slicers and Finalization
Created the `Management Dashboard` worksheet.
Merged cells `A1:F1` for the dashboard title.
Entered:
`MANAGEMENT DASHBOARD`
Formatted the title using bold white text and a dark background.
Created eight KPI cards.
Added Total Production.
Added Production Achievement %.
Added OEE %.
Added Defect Rate %.
Added Scrap Cost.
Added Downtime Hours.
Added Revenue.
Added Profit.
Created the following Management Dashboard PivotCharts:
`Production Trend`
Used Production_Date in Rows and Produced_Quantity in Values.
Grouped Production_Date by Months and Years.
Changed the chart type to Line.
Set the chart title to `Production Trend`.
`OEE by Line`
Placed Production_Line in Rows.
Placed OEE_% in Values.
Changed the calculation to Average.
Formatted the values as percentages.
Created a Clustered Column chart titled `OEE by Line`.
`Downtime by Machine`
Placed Machine_ID in Rows.
Placed Downtime_Hours in Values.
Used Sum.
Created a Clustered Bar chart titled `Downtime by Machine`.
`Defect Rate by Product`
Placed Product_Type in Rows.
Placed Defect_Rate_% in Values.
Changed the calculation to Average.
Created a Clustered Column chart titled `Defect Rate by Product`.
`Scrap Cost by Line`
Placed Production_Line in Rows.
Placed Scrap_Cost in Values.
Used Sum.
Formatted the values as Indian Rupee currency.
Created a Clustered Column chart titled `Scrap Cost by Line`.
`Production by Shift`
Placed Shift in Rows.
Placed Produced_Quantity in Values.
Used Sum.
Formatted the values with comma separators.
Created a Clustered Column chart titled `Production by Shift`.
`Maintenance Risk`
Placed Machine_ID in Rows.
Placed Maintenance_Risk in Columns.
Placed Record_ID in Values.
Changed Record_ID from Sum to Count.
Created a Clustered Column chart titled `Maintenance Risk`.
Added five slicers to make the dashboard interactive.
Added the grouped `Months (Production_Date)` field as a slicer and renamed it `Month`.
Added `Production_Line` as a slicer and renamed it `Production Line`.
Added `Machine_ID` as a slicer and renamed it `Machine`.
Added `Shift` as a slicer and renamed it `Shift`.
Added `Product_Type` as a slicer and renamed it `Product Type`.
Used the PivotTable connection settings to connect the slicers to the relevant PivotTables.
Formatted the dashboards by adjusting chart positions, titles, spacing, number formats, percentage formats, and currency formats.
Adjusted PivotTable settings so column widths would not automatically change after updates.
Prepare the final business insights and conclusions from the completed analysis.
Prepared recommendations based on machine reliability, production-line performance, downtime, quality, maintenance risk, shift performance, and scrap cost.
Created the `AI_Usage_Log` worksheet.
Recorded the AI tool used, prompts, purpose, AI output summary, Excel validation, and final decision for the AI-assisted parts of the project.


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
