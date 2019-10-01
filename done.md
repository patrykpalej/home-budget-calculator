v.1.0

- monthAnalysis.py which includes a scenario for a single month analysis was written. Plots and reports generation implemented
- MyWorksheet class for parsing a single month sheet was implemented
- first plotting functions were implemented
- main.py initialized

v.1.1

- first version of yearAnalysis.py scenario implemented
- sumAllSheets() method added to class MyWorkbook to sum all sheets (months) in a one year workbook 

v.1.2

 - year Analysis divided into three sections - year as a whole, average month and sequence of months
 - 4 new plots added for the year as a sequence of months part
 - sumAllSheets() method deleted and the code included to the constructor

v.1.3

 - small changes in year- and month analysis
 - totalAnalysis created
 - piechart design changed

v.1.4

_YearSpendingsTable_
 - Changed way of using sys.argv
 - Better description of the columns in spendings table in year scenario
 - Colors and monthly sums added to the table

_PlotsModification_
 - New plot added for year and total scenario - incomes and spendings in subsequent months (not cummulated)
 - For lineplots x tick labels changed to dates

_TotalTable_
 - Spendings table for the total scenario added

v.1.5

_PartAnalysis_
- Created new file: partAnalysis.py, code from totalAnalysis.py copied
- Additional plots added
- Choosing the period to analyse implemented
- Readme updated
- Variables names adjusted 
