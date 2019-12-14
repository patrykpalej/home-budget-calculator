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

v.1.6

_BugFix_
- Limitting number of incomes and earnings on a piechart for a total and year scenario
- adding food subcategories structure for part and total analysis

_ResultsGeneration_
- Scatterplot of spendings vs. incomes for the subsequent months in year and total scenarios
- For year and total scenario summary tables added to xlsx file. The first table contains mean, median and std for spendings, earnings, incomes and surplus. The second table contains total surplus during the period
- Lineplot of basic, additional and giftdon spendings for the whole period added to year and total scenario

v.1.7

_BugFix_
- Dates range added to total output xlsx

_SpendingsFinder_
- for the set of given phrases script generates a report of when spendings with that phrase occured and what was the category

v.1.8

*path_from_txt*

- implement reading data and results directory from .txt file
- divide code into functions

*beautify*

- order imports by length

- delete hardcoded sys.args

- shorten lines to 79 chars

  