This is a project of a home budget calculator. Data is stored in excel worksheets. Sample data can be found in a "data" folder. There are different analysis scenarios. The output is a powerpoint presentation which is automatically generated. The main need-to-know information is described below. 


1. Names of the files and what they are for
-------------------------------------------
- main.py		- a potential main script for a complex scenario
- monthAnalysis.py	- main script for a single month analysis scenario
- yearAnalysis.py 	- main script for a single year analysis scenario
- totalAnalysis.py      - main script for the analysis of the whole period
- partAnalysis.py       - main script for the analysis of a chosen period
- classes.py 		- contains classes such as MyWorksheet or MyWorkbook whose instances are single worksheets/workbooks
- plotFuncs.py		- file with functions for plotting data

-------------------------------------------------------------------------------------
-------------------------------------------------------------------------------------

2. How to run
-------------
To run the month analysis enter "python monthAnalysis.py mm yy" to cmd
where mm and yy stand for month and year respectively. 

For example "python monthAnalysis.py 04 18" is for April 2018 

There is sample data attached to the code for Jan, Feb and Mar 2099.
- - -
To run the year analysis enter "python yearAnalysis.py yy" 
- - -
To run the total analysis enter "python totalAnalysis.py" without arguments
- - -
To run the partial analysis enter "python partAnalysis.py mm_start yy_year mm_end yy_end"

For example "python partAnalysis.py 02 18 10 19" is for the period from Feb 2018 to Oct 2018
 
-------------------------------------------------------------------------------------
-------------------------------------------------------------------------------------

3. Dictionary
-------------
- categories - arbitrarily chosen categories of spendings such as "Jedzenie" or "Rzeczy i sprzêty"
- metacategories - high-level categories indicating type of spendings. There are three: "Podstawowe", "Dodatkowe","Donacje/prezenty"
- subcategories - to specify type of the spendings for category "Jedzenie" there are subcategories included to indicate what kind of food is this
- incomes - money that comes in a given month
- earnings - money that is earned in a given month
- source - on of incomes/earnings sources
- ballance - difference between incomes/earnings and spendings

-------------------------------------------------------------------------------------
-------------------------------------------------------------------------------------

4. Code conventions
-------------------
- variables starting with "_" are temporary 
