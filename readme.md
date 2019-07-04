1. Names of the files and what they are for
-------------------------------------------
- main.py		a potential main script for a complex scenario
- monthAnalysis.py	main script for a single month analysis scenario
- classes.py 		contains classes such as MyWorksheet or MyWorkbook whose instances are single worksheets/workbooks
- plotFuncs.py		file with functions for plotting data

-------------------------------------------------------------------------------------
-------------------------------------------------------------------------------------

2. How to run
-------------
To run the month analysis type in the cmd "python monthAnalysis.py mm yy"
where mm and yy stand for month and year respectively. 

For example "python monthAnalysis.py 04 18" is for April 2018 

There is sample data attached to the code for Jan, Feb and Mar 2099.
 
-------------------------------------------------------------------------------------
-------------------------------------------------------------------------------------

3. Dictionary
----------
- categories - arbitrarily chosen categories of spendings such as "Jedzenie" or "Rzeczy i sprzêty"
- metacategories - high-level categories indicating type of spendings. There are three: "Podstawowe", "Dodatkowe","Donacje/prezenty"
- subcategories - to specify type of the spendings for category "Jedzenie" there are subcategories included to indicate what kind of food is this
- incomes - money that come in a given month
- earnings - money that is earned in a given month
- source - on of incomes/earnings sources
- ballance - difference between incomes/earnings and spendings

-------------------------------------------------------------------------------------
-------------------------------------------------------------------------------------

4. Code
-------
- variables starting with "_" are temporary 
