# Home Budget Calculator 

This is a project of a home budget calculator. Data is stored in excel worksheets. Sample data can be found in a "data" folder. There are different analysis scenarios. The output is a powerpoint presentation which is automatically generated. The main need-to-know information is described below.  // !!! info about storing results in another directory // !!!

### 1. Names of the files and what they are for 

- monthAnalysis.py - main script for a single month analysis scenario
- yearAnalysis.py - main script for a single year analysis scenario
- totalAnalysis.py - main script for the analysis of the whole registered period
- partAnalysis.py - main script for the analysis of a chosen period
- main.py - a main script for a potential complex scenario
- spendingsFinder.py  - script which searches for a set of phrases in all spendings for a given period and presents them in tables
- classes.py - contains classes such as MyWorksheet or MyWorkbook whose instances are single worksheets/workbooks
- plotFuncs.py - file with functions for plotting data
- functions/
  - plotFuncs.py - file with functions for plots generation
  - monthFuncs.py - file with functions used in month scenario
  - yearFuncs.py - file with functions used in year scenario
  - totalFuncs.py - file with functions used in total scenario
  - ...



### 2. How to run

#### 2.1. Month 

To run the month analysis enter "python monthAnalysis.py mm yy" to the console where mm and yy stand for month and year respectively. 

For example "python monthAnalysis.py 04 18" is for April 2018. 

There is sample data attached to the code for Jan, Feb and Mar 2099 in data/monthly directory.
#### 2.2. Year 

To run the year analysis enter "python yearAnalysis.py yy" 

For example "python yearAnalysis 19" is for 2019. 

There is sample data attached to the code for 2099 in data/yearly directory.

#### 2.3. Total  

To run the total analysis enter "python totalAnalysis.py" without any arguments.

There is sample data attached to the code in data/ directory.

#### 2.4 Part

To run the partial analysis enter "python partAnalysis.py mm_start yy_year mm_end yy_end"

For example "python partAnalysis.py 02 18 10 19" is for the period from Feb 2018 to Oct 2018
#### 2.5 Spendings finder

To run a spendings finder enter "python spendingsFinder.py" without any arguments. Update keywords in config.json in spendings_finder/ directory beforehand.



### 3. Dictionary

- categories - arbitrarily chosen categories of spendings. They are specified in worksheets with input data and can be modified
- metacategories - high-level arbitrarily chosen categories indicating type of spendings. There are three: one for basic spendings, one for additional (non-basic) and one for gifts, donations etc.
- subcategories - to specify type of the spendings for food category there are subcategories included to indicate what kind of food it is
- incomes - money that incomes in a given month
- earnings - money that is earned in a given month
- sources - sources of incomes/earnings sources
- balance - difference between incomes/earnings and spendings



### 4. Code conventions

- variables starting with "_" are temporary 



### 5. More information

The project is described in details on the blog: 

https://ds-ml.blog/2019/09/10/kalkulator-budzetu-domowego-1-wprowadzenie/

