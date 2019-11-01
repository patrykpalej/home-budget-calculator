import pandas as pd
import os
from classes import *
from openpyxl.styles import Alignment
from openpyxl.styles.borders import Border, Side
from openpyxl.styles import PatternFill


# 1. Importing config file and total data
config_txt = pd.read_csv("config/spends_to_find.txt", header=None)
config_list = list(config_txt[0])

spendings_file_path = os.getcwd() + "/data/total.xlsx"
myWorkbook = MyWorkbook(spendings_file_path)
spend_values = myWorkbook.spends_values_yr
spend_items = myWorkbook.spends_items_yr

# 2. Preparation of list of encoded dictionaries of spendings
spends_dicts_list = list()
trial_dict = {"cat_sums": [["Mieszkanie i przyjemności", 250], ["Podróże",
														20], ["Ubrania", 150]],
			  "date_spends": {"Styczeń 2053": [["szermierka", 12],
											   ["spodnie", 120],
											   ["audiobook", 40]],
							  "Marzec 2053": [["herbaciarnia", 12],
											  ["teatr", 34]],
							  "Lipiec 2053": [["Wyjście na piwo", 20]],
							  "Luty 20154": [["katowice, przejazd", 1],
											 ["wydatki", 14]]
							  }
			  }
spends_dicts_list.append(trial_dict)

# 3. Creating xlsx file for the results
output_xlsx = openpyxl.Workbook()
ws = output_xlsx.active
ws.title = "Słowo"

# a) cat_sums table
thin_border = Border(left=Side(style='thin'),
                     right=Side(style='thin'),
                     top=Side(style='thin'),
                     bottom=Side(style='thin'))

try:
	ws.column_dimensions["A"].width = max([len(cat[0]) for cat in
										   trial_dict["cat_sums"]]) + 2
except:
	pass

row = 1
for cat in trial_dict["cat_sums"]:
	ws.cell(row, 1).value = cat[0]
	ws.cell(row, 2).value = cat[1]

	ws.cell(row, 1).border = thin_border
	ws.cell(row, 2).border = thin_border

	row += 1

# b) date_spends table


# 4. Postprocessing
results_dir = os.getcwd() + "/spendings_finder"
if not os.path.exists(results_dir):
	os.mkdir(results_dir)

output_xlsx.save(results_dir + "/Wydatki.xlsx")
