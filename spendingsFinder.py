import json
import os
from classes import *
from openpyxl.styles import Alignment
from openpyxl.styles.borders import Border, Side
from openpyxl.styles import PatternFill


# 1. Importing config file and total data
with open('spendings_finder/config.json') as handle:
    config_json = json.loads(handle.read())

config_list = config_json["keywords"]
period = config_json["period"]

spendings_file_path = os.getcwd() + "/data/total_test.xlsx"
myWorkbook = MyWorkbook(spendings_file_path)
spend_values = myWorkbook.spends_values_yr
spend_items = myWorkbook.spends_items_yr


# 2. Preparation of list of encoded dictionaries of spendings
spends_dicts_list = list()
phrase = "test"
# -----
good_dict = {"cat_sums": {}, "date_spends": {}}

for cat in spend_items.keys():
    for i, item in enumerate(spend_items[cat]):
        if phrase.lower() in item.lower():
            if cat in good_dict["cat_sums"].keys():
                good_dict["cat_sums"][cat] += spend_values[cat][i]
            else:
                good_dict["cat_sums"][cat] = spend_values[cat][i]


# -----
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

thin_border = Border(left=Side(style='thin'),
                     right=Side(style='thin'),
                     top=Side(style='thin'),
                     bottom=Side(style='thin'))

# a) cat_sums table
try:
    ws.column_dimensions["A"].width = 2 + max([len(cat[0])
                                               for cat in
                                               trial_dict["cat_sums"]])
except:
    pass

row = 1
for cat in trial_dict["cat_sums"]:
    ws.cell(row, 1).value = cat[0]
    ws.cell(row, 2).value = cat[1]

    ws.cell(row, 1).border = thin_border
    ws.cell(row, 2).border = thin_border

    ws.cell(row, 1).fill = PatternFill(fgColor='daf6f7', fill_type='solid')

    row += 1

# b) date_spends table
row = 1
for date in list(trial_dict["date_spends"].keys()):
    ws.merge_cells(start_row=row, start_column=8, end_row=row, end_column=9)
    ws.cell(row, 8).value = date
    ws.cell(row, 8).alignment = Alignment(horizontal='center')
    ws.cell(row, 8).fill = PatternFill(fgColor='93e1e6', fill_type='solid')
    ws.cell(row, 8).border = thin_border
    row += 1

    ws.cell(row, 8).value = "Co?"
    ws.cell(row, 9).value = "Kwota [zł]"
    ws.cell(row, 8).alignment = Alignment(horizontal='center')
    ws.cell(row, 9).alignment = Alignment(horizontal='center')
    ws.cell(row, 8).fill = PatternFill(fgColor='daf6f7', fill_type='solid')
    ws.cell(row, 9).fill = PatternFill(fgColor='daf6f7', fill_type='solid')
    ws.cell(row, 8).border = thin_border
    ws.cell(row, 9).border = thin_border

    row += 1
    max_length = 0
    for spend in trial_dict["date_spends"][date]:
        ws.cell(row, 8).value = spend[0]
        ws.cell(row, 9).value = spend[1]
        ws.cell(row, 8).border = thin_border
        ws.cell(row, 9).border = thin_border

        max_length = max(max_length, len(spend[0]))
        row += 1

try:
    ws.column_dimensions["H"].width = max_length + 2
except:
    pass


# 4. Postprocessing
results_dir = os.getcwd() + "/spendings_finder"
if not os.path.exists(results_dir):
    os.mkdir(results_dir)

output_xlsx.save(results_dir + "/Przedział czasu.xlsx")
