import pandas as pd
import os
from classes import *


# 1. Importing config file and total data
config_txt = pd.read_csv("config/spends_to_find.txt", header=None)
config_list = list(config_txt[0])

spendings_file_path = os.getcwd() + "/data/total.xlsx"
myWorkbook = MyWorkbook(spendings_file_path)
spend_values = myWorkbook.spends_values_yr
spend_items = myWorkbook.spends_items_yr

# 2. Preparation of list of encoded dictionaries of spendings


# 3. Creating xlsx file for the results
output_xlsx = openpyxl.Workbook()
ws = output_xlsx.active
ws.title = "Przedział"


# 4. Postprocessing
results_dir = os.getcwd() + "/spendings_finder"
if not os.path.exists(results_dir):
    os.mkdir(results_dir)

output_xlsx.save(results_dir + "/Wydatki.xlsx")
