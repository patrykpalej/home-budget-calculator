import pandas as pd
import os
import openpyxl


# 1. Importing config file
config_txt = pd.read_csv("config/spends_to_find.txt", header=None)
config_list = list(config_txt[0])

# 2. Creating xlsx file for the results
output_xlsx = openpyxl.Workbook()

# 3. Postprocessing
results_dir = os.getcwd() + "/spendings_finder"
if not os.path.exists(results_dir):
    os.mkdir(results_dir)

output_xlsx.save(results_dir + "/Wydatki.xlsx")

