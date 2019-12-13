import os
import sys

from classes import MyWorkbook
from functions.month_funcs.create_all_plots import create_all_plots
from functions.month_funcs.create_pptx_presentation \
    import create_pptx_presentation


month_num = sys.argv[1]
year_num = sys.argv[2]
# month_num = "01"  # hardcoded if no arguments
# year_num = "99"  # hardcoded if no arguments

# Preparing the analysis
folder_path_file = open("path.txt", "r")
folder_path = folder_path_file.read()
folder_path_file.close()

file_path = folder_path + "/monthly/20" + year_num + "." + month_num + ".xlsx"
month_label = "20" + year_num + "." + month_num

myWorkbook = MyWorkbook(file_path)
myWorksheet = myWorkbook.sheets_list[0]

results_dir = folder_path + "/!Raporty/monthly_reports/" + month_label \
              + " - wyniki"
if not os.path.exists(results_dir):
    os.mkdir(results_dir)
    os.mkdir(results_dir + "/plots/")


# Actual analysis

create_all_plots(myWorksheet, month_label, results_dir)

create_pptx_presentation(month_num, year_num, results_dir)
