import os
import sys

from classes import MyWorkbook
from functions.yearFuncs import create_all_plots, create_pptx_presentation, \
    create_xlsx_report


year_num = sys.argv[1]
# year_num = "99"  # hardcoded if no arguments

# Preparing the analysis
folder_path_file = open("path.txt", "r")
folder_path = folder_path_file.read()
folder_path_file.close()

file_path = folder_path + "/yearly/20" + year_num + ".xlsx"
year_label = "20" + year_num

myWorkbook = MyWorkbook(file_path)
myWorksheets = myWorkbook.mywb.sheetnames
start_label = [int(myWorksheets[0][:2]), int(year_num)]

results_dir = folder_path + "/!Raporty/yearly_reports/" + year_label \
              + " - wyniki"
if not os.path.exists(results_dir):
    os.mkdir(results_dir)
    os.mkdir(results_dir + "/plots")


# Actual analysis
spendings_list, incomes_list, earnings_list, surplus_list \
    = create_all_plots(myWorkbook, myWorksheets, year_label, results_dir,
                       start_label)

create_pptx_presentation(year_num, results_dir)

create_xlsx_report(results_dir, year_num, spendings_list, incomes_list,
                   earnings_list, surplus_list, myWorkbook, start_label)

