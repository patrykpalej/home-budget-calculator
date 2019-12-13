import os
import math

from classes import MyWorkbook
from functions.totalFuncs import create_all_plots, create_pptx_presentation, \
    create_xlsx_report


# Preparing the analysis
file_path = os.getcwd() + "/data/total.xlsx"

myWorkbook = MyWorkbook(file_path)
myWorksheets = myWorkbook.mywb.sheetnames
n_of_months = len(myWorkbook.sheets_list)
start_label = [int(myWorksheets[0][:2]), int(myWorksheets[0][-2:])]
end_label = [start_label[0] + n_of_months%12 - 1,
             start_label[1] + math.floor(n_of_months/12)]

total_label = "Total (" + ("0" if start_label[0] <= 9 else "") \
              + str(start_label[0]) + "." + str(start_label[1]) + " - " \
              + ("0" if end_label[0] <= 9 else "") + str(end_label[0]) + "." \
              + str(end_label[1]) + ")"

results_dir = os.getcwd() + "/results/" + total_label + " - wyniki"
if not os.path.exists(results_dir):
    os.mkdir(results_dir)
    os.mkdir(results_dir + "/plots")


# Actual analysis
spendings_list, incomes_list, earnings_list, surplus_list = \
    create_all_plots(myWorkbook, myWorksheets, total_label, results_dir,
                     start_label, n_of_months)

create_pptx_presentation(total_label, results_dir)

create_xlsx_report(results_dir, total_label, spendings_list, incomes_list,
                   earnings_list, surplus_list, myWorkbook, start_label,
                   n_of_months)
