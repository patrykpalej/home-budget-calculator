import sys

from classes import MyWorkbook
from functions.year_funcs.get_year_parameters import get_year_parameters
from functions.year_funcs.create_all_plots import create_all_plots
from functions.year_funcs.create_pptx_presentation \
    import create_pptx_presentation
from functions.year_funcs.create_xlsx_report import create_xlsx_report

year_num = sys.argv[1]


# Preparing the analysis
folder_path, file_path, year_label, results_dir \
    = get_year_parameters(year_num)

myWorkbook = MyWorkbook(file_path)
myWorksheets = myWorkbook.mywb.sheetnames
start_label = [int(myWorksheets[0][:2]), int(year_num)]


# Actual analysis
spendings_list, incomes_list, earnings_list, surplus_list \
    = create_all_plots(myWorkbook, myWorksheets, year_label, results_dir,
                       start_label)

create_pptx_presentation(year_num, results_dir)

create_xlsx_report(results_dir, year_num, spendings_list, incomes_list,
                   earnings_list, surplus_list, myWorkbook, start_label)
