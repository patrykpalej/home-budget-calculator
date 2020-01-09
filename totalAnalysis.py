import os
import shutil

from functions.total_funcs.get_total_parameters import get_total_parameters
from functions.total_funcs.create_all_plots import create_all_plots
from functions.total_funcs.create_pptx_presentation \
    import create_pptx_presentation
from functions.total_funcs.create_xlsx_report import create_xlsx_report


# Preparing the analysis
folder_path, file_path, total_label, results_dir, myWorkbook, myWorksheets, \
    start_label = get_total_parameters()


# Actual analysis
if not os.path.exists(results_dir + "/plots/"):
    os.mkdir(results_dir + "/plots/")

spendings_list, incomes_list, earnings_list, surplus_list = \
    create_all_plots(myWorkbook, myWorksheets, total_label, results_dir,
                     start_label, len(myWorksheets))

create_pptx_presentation(total_label, results_dir)

shutil.rmtree(results_dir + "/plots/")

create_xlsx_report(results_dir, total_label, spendings_list, incomes_list,
                   earnings_list, surplus_list, myWorkbook, start_label,
                   len(myWorksheets))
