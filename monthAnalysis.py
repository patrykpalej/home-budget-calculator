import os
import sys
import shutil

from classes import MyWorkbook
from functions.month_funcs.get_month_parameters import get_month_parameters
from functions.month_funcs.create_all_plots import create_all_plots
from functions.month_funcs.create_pptx_presentation \
    import create_pptx_presentation

month_num = sys.argv[1]
year_num = sys.argv[2]


# Preparing the analysis
folder_path, file_path, month_label, results_dir \
    = get_month_parameters(month_num, year_num)

myWorkbook = MyWorkbook(file_path)
myWorksheet = myWorkbook.sheets_list[0]

# Actual analysis
if not os.path.exists(results_dir + "/plots/"):
    os.mkdir(results_dir + "/plots/")

create_all_plots(myWorksheet, month_label, results_dir)

create_pptx_presentation(month_num, year_num, results_dir)

shutil.rmtree(results_dir + "/plots/")