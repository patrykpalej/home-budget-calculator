import os
import sys
import shutil

from classes import MyWorkbook
from functions.month_funcs.get_month_parameters import get_month_parameters
from functions.month_funcs.create_all_plots import create_all_plots
from functions.month_funcs.create_pptx_presentation \
    import create_pptx_presentation

try:
    month_num = sys.argv[1]
    year_num = sys.argv[2]
except IndexError:
    folder_path_file = open("path.txt", "r")
    folder_path = folder_path_file.read()
    folder_path_file.close()
    folder_path += "/monthly data/"
    second_newest = os.listdir(folder_path)[-2]

    year_num = second_newest[:2]
    month_num = second_newest[5:7]


# Preparing the analysis
folder_path, file_path, month_label, results_dir \
    = get_month_parameters(month_num, year_num)

myWorkbook = MyWorkbook(file_path)
myWorksheet = myWorkbook.sheets_list[0]

# Actual analysis
if not os.path.exists(results_dir + "/plots/"):
    os.mkdir(results_dir + "/plots/")

plot_numbers_list = create_all_plots(myWorksheet, month_label, results_dir)

create_pptx_presentation(month_num, year_num, results_dir, plot_numbers_list)

shutil.rmtree(results_dir + "/plots/")
