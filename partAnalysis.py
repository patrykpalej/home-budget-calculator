import os
import sys
import shutil

from classes import MyWorkbook
from functions.part_funcs.get_part_parameters import get_part_parameters
from functions.part_funcs.create_all_plots import create_all_plots
from functions.part_funcs.create_pptx_presentation \
    import create_pptx_presentation
from functions.part_funcs.create_xlsx_report import create_xlsx_report

start_month = int(sys.argv[1])
start_year = int(sys.argv[2])
end_month = int(sys.argv[3])
end_year = int(sys.argv[4])


# Preparing the analysis
folder_path, file_path, part_label, results_dir, start_label, \
    list_of_sheetnames \
    = get_part_parameters(start_month, start_year, end_month, end_year)

myWorkbook = MyWorkbook(file_path, list_of_sheetnames)


# Actual analysis
if not os.path.exists(results_dir + "/plots/"):
    os.mkdir(results_dir + "/plots/")

spendings_list, incomes_list, plot_numbers_list \
    = create_all_plots(myWorkbook, part_label, results_dir, start_label,
                       len(list_of_sheetnames))

create_pptx_presentation(part_label, results_dir, plot_numbers_list)

shutil.rmtree(results_dir + "/plots/")

create_xlsx_report(results_dir, part_label, myWorkbook, start_label,
                   len(list_of_sheetnames))
