import os
import math

from classes import MyWorkbook


def get_total_parameters():
    folder_path_file = open("path.txt", "r")
    folder_path = folder_path_file.read()
    folder_path_file.close()

    file_path = folder_path + "/total data/total.xlsx"

    my_workbook = MyWorkbook(file_path)
    my_worksheets = my_workbook.mywb.sheetnames

    n_of_months = len(my_workbook.sheets_list)
    start_label = [int(my_worksheets[0][:2]), int(my_worksheets[0][-2:])]
    end_label = [start_label[0] + n_of_months % 12 - 1,
                 start_label[1] + math.floor(n_of_months / 12)]

    total_label = "Total (" + ("0" if start_label[0] <= 9 else "") \
                  + str(start_label[0]) + "." + str(start_label[1]) + " - " \
                  + ("0" if end_label[0] <= 9 else "") + str(
        end_label[0]) + "." + str(end_label[1]) + ")"

    results_dir = folder_path + "/!Reports/total_partial_results/" + \
        total_label + " - wyniki/"
    if not os.path.exists(results_dir):
        os.mkdir(results_dir)
        os.mkdir(results_dir + "/plots")

    return folder_path, file_path, total_label, results_dir, my_workbook, \
        my_worksheets, start_label
