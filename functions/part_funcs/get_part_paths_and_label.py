import os
import sys


def get_part_paths_and_label(start_month, start_year, end_month, end_year):

    folder_path_file = open("path.txt", "r")
    folder_path = folder_path_file.read()
    folder_path_file.close()

    file_path = folder_path + "/total/total.xlsx"
    part_label = str(start_month) + ".20" + str(start_year) + "-" \
        + str(end_month) + ".20" + str(end_year)

    if (start_year > end_year) or \
            (start_year == end_year and start_month > end_month):
        print("Podany zakres dat jest błędny")
        sys.exit()
    else:
        # generating list of sheetnames for a given range
        list_of_sheetnames = []
        n_of_months = 12 * (end_year - start_year) \
            + end_month - start_month + 1
        month = start_month
        year = start_year
        for i in range(n_of_months):
            name_of_sheet = ("0" + str(month) if month <= 9 else str(month)) \
                            + "." + str(year)
            list_of_sheetnames.append(name_of_sheet)

            month += 1
            if month > 12:
                month -= 12
                year += 1

    start_label = [int(list_of_sheetnames[0][:2]),
                   int(list_of_sheetnames[0][-2:])]

    results_dir = folder_path + "/!Raporty/total_partial/" + part_label \
        + " - wyniki/"
    if not os.path.exists(results_dir):
        os.mkdir(results_dir)
        os.mkdir(results_dir + "/plots")

    return folder_path, file_path, part_label, results_dir, start_label, \
        list_of_sheetnames
