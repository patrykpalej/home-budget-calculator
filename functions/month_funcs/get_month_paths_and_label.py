import os


def get_month_paths_and_label(month_num, year_num):

    folder_path_file = open("path.txt", "r")
    folder_path = folder_path_file.read()
    folder_path_file.close()

    file_path = folder_path + "/monthly/20" + year_num + "." + month_num \
        + ".xlsx"
    month_label = "20" + year_num + "." + month_num

    results_dir = folder_path + "/!Raporty/monthly_reports/" + month_label \
        + " - wyniki"
    if not os.path.exists(results_dir):
        os.mkdir(results_dir)
        os.mkdir(results_dir + "/plots/")

    return folder_path, file_path, month_label, results_dir
