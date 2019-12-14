import os


def get_year_parameters(year_num):
    folder_path_file = open("path.txt", "r")
    folder_path = folder_path_file.read()
    folder_path_file.close()

    file_path = folder_path + "/yearly/20" + year_num + ".xlsx"
    year_label = "20" + year_num

    results_dir = folder_path + "/!Raporty/yearly_reports/" + year_label \
        + " - wyniki"
    if not os.path.exists(results_dir):
        os.mkdir(results_dir)
        os.mkdir(results_dir + "/plots")

    return folder_path, file_path, year_label, results_dir
