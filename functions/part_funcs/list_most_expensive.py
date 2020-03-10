import openpyxl


def list_most_expensive(results_dir, part_label, my_workbook, start_label,
                        n_of_months):

    wb_to_export = openpyxl.Workbook()

    wb_to_export.save(results_dir + "/" + part_label
                      + " - największe wydatki.xlsx")