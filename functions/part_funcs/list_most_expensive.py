import openpyxl
from math import floor


def list_most_expensive(results_dir, part_label, my_workbook, start_label,
                        n_of_months):

    wb_to_export = openpyxl.Workbook()

    # Data preparation
    all_spends_values = my_workbook.spends_values_yr
    all_spends_items = my_workbook.spends_items_yr
    all_spends_months = my_workbook.spends_monthlabel_yr

    high_spends_values = []
    high_spends_items = []
    high_spends_dates = []
    high_spends_categories = []

    value_treshold = 100

    for values, items, months, cat_name in zip(all_spends_values.values(),
                                               all_spends_items.values(),
                                               all_spends_months.values(),
                                               all_spends_values.keys()):
        for value, item, month in zip(values, items, months):
            if value >= value_treshold and cat_name not in ["Mieszkanie",
                                                            "Jedzenie"]:
                high_spends_values.append(value)
                high_spends_items.append(item)

                month_label = (month + start_label[0]) % 12
                year_label = start_label[1] \
                    + floor((month-1 + start_label[0] - 1) / 12)
                high_spends_dates.append(str(month_label) + ".20" +
                                         str(year_label))

                high_spends_categories.append(cat_name)

    # Excel filling
    ws = wb_to_export.active

    ws.cell(1, 1).value = "id"
    ws.cell(1, 2).value = "Nazwa"
    ws.cell(1, 3).value = "Wartość"
    ws.cell(1, 4).value = "Kategoria"
    ws.cell(1, 5).value = "Miesiąc"

    for row in range(len(high_spends_values)):
        ws.cell(2+row, 1).value = row+1
        ws.cell(2+row, 2).value = high_spends_items[row]
        ws.cell(2+row, 3).value = high_spends_values[row]
        ws.cell(2+row, 4).value = high_spends_categories[row]
        ws.cell(2+row, 5).value = high_spends_dates[row]

    wb_to_export.save(results_dir + "/" + part_label
                      + " - największe wydatki.xlsx")