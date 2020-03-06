import numpy as np
import matplotlib.pyplot as plt

from functions.plotFuncs import plotPie, plotBar


def create_all_plots(my_worksheet, month_label, results_dir):
    plot_nr = 1

    # -- Barplot for all categories
    # region
    index_order = np.flip(np.argsort(my_worksheet.cats_sums_list), axis=0)
    values_desc = [my_worksheet.cats_sums_list[i] for i in index_order
                   if my_worksheet.cats_sums_list[i] > 0]
    labels_desc = [my_worksheet.cats_names[i] for i in index_order
                   if my_worksheet.cats_sums_list[i] > 0]

    values = values_desc
    labels = labels_desc
    title = month_label + " - Kwoty wydawane miesięcznie \n na kolejne " \
                          "kategorie\n"
    fig_name = results_dir + "/plots/plot{}.png".format(plot_nr)
    plot_nr += 1

    fig = plotBar(values, labels, title)
    plt.savefig(figure=fig, fname=fig_name)
    # endregion

    # -- Piechart of spendings for the main categories
    # region
    index_order = np.flip(np.argsort(my_worksheet.cats_sums_list), axis=0)
    _top_indices = index_order[0:5]
    _low_indices = index_order[5:len(index_order)]

    top_labels = [my_worksheet.cats_names[i] for i in _top_indices]
    top_values = [my_worksheet.cats_sums_list[i] for i in _top_indices]
    others_values = [my_worksheet.cats_sums_list[i] for i in _low_indices]

    top_plus_others_values = top_values + [sum(others_values)]
    top_plus_others_labels = top_labels + ["Pozostałe"]
    top_plus_others_labels = [top_plus_others_labels[i] + " - "
                              + str(round(top_plus_others_values[i], 2))+" zł"
                              for i in range(len(top_plus_others_values))]

    values = top_plus_others_values
    labels = top_plus_others_labels
    title = month_label + " - Struktura miesięcznych wydatków\n z podziałem" \
                          " na kategorie\n\n" + "Suma wydatków: " \
                          + str(round(my_worksheet.sum_total, 2)) + "zł\n"
    fig_name = results_dir + "/plots/plot{}.png".format(plot_nr)
    plot_nr += 1

    fig = plotPie(values, labels, title)
    plt.savefig(figure=fig, fname=fig_name)
    # endregion

    # -- Piechart of the metacategories
    # region
    metacats_values = [my_worksheet.sum_basic, my_worksheet.sum_addit,
                       my_worksheet.sum_giftdon]
    metacats_labels = [
        "Podstawowe - " + str(round(my_worksheet.sum_basic, 2)) + "zł",
        "Dodatkowe - " + str(round(my_worksheet.sum_addit, 2)) + "zł",
        "Prezenty i donacje - "
        + str(round(my_worksheet.sum_giftdon, 2)) + "zł"]

    values = metacats_values
    labels = metacats_labels
    title = month_label + " - Podział wydatków na: \n" \
        + "Podstawowe, Dodatkowe i Prezenty/Donacje\n\n" + "Suma wydatków: " \
        + str(round(my_worksheet.sum_total, 2)) + "zł\n"
    fig_name = results_dir + "/plots/plot{}.png".format(plot_nr)
    plot_nr += 1

    fig = plotPie(values, labels, title)
    plt.savefig(figure=fig, fname=fig_name)
    # endregion

    # -- Piechart of incomes
    # region
    _values_list = list(my_worksheet.incomes_dict.values())
    _labels_list = list(my_worksheet.incomes_dict.keys())
    incomes_values = [inc for inc in _values_list if inc > 0]
    incomes_labels = [inc + " - " + str(_values_list[i]) + "zł" for i, inc in
                      enumerate(_labels_list) if _values_list[i] > 0]

    values = incomes_values
    labels = incomes_labels
    title = month_label + " - Podział przychodów na poszczególne źródła\n\n" \
        + "Suma przychodów: " + str(my_worksheet.incomes) + "zł\n" \
        + "Nadwyżka przychodów: " + str(round(my_worksheet.balance[0], 2)) \
        + "zł (" + str(round(100 * my_worksheet.balance[0]
                             / my_worksheet.incomes, 2)) + "%)\n"
    fig_name = results_dir + "/plots/plot{}.png".format(plot_nr)
    plot_nr += 1

    fig = plotPie(values, labels, title)
    plt.savefig(figure=fig, fname=fig_name)
    # endregion

    # -- Piechart of earnings
    # region
    _values_list = list(my_worksheet.earnings_dict.values())
    _labels_list = list(my_worksheet.earnings_dict.keys())
    earnings_values = [ear for ear in _values_list if ear > 0]
    earnings_labels = [ear + " - " + str(_values_list[i]) + "zł" for i, ear in
                       enumerate(_labels_list) if _values_list[i] > 0]

    values = earnings_values
    labels = earnings_labels
    title = month_label + " - Podział zarobków na poszczególne źródła\n\n" \
        + "Suma zarobków: " + str(my_worksheet.earnings) + "zł\n" \
        + "Nadwyżka zarobków: " + str(round(my_worksheet.balance[1], 2)) \
        + "zł (" + str(round(100 * my_worksheet.balance[1]
                             / my_worksheet.earnings, 2)) + "%)\n"
    fig_name = results_dir + "/plots/plot{}.png".format(plot_nr)
    plot_nr += 1

    fig = plotPie(values, labels, title)
    plt.savefig(figure=fig, fname=fig_name)
    # endregion

    # -- Piechart of food subcategories
    # region
    amounts = my_worksheet.spends_values["Jedzenie"]
    subcats = my_worksheet.spends_items["Jedzenie"]
    subcats_dict = {}

    for i, subcat in enumerate(subcats):
        if subcat in list(subcats_dict.keys()):
            subcats_dict[subcat] += amounts[i]
        else:
            subcats_dict[subcat] = amounts[i]

    subcats_values = list(subcats_dict.values())
    subcats_labels = [list(subcats_dict.keys())[i] + " - "
                      + str(round(list(subcats_dict.values())[i], 2)) + "zł"
                      for i, sc in enumerate(list(subcats_dict.keys()))]

    _subcats_fractions = [sc / sum(subcats_values) for sc in subcats_values]
    subcats_values_with_others = []
    subcats_labels_with_others = []
    others_sum = 0
    for i, sc_f in enumerate(_subcats_fractions):
        if sc_f > 0.02:
            subcats_values_with_others.append(subcats_values[i])
            subcats_labels_with_others.append(subcats_labels[i])
        else:
            others_sum += subcats_values[i]

    if others_sum > 0:
        subcats_values_with_others.append(others_sum)
        subcats_labels_with_others.append("inne - " + str(others_sum) + "zł")

    values = subcats_values_with_others
    labels = subcats_labels_with_others
    title = month_label + " - Podział wydatków spożywczych\n\nCałkowita" \
        " kwota: " + str(my_worksheet.cats_sums["Jedzenie"]) + " zł\n"
    fig_name = results_dir + "/plots/plot{}.png".format(plot_nr)
    plot_nr += 1

    fig = plotPie(values, labels, title)
    plt.savefig(figure=fig, fname=fig_name)
    # endregion

    # -- Piechart of "Rozrywka" items
    # region
    amounts = my_worksheet.spends_values["Rozrywka"]
    subcats = my_worksheet.spends_items["Rozrywka"]
    subcats_dict = {}

    for i, subcat in enumerate(subcats):
        if subcat in list(subcats_dict.keys()):
            subcats_dict[subcat] += amounts[i]
        else:
            subcats_dict[subcat] = amounts[i]

    subcats_values_1 = list(subcats_dict.values())
    subcats_labels_1 = [list(subcats_dict.keys())[i] + " - "
                        + str(round(list(subcats_dict.values())[i], 2)) + "zł"
                        for i in range(len(list(subcats_dict.keys())))]

    values = subcats_values_1
    labels = subcats_labels_1
    title = month_label + " - Podział wydatków kategorii Rozrywka\n\n" \
        + "Całkowita kwota: " \
        + str(my_worksheet.cats_sums["Rozrywka"]) + " zł\n"
    fig_name = results_dir + "/plots/plot{}.png".format(plot_nr)
    plot_nr += 1

    fig = plotPie(values, labels, title)
    plt.savefig(figure=fig, fname=fig_name)
    # endregion

    # -- Piechart of "Rzeczy i sprzęty" items
    # region
    amounts = my_worksheet.spends_values["Rzeczy i sprzęty"]
    subcats = my_worksheet.spends_items["Rzeczy i sprzęty"]
    subcats_dict = {}

    for i, subcat in enumerate(subcats):
        if subcat in list(subcats_dict.keys()):
            subcats_dict[subcat] += amounts[i]
        else:
            subcats_dict[subcat] = amounts[i]

    subcats_values_2 = list(subcats_dict.values())
    subcats_labels_2 = [list(subcats_dict.keys())[i] + " - "
                        + str(round(list(subcats_dict.values())[i], 2)) + "zł"
                        for i in range(len(list(subcats_dict.keys())))]

    values = subcats_values_2
    labels = subcats_labels_2
    title = month_label + " - Podział wydatków kategorii Rzeczy i " \
        "sprzęty\n\nCałkowita kwota: " \
        + str(my_worksheet.cats_sums["Rzeczy i sprzęty"]) + " zł\n"
    fig_name = results_dir + "/plots/plot{}.png".format(plot_nr)
    plot_nr += 1

    fig = plotPie(values, labels, title)
    plt.savefig(figure=fig, fname=fig_name)
    # endregion

    plt.close("all")
