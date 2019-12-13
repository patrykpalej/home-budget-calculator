import numpy as np
import matplotlib.pyplot as plt

from functions.plotFuncs import plotPie, plotBar, plotLine, plotStack


def create_all_plots(myWorkbook, part_label, results_dir, start_label,
                     n_of_months):

    # -- Part as a whole
    # a) Barplot for all categories
    index_order = np.flip(np.argsort(myWorkbook.cats_sums_list), axis=0)
    values_desc = [myWorkbook.cats_sums_list[i] for i in index_order
                   if myWorkbook.cats_sums_list[i] > 0]
    labels_desc = [myWorkbook.cats_names[i] for i in index_order
                   if myWorkbook.cats_sums_list[i] > 0]

    values = values_desc
    labels = labels_desc
    title = part_label + " - Kwoty wydane w ciągu całego okresu \nna " \
                         "kolejne kategorie"
    fig_name = results_dir + "/plots/plot1.png"

    fig = plotBar(values, labels, title)
    plt.savefig(figure=fig, fname=fig_name)

    # b) Piechart of spendings for the main categories
    index_order = np.flip(np.argsort(myWorkbook.cats_sums_list), axis=0)
    _top_indices = index_order[0:5]
    _low_indices = index_order[5:len(index_order)]

    top_labels = [myWorkbook.cats_names[i] for i in _top_indices]
    top_values = [myWorkbook.cats_sums_list[i] for i in _top_indices]
    others_values = [myWorkbook.cats_sums_list[i] for i in _low_indices]

    top_plus_others_values = top_values + [sum(others_values)]
    top_plus_others_labels_0 = top_labels + ["Pozostałe"]
    top_plus_others_labels = [top_plus_others_labels_0[i] + " - "
                              + str(
        round(top_plus_others_values[i], 2)) + " zł"
                              for i in range(len(top_plus_others_values))]

    values = top_plus_others_values
    labels = top_plus_others_labels
    title = part_label \
            + " - Struktura wydatków z podziałem \nna kategorie\n\n" \
              "Suma wydatków: " + str(round(myWorkbook.sum_total, 2)) + "zł\n"
    fig_name = results_dir + "/plots/plot2.png"

    fig = plotPie(values, labels, title)
    plt.savefig(figure=fig, fname=fig_name)

    # c) Piechart of the metacategories
    metacats_values = [myWorkbook.sum_basic, myWorkbook.sum_addit,
                       myWorkbook.sum_giftdon]
    metacats_labels = ["Podstawowe - " + str(round(myWorkbook.sum_basic, 2))
                       + "zł",
                       "Dodatkowe - " + str(round(myWorkbook.sum_addit, 2))
                       + "zł", "Prezenty i donacje - "
                       + str(round(myWorkbook.sum_giftdon, 2)) + "zł"]

    values = metacats_values
    labels = metacats_labels
    title = part_label + " - Podział wydatków na:\n" \
            + "Podstawowe, Dodatkowe i Prezenty/Donacje\n\nSuma wydatków: " \
            + str(round(myWorkbook.sum_total, 2)) + "zł\n"
    fig_name = results_dir + "/plots/plot3.png"

    fig = plotPie(values, labels, title)
    plt.savefig(figure=fig, fname=fig_name)

    # d) Piechart of incomes
    _values_list_inc = list(myWorkbook.incomes_dict.values())
    _labels_list_inc = list(myWorkbook.incomes_dict.keys())
    incomes_values = [inc for inc in _values_list_inc if inc > 0]
    incomes_labels = [inc + " - " + str(_values_list_inc[i]) + "zł" for i, inc
                      in
                      enumerate(_labels_list_inc) if _values_list_inc[i] > 0]

    values = incomes_values
    labels = incomes_labels
    title = part_label + " - Podział przychodów na \nposzczególne źródła\n\n" \
                         "Suma przychodów: " + str(
        round(myWorkbook.incomes, 2)) + "zł\n" \
                                        "Nadwyżka przychodów: " + str(
        round(myWorkbook.balance[0], 2)) \
            + "zł  (" + str(round(100 * myWorkbook.balance[0]
                                  / myWorkbook.incomes, 2)) \
            + "%)"

    fig_name = results_dir + "/plots/plot4.png"

    fig = plotPie(values, labels, title)
    plt.savefig(figure=fig, fname=fig_name)

    # e) Piechart of earnings
    _values_list_ear = list(myWorkbook.earnings_dict.values())
    _labels_list_ear = list(myWorkbook.earnings_dict.keys())
    earnings_values = [ear for ear in _values_list_ear if ear > 0]
    earnings_labels = [ear + " - " + str(_values_list_ear[i]) + "zł" for i, ear
                       in
                       enumerate(_labels_list_ear) if _values_list_ear[i] > 0]

    values = earnings_values
    labels = earnings_labels
    title = part_label + " - Podział zarobków na \nposzczególne źrodła\n\n" \
            + "Suma zarobków: " + str(myWorkbook.earnings) + "zł\n" \
            + "Nadwyżka zarobków: " + str(round(myWorkbook.balance[1], 2)) \
            + "zł  (" + str(round(100 * myWorkbook.balance[1]
                                  / myWorkbook.earnings, 2)) \
            + "%)"
    fig_name = results_dir + "/plots/plot5.png"

    fig = plotPie(values, labels, title)
    plt.savefig(figure=fig, fname=fig_name)

    # f) Piechart of food subcategories
    amounts = myWorkbook.spends_values_yr["Jedzenie"]
    subcats = myWorkbook.spends_items_yr["Jedzenie"]
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
    # ----
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
    title = part_label + " - Podział wydatków spożywczych\n\nCałkowita kwota: " \
            + str(round(myWorkbook.cats_sums["Jedzenie"], 2)) + " zł\n"
    fig_name = results_dir + "/plots/plot6.png"

    fig = plotPie(values, labels, title)
    plt.savefig(figure=fig, fname=fig_name)

    # -- Averaged month --
    # g) Piechart of spendings for the main categories
    top_plus_others_values_avg = [i / n_of_months for i in
                                  top_plus_others_values]
    top_plus_others_labels_avg = [top_plus_others_labels_0[i] + " - "
                                  + str(round(top_plus_others_values[i]
                                              / n_of_months, 2)) + " zł"
                                  for i in range(len(top_plus_others_values))]

    values = top_plus_others_values_avg
    labels = top_plus_others_labels_avg
    title = part_label \
            + " - Struktura wydatków w uśrednionym \nmiesiącu z podziałem na " \
              "kategorie\n\n" + "Suma wydatków: " \
            + str(round(myWorkbook.sum_total / n_of_months, 2)) + "zł\n"
    fig_name = results_dir + "/plots/plot7.png"

    fig = plotPie(values, labels, title)
    plt.savefig(figure=fig, fname=fig_name)

    # h) Piechart of the metacategories
    metacats_values_avg = [i / n_of_months for i in metacats_values]
    metacats_labels_avg = ["Podstawowe - "
                           + str(round(myWorkbook.sum_basic / n_of_months, 2))
                           + "zł", "Dodatkowe - "
                           + str(round(myWorkbook.sum_addit / n_of_months, 2))
                           + "zł", "Prezenty i donacje - "
                           + str(
        round(myWorkbook.sum_giftdon / n_of_months, 2))
                           + "zł"]

    values = metacats_values_avg
    labels = metacats_labels_avg
    title = part_label + " - Podział wydatków na: " \
            + "Podstawowe, \nDodatkowe i Prezenty/Donacje " \
              "w uśrednionym miesiącu\n\n" \
              "Suma wydatków: " + str(
        round(myWorkbook.sum_total / n_of_months, 2)) \
            + "zł\n"
    fig_name = results_dir + "/plots/plot8.png"

    fig = plotPie(values, labels, title)
    plt.savefig(figure=fig, fname=fig_name)

    # i) Piechart of incomes
    incomes_values_avg = [i / n_of_months for i in incomes_values]
    incomes_labels_avg = [inc + " - " + str(round(_values_list_inc[i]
                                                  / n_of_months, 2)) + "zł"
                          for i, inc in enumerate(_labels_list_inc)
                          if _values_list_inc[i] > 0]

    values = incomes_values_avg
    labels = incomes_labels_avg
    title = part_label + " - Podział przychodów na poszczególne \nźródła " \
                         "w uśrednionym miesiącu\n\n" + "Suma przychodów: " \
            + str(round(myWorkbook.incomes / n_of_months, 2)) + "zł\n" \
            + "Nadwyżka przychodów: " \
            + str(round(myWorkbook.balance[0] / n_of_months, 2)) \
            + "zł  (" + str(round(100 * myWorkbook.balance[0] / n_of_months
                                  / (myWorkbook.incomes / n_of_months),
                                  2)) + "%)"

    fig_name = results_dir + "/plots/plot9.png"

    fig = plotPie(values, labels, title)
    plt.savefig(figure=fig, fname=fig_name)

    # j) Piechart of earnings
    earnings_values_avg = [i / n_of_months for i in earnings_values]
    earnings_labels_avg = [ear + " - "
                           + str(round(_values_list_ear[i] / n_of_months, 2))
                           + "zł"
                           for i, ear in enumerate(_labels_list_ear)
                           if _values_list_ear[i] > 0]

    values = earnings_values_avg
    labels = earnings_labels_avg
    title = part_label + " - Podział zarobków na poszczególne \nźródła " \
                         "w uśrednionym miesiącu\n\nSuma zarobków: " \
            + str(round(myWorkbook.earnings / n_of_months, 2)) + "zł\n" \
                                                                 "Nadwyżka zarobków: " \
            + str(round(myWorkbook.balance[1] / n_of_months, 2)) \
            + "zł  (" + str(round(100 * myWorkbook.balance[1] / n_of_months
                                  / (myWorkbook.earnings / n_of_months), 2)) \
            + "%)"

    fig_name = results_dir + "/plots/plot10.png"

    fig = plotPie(values, labels, title)
    plt.savefig(figure=fig, fname=fig_name)

    # -- Part as a sequence of months --
    # k) Stackplot of cummulated spendings for the top categories
    top_spends_seqs = []
    # top categories
    for c, cat in enumerate(top_labels):
        top_spends_seqs.append([])
        for month in myWorkbook.sheets_list:
            top_spends_seqs[c].append(month.cats_sums[cat])
    # others
    top_spends_seqs.append([])
    for m, month in enumerate(myWorkbook.sheets_list):
        top_spends_seqs[-1].append(np.sum(month.cats_sums_list)
                                   - sum(cat[m] for cat in top_spends_seqs[:-1]
                                         if cat[m]))
    # cummulated sums of the lists
    top_spends_seqs_cumsum = []
    for seq in top_spends_seqs:
        top_spends_seqs_cumsum.append(np.cumsum(seq))

    values = top_spends_seqs_cumsum
    labels = top_labels + ["Pozostałe"]
    title = part_label + " - Skumulowane wartości wydatków na\n poszczególne " \
                         "kategorie"
    fig_name = results_dir + "/plots/plot11.png"

    fig = plotStack(values, labels, title, start_label)
    plt.savefig(figure=fig, fname=fig_name)

    # l) Lineplot of cummulated spendings, incomes and savings
    line_incomes = []
    line_spendings = []
    line_balance = []
    line_savings = []

    for month in myWorkbook.sheets_list:
        line_incomes.append(month.incomes)
        line_spendings.append(sum(month.cats_sums_list))
        line_balance.append(month.balance[0])
        line_savings.append(month.savings)

    line_incomes = np.cumsum(line_incomes)
    line_spendings = np.cumsum(line_spendings)
    line_balance = np.cumsum(line_balance)
    line_savings = np.cumsum(line_savings)

    values = [line_incomes, line_spendings, line_balance, line_savings]
    labels = ["Przychody", "Wydatki", "Nadwyżka\nprzychodów",
              "Oszczędności\ndługoterminowe"]
    title = part_label + " - Skumulowane wartości przychodów, wydatków \n" \
                         "i oszczędności"
    fig_name = results_dir + "/plots/plot12.png"

    fig = plotLine(values, labels, title, start_label)
    plt.savefig(figure=fig, fname=fig_name)

    # m) Lineplot of spendings and incomes in subsequent months
    spendings_list = []
    incomes_list = []
    for month in myWorkbook.sheets_list:
        spendings_list.append(month.sum_total)
        incomes_list.append(month.incomes)

    values = [incomes_list, spendings_list]
    labels = ["Przychody", "Wydatki"]
    title = part_label + " - Przychody i wydatki w kolejnych miesiącach"
    fig_name = results_dir + "/plots/plot13.png"

    fig = plotLine(values, labels, title, start_label)
    plt.savefig(figure=fig, fname=fig_name)

    # n) Lineplot of average spendings for subsequent categories so far
    current_means_seqs = []
    for c, cat in enumerate(top_spends_seqs):
        current_means_seqs.append([])
        for m, month in enumerate(cat):
            current_means_seqs[-1].append(np.mean(top_spends_seqs[c][:m + 1]))

    values = current_means_seqs
    labels = top_labels + ["Pozostałe"]
    title = part_label \
            + " - Dotychczasowe średnie miesięczne wydatki na \nposzczególne " \
              "kategorie"
    fig_name = results_dir + "/plots/plot14.png"

    fig = plotLine(values, labels, title, start_label)
    plt.savefig(figure=fig, fname=fig_name)

    # o) Lineplot of main sources
    #  choosing of the relevant sources
    incomes_main = []
    for inc in myWorkbook.incomes_dict:
        if myWorkbook.incomes_dict[inc] > 0.02 * myWorkbook.incomes:
            incomes_main.append(inc)

    # preprocessing
    incomes_seqs = [[] for i in range(len(incomes_main))]
    for p, inc in enumerate(incomes_main):
        for m, month in enumerate(myWorkbook.sheets_list):
            try:
                incomes_seqs[p].append(month.incomes_dict[inc])
            except:
                incomes_seqs[p].append(0)

    values = incomes_seqs
    labels = incomes_main
    title = part_label + " - Kwoty przychodów z najważniejszych \n źrodeł"
    fig_name = results_dir + "/plots/plot15.png"

    fig = plotLine(values, labels, title, start_label)
    plt.savefig(figure=fig, fname=fig_name)

    plt.close("all")

    return spendings_list, incomes_list
