from classes import *
from plotFuncs import *
import os, sys
import numpy as np
from pptx import Presentation
from pptx.util import Inches

# 1. Loading the file with data for one year
file_path = os.getcwd() + '/data/yearly/20' + sys.argv[1]+'.xlsx'
#file_path = os.getcwd() + '/data/yearly/2099.xlsx'
year_label = '20' + sys.argv[1]
#year_label = '2099'

myWorkbook = MyWorkbook(file_path)
myWorksheets = myWorkbook.sheets_list

myWorkbook.sumAllSheets(myWorksheets)


# 2. Preparing data from the parsed sheets for visualization
# -- Year as a whole
# a) Barplot for all categories
index_order = np.flip(np.argsort(myWorkbook.cats_sums_list))
values_desc = [myWorkbook.cats_sums_list[i] for i in index_order
               if myWorkbook.cats_sums_list[i] > 0]
labels_desc = [myWorkbook.cats_names[i] for i in index_order
               if myWorkbook.cats_sums_list[i] > 0]

# b) Piechart of spendings for the main categories
index_order = np.flip(np.argsort(myWorkbook.cats_sums_list))
top_indices = index_order[0:5]
low_indices = index_order[5:len(index_order)]

top_labels = [myWorkbook.cats_names[i] for i in top_indices]
top_values = [myWorkbook.cats_sums_list[i] for i in top_indices]
others_values = [myWorkbook.cats_sums_list[i] for i in low_indices]

top_plus_others_values = top_values + [sum(others_values)]
top_plus_others_labels = top_labels + ['Pozostałe']
top_plus_others_labels = [top_plus_others_labels[i] + '\n'
                        + str(round(top_plus_others_values[i],2)) + ' zł'
                        for i in range(len(top_plus_others_values))]

# c) Piechart of the metacategories
metacats_values = [myWorkbook.sum_basic, myWorkbook.sum_addit,
                   myWorkbook.sum_giftdon]
metacats_labels = ['Podstawowe\n'+str(round(myWorkbook.sum_basic, 2))+'zł',
                   'Dodatkowe\n'+str(round(myWorkbook.sum_addit, 2))+'zł',
                   'Prezenty i donacje\n'+
                   str(round(myWorkbook.sum_giftdon,2))+'zł']

# d) Piechart of incomes
_values_list = list(myWorkbook.incomes_dict.values())
_labels_list = list(myWorkbook.incomes_dict.keys())
incomes_values = [inc for inc in _values_list if inc>0]
incomes_labels = [inc+'\n'+str(_values_list[i])+'zł' for i, inc in
                  enumerate(_labels_list) if _values_list[i] > 0]

# e) Piechart of earnings
_values_list = list(myWorkbook.earnings_dict.values())
_labels_list = list(myWorkbook.earnings_dict.keys())
earnings_values = [ear for ear in _values_list if ear>0]
earnings_labels = [ear+'\n'+str(_values_list[i])+'zł' for i, ear in \
                   enumerate(_labels_list) if _values_list[i] > 0]

# f) Piechart of food subcategories
amounts = myWorkbook.spends_values['Jedzenie']
subcats = myWorkbook.spends_items['Jedzenie']
subcats_dict = {}

for i, subcat in enumerate(subcats):
    if subcat in list(subcats_dict.keys()):
        subcats_dict[subcat] += amounts[i]
    else:
        subcats_dict[subcat] = amounts[i]

subcats_values = list(subcats_dict.values())
subcats_labels = [list(subcats_dict.keys())[i]+'\n'
                  + str(round(list(subcats_dict.values())[i],2))+'zł'
                  for i, sc in enumerate(list(subcats_dict.keys()))]


# 3. Visualization and saving the plots
results_dir = os.getcwd() + '/results/' + year_label + ' - wyniki'
if not os.path.exists(results_dir):
    os.mkdir(results_dir)

# a) Barplot for all categories
values = values_desc
labels = labels_desc
title = year_label + ' - Kwoty wydane w ciągu roku \n na kolejne kategorie\n'
fig_name = results_dir + '/plot1.png'

fig = plotBar(values, labels, title)
plt.savefig(figure=fig, fname=fig_name)

# b) Piechart of spendings for the main categories
values = top_plus_others_values
labels = top_plus_others_labels
title = year_label \
        + ' - Struktura rocznych wydatków\n z podziałem na kategorie\n\n' \
        + 'Całkowita kwota: ' + str(round(myWorkbook.sum_total, 2)) + 'zł\n'
fig_name = results_dir + '/plot2.png'

fig = plotPie(values, labels, title)
plt.savefig(figure=fig, fname=fig_name)

# c) Piechart of the metacategories
values = metacats_values
labels = metacats_labels
title = year_label + ' - Podział wydatków na: \n' \
        + 'Podstawowe, Dodatkowe i Prezenty/Donacje\n\n'
fig_name = results_dir + '/plot3.png'

fig = plotPie(values, labels, title)
plt.savefig(figure=fig, fname=fig_name)

# d) Piechart of incomes
values = incomes_values
labels = incomes_labels
title = year_label + ' - Podział przychodów na poszczególne źródła\n\n' \
        + 'Całkowita kwota: ' + str(myWorkbook.incomes) + 'zł\n' \
        + 'Bilans przychodów: ' + str(round(myWorkbook.ballance[0], 2)) + 'zł\n'
fig_name = results_dir + '/plot4.png'

fig = plotPie(values, labels, title)
plt.savefig(figure=fig, fname=fig_name)

# e) Piechart of earnings
values = earnings_values
labels = earnings_labels
title = year_label + ' - Podział zarobków na poszczególne źrodła\n\n' \
        + 'Całkowita kwota: ' + str(myWorkbook.earnings) + 'zł\n' \
        + 'Bilans zarobków: ' + str(round(myWorkbook.ballance[1], 2))+'zł\n'
fig_name = results_dir + '/plot5.png'

fig = plotPie(values, labels, title)
plt.savefig(figure=fig, fname=fig_name)

# f) Piechart of food subcategories
values = subcats_values
labels = subcats_labels
title = year_label + ' - Podział wydatków spożywczych\n\n'
fig_name = results_dir + '/plot6.png'

fig = plotPie(values, labels, title)
plt.savefig(figure=fig, fname=fig_name)

plt.close('all')

# 4. Pptx presentation
prs = Presentation()
title_slide_layout = prs.slide_layouts[0]
blank_slide_layout = prs.slide_layouts[6]
slides = []

slides.append(prs.slides.add_slide(title_slide_layout))
title = slides[-1].shapes.title
title.text = '20' + sys.argv[1] + ' - raport finansowy'

for i in range(6):
    slides.append(prs.slides.add_slide(blank_slide_layout))
    left = Inches(1.2)
    top = Inches(0.0)
    height = width = Inches(7.5)

    pic_path = 'results/20' + sys.argv[1] + ' - wyniki/plot' + \
               str(i+1) + '.png'
    pic = slides[-1].shapes.add_picture(pic_path, left, top, height, width)

prs.save(results_dir + '/20' + sys.argv[1] + ' - raport finansowy.pptx')
