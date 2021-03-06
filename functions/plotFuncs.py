import numpy as np
import seaborn as sns
from math import floor
import matplotlib.pyplot as plt
from datetime import date
from math import ceil
from dateutil.relativedelta import *
from matplotlib.ticker import AutoMinorLocator, MultipleLocator, MaxNLocator, AutoLocator


def plotPie(values, labels, plot_title):

    plt.style.use('default')
    fig = plt.figure(figsize=(13, 13*0.83))
    plt.pie(x=values, autopct='%1.0f%%', textprops={'fontsize': 20},
            startangle=0)

    plt.title(plot_title, fontsize=24, pad=1)
    plt.subplots_adjust(left=-0.02, bottom=0.0, right=0.78, top=0.78)
    plt.legend(labels=labels, loc=(0.825, 0.75), fontsize='xx-large')

    return fig


def plotBar(values, labels, plot_title):

    plt.style.use('default')
    fig, ax = plt.subplots(figsize=(12, 12*0.83), facecolor='white')

    sns.barplot(x=values, y=labels, zorder=2)
    ax.xaxis.set_minor_locator(AutoLocator())

    plt.tick_params(axis='both', which='major', labelsize=14)
    plt.subplots_adjust(left=0.3, bottom=0.1, right=0.98, top=0.85)
    plt.title(plot_title, fontsize=24, pad=10, loc='left')
    plt.xlabel('Kwota wydana [zł]', fontsize=20)
    plt.grid(zorder=0, axis='x', which='both')

    return fig


def plotLine(values, labels, plot_title, start_label):

    plt.style.use('default')

    fig, ax = plt.subplots(figsize=(14, 9))
    for i, vec in enumerate(values):
        fig = plt.plot(vec, label=labels[i], marker='.', markersize=14)

    plt.title(plot_title, fontsize=22, pad=15)
    plt.grid()
    plt.ylim([-2, 1.1*max([max(v) for v in values])])

    box = ax.get_position()
    ax.set_position([box.x0, box.y0, box.width * 0.8, box.height])
    ax.legend(loc='center right', bbox_to_anchor=(1.3, 0.5),
              fontsize='x-large')
    plt.ylabel('Kwota [zł]', fontdict={'size': 18})

    # -- xtick labels
    x_tick_labels = []
    month = start_label[0]
    year = start_label[1]
    for i in range(1+floor(len(values[0])/2)):
        # to prevent xticklabels exceeding the actual time period
        if 2*i+1 > len(values[0]):
            break

        x_tick_labels.append(str(month)+'.20'+str(year))
        month += 2

        if month > 12:
            month -= 12
            year += 1
    xticks_idx = [2 * i for i in range(1 + round(len(values[0]) / 2))]

    if len(values[0]) > 20:
        start_date = date(2000 + start_label[1], start_label[0], 1)
        dates_range = [start_date + relativedelta(months=i)
                       for i in range(len(values[0]))]
        xticks_values = dates_range[0::ceil(len(values[0])/10)]
        xticks_idx = list(range(len(values[0])))[0::ceil(len(values[0])/10)]
        x_tick_labels = [str(date(x.year, x.month, 1).strftime("%m.%Y"))
                         for x in xticks_values]

    plt.xticks(xticks_idx, x_tick_labels, rotation=15)
    ax.tick_params(labelsize=16)

    return fig


def twoAxisLinePlot(values, labels, plot_title, start_label):
    plt.style.use('default')

    fig, ax1 = plt.subplots(figsize=(14, 9))
    ax2 = ax1.twinx()
    plt.grid(axis='both')
    # - - -
    ax1.plot(values[0], label=labels[0], marker='.', markersize=14,
             color='#1f77b4')
    trend = np.polyfit([i for i in range(len(values[0]))], values[0], 1)
    p = np.poly1d(trend)
    ax1.plot([0, len(values[0])], [p[0], p[0]+p[1]*len(values[0])],
             color='#1f77b4', linestyle='dashed', label="Linia trendu")
    ax1.tick_params(axis='y', labelcolor='#1f77b4', labelsize=16)
    # --
    ax2.plot(values[1], label=labels[1], marker='.', markersize=14,
             color='orange')
    relevant_values = [val for val in values[1] if -1 < val < 1]
    trend = np.polyfit([i for i in range(len(relevant_values))],
                       relevant_values, 1)
    p = np.poly1d(trend)
    ax2.plot([0, len(values[1])], [p[0], p[0] + p[1] * len(values[1])],
             color='orange', linestyle='dashed', label="Linia trendu")
    ax2.tick_params(axis='y', labelcolor='orange', labelsize=16)

    # - - -
    plt.title(plot_title, fontsize=22, pad=15)

    box = ax2.get_position()
    ax2.set_position([box.x0, box.y0, box.width * 0.8, box.height])
    ax1.legend(loc='center right', bbox_to_anchor=(1.34, 0.58),
               fontsize='x-large')
    ax2.legend(loc='center right', bbox_to_anchor=(1.33, 0.38),
               fontsize='x-large')
    ax2.set_ylim([-1, 1])

    # -- xtick labels
    x_tick_labels = []
    month = start_label[0]
    year = start_label[1]
    for i in range(1 + floor(len(values[0]) / 2)):
        # to prevent xticklabels exceeding the actual time period
        if 2 * i + 1 > len(values[0]):
            break

        x_tick_labels.append(str(month) + '.20' + str(year))
        month += 2

        if month > 12:
            month -= 12
            year += 1
    xticks_idx = [2 * i for i in range(1 + round(len(values[0]) / 2))]

    if len(values[0]) > 20:
        start_date = date(2000 + start_label[1], start_label[0], 1)
        dates_range = [start_date + relativedelta(months=i)
                       for i in range(len(values[0]))]
        xticks_values = dates_range[0::ceil(len(values[0])/10)]
        xticks_idx = list(range(len(values[0])))[0::ceil(len(values[0])/10)]
        x_tick_labels = [str(date(x.year, x.month, 1).strftime("%m.%Y"))
                         for x in xticks_values]

    plt.xticks(xticks_idx, x_tick_labels)
    ax1.tick_params(axis='x', labelsize=16, rotation=15)

    return fig


def plotStack(values, labels, plot_title, start_label):

    plt.style.use('default')

    fig, ax = plt.subplots(figsize=(14, 9))
    fig = plt.stackplot(range(len(values[0])), values, labels=labels)
    plt.title(plot_title, fontsize=22, pad=15)
    plt.grid()
    ax.set_axisbelow(True)
    ax.legend(loc=2, fontsize='x-large')
    plt.ylabel('Kwota[zł]', fontdict={'size': 18})

    x_tick_labels = []
    month = start_label[0]
    year = start_label[1]
    for i in range(1 + round(len(values[0]) / 2)):
        # to prevent xticklabels exceeding the actual time period
        if 2*i+1 > len(values[0]):
            break

        x_tick_labels.append(str(month) + '.20' + str(year))
        month += 2

        if month > 12:
            month -= 12
            year += 1

    xticks_idx = [2 * i for i in range(1 + round(len(values[0]) / 2))]

    if len(values[0]) > 20:
        start_date = date(2000 + start_label[1], start_label[0], 1)
        dates_range = [start_date + relativedelta(months=i)
                       for i in range(len(values[0]))]
        xticks_values = dates_range[0::ceil(len(values[0]) / 10)]
        xticks_idx = list(range(len(values[0])))[0::ceil(len(values[0]) / 10)]
        x_tick_labels = [str(date(x.year, x.month, 1).strftime("%m.%Y"))
                         for x in xticks_values]

    plt.xticks(xticks_idx, x_tick_labels)
    ax.tick_params(labelsize=16, rotation=15)

    return fig


def plotScatter(values, plot_title):

    plt.style.use('default')

    fig, ax = plt.subplots(figsize=(14, 9))
    fig = plt.scatter(values[0], values[1], marker='.', s=100,
                      label='Wartości zarejestrowane')
    trend = np.polyfit(values[0], values[1], 1)
    p = np.poly1d(trend)
    x1 = 0.95 * min(min(values[0]), min(values[1]))
    x2 = 1.05 * max(max(values[0]), max(values[1]))
    plt.plot([x1, x2], [p[1]*x1+p[0], p[1]*x2+p[0]], 'k--',
             label='Linia trendu')

    fig = plt.plot([x1, x2],
                   [0.95 * min(min(values[0]), min(values[1])),
                   1.05 * max(max(values[0]), max(values[1]))], color='red',
                   linewidth=2, label='y=x')
    plt.title(plot_title, fontsize=22, pad=15)
    ax.axis('equal')
    plt.xlabel('Przychody [zł]', fontdict={'size': 18})
    plt.ylabel('Wydatki [zł]', fontdict={'size': 18})
    ax.tick_params(labelsize=16)
    plt.grid()
    plt.legend()

    return fig


def twoAxisLinePlot1(values, labels, plot_title, start_label, n_rolling):
    plt.style.use('default')

    fig, ax1 = plt.subplots(figsize=(14, 9))
    ax2 = ax1.twinx()
    plt.grid(axis='both')
    # - - -
    ax1.plot(values[0], label=labels[0], marker='.', markersize=14,
             color='#1f77b4')
    trend = np.polyfit([i for i in range(len(values[0]))], values[0], 1)
    p = np.poly1d(trend)
    ax1.plot([0, len(values[0])], [p[0], p[0]+p[1]*len(values[0])],
             color='#1f77b4', linestyle='dashed', label="Linia trendu")
    ax1.tick_params(axis='y', labelcolor='#1f77b4', labelsize=16)
    # --
    ax2.plot(np.arange(n_rolling-1, len(values[1])+n_rolling-1), values[1],
             label=labels[1], marker='.', markersize=14, color='orange')
    trend = np.polyfit([i for i in range(len(values[1]))], values[1], 1)
    p = np.poly1d(trend)
    ax2.plot([n_rolling-1, len(values[1])+n_rolling-1],
             [p[0], p[0] + p[1] * len(values[1])],
             color='orange', linestyle='dashed', label="Linia trendu")
    ax2.tick_params(axis='y', labelcolor='orange', labelsize=16)

    # - - -
    plt.title(plot_title, fontsize=22, pad=15)

    box = ax2.get_position()
    ax2.set_position([box.x0, box.y0, box.width * 0.8, box.height])
    ax1.legend(loc='center right', bbox_to_anchor=(1.34, 0.58),
               fontsize='x-large')
    ax2.legend(loc='center right', bbox_to_anchor=(1.33, 0.38),
               fontsize='x-large')

    x_tick_labels = []
    month = start_label[0]
    year = start_label[1]
    for i in range(1 + floor(len(values[0]) / 2)):
        # to prevent xticklabels exceeding the actual time period
        if 2 * i + 1 > len(values[0]):
            break

        x_tick_labels.append(str(month) + '.20' + str(year))
        month += 2

        if month > 12:
            month -= 12
            year += 1

    xticks_idx = [2 * i for i in range(1 + round(len(values[0]) / 2))]

    if len(values[0]) > 20:
        start_date = date(2000 + start_label[1], start_label[0], 1)
        dates_range = [start_date + relativedelta(months=i)
                       for i in range(len(values[0]))]
        xticks_values = dates_range[0::ceil(len(values[0]) / 10)]
        xticks_idx = list(range(len(values[0])))[0::ceil(len(values[0]) / 10)]
        x_tick_labels = [str(date(x.year, x.month, 1).strftime("%m.%Y"))
                         for x in xticks_values]

    plt.xticks(xticks_idx, x_tick_labels)
    ax1.tick_params(axis='x', labelsize=16, rotation=15)

    ax1.set_ylim([0, 1.1*max(max(values[0]), max(values[1]))])
    ax2.set_ylim([0, 1.1*max(max(values[0]), max(values[1]))])

    return fig
