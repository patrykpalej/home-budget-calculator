import matplotlib.pyplot as plt
import seaborn as sns
from math import floor


def plotPie(values, labels, plot_title):

    plt.style.use('default')
    fig = plt.figure(figsize=(12, 12))
    plt.pie(x=values, autopct='%1.0f%%', textprops={'fontsize': 20},
            startangle=0)

    plt.title(plot_title, fontsize=24, pad=1)
    plt.subplots_adjust(left=0.0, bottom=0.0, right=0.78, top=0.78)
    plt.legend(labels=labels, loc=(0.8, 0.8), fontsize='xx-large')

    return fig


def plotBar(values, labels, plot_title):

    plt.style.use('default')
    fig = plt.figure(figsize=(12, 12), facecolor='white')

    sns.barplot(x=values, y=labels, zorder=2)

    plt.tick_params(axis='both', which='major', labelsize=20)
    plt.subplots_adjust(left=0.26, bottom=0.1, right=0.9, top=0.85)
    plt.title(plot_title, fontsize=24, pad=10)
    plt.xlabel('Kwota wydana [zł]', fontsize=20)
    plt.grid(zorder=0, axis='x')

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
    ax.legend(loc='center right', bbox_to_anchor=(1.3, 0.5), fontsize='x-large')
    plt.ylabel('Kwota [zł]', fontdict={'size': 18})

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

    plt.xticks([2*i for i in range(1+round(len(values[0])/2))],
               x_tick_labels, rotation=15)
    ax.tick_params(labelsize=16)

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

    plt.xticks([2 * i for i in range(1 + round(len(values[0]) / 2))],
               x_tick_labels, rotation=15)

    ax.tick_params(labelsize=16)

    return fig


def plotScatter(values, plot_title):

    plt.style.use('default')

    fig, ax = plt.subplots(figsize=(14, 9))
    fig = plt.scatter(values[0], values[1], marker='.', s=100)
    fig = plt.plot([0.95 * min(min(values[0]), min(values[1])),
                   1.05 * max(max(values[0]), max(values[1]))],
                   [0.95 * min(min(values[0]), min(values[1])),
                   1.05 * max(max(values[0]), max(values[1]))], color='red',
                   linewidth=2)
    plt.title(plot_title, fontsize=22, pad=15)
    ax.axis('equal')
    plt.xlabel('Przychody [zł]', fontdict={'size': 18})
    plt.ylabel('Wydatki [zł]', fontdict={'size': 18})
    ax.tick_params(labelsize=16)

    return fig
