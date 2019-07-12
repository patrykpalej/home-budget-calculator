import matplotlib.pyplot as plt
import seaborn as sns


def plotPie(values, labels, plot_title):

    plt.style.use('default')
    fig = plt.figure(figsize=(12, 12))
    plt.pie(x=values, labels=labels, autopct='%1.0f%%',
            textprops={'fontsize': 20}, labeldistance=1.1, startangle = 140)

    plt.title(plot_title, fontsize=24, pad=21)
    plt.subplots_adjust(left=0.0, bottom=0.00, right=1.0, top=0.8)
    
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


def plotLine(values, labels, plot_title):

    plt.style.use('default')

    fig, ax = plt.subplots(figsize=(14, 9))
    for i, vec in enumerate(values):
        fig = plt.plot(vec, label=labels[i], marker='.', markersize=14)

    plt.title(plot_title, fontsize=22, pad=15)
    plt.grid()

    box = ax.get_position()
    ax.set_position([box.x0, box.y0, box.width * 0.8, box.height])
    ax.legend(loc='center right', bbox_to_anchor=(1.3, 0.5), fontsize='x-large')
    plt.xlabel('Miesiące', fontdict={'size': 18})
    plt.ylabel('Kwota [zł]', fontdict={'size': 18})
    plt.xticks(range(len(values[0])),
               [str(i+1) for i in range(len(values[0]))])
    ax.tick_params(labelsize=16)

    return fig


def plotStack(values, labels, plot_title):

    plt.style.use('default')

    fig, ax = plt.subplots(figsize=(14, 9))
    fig = plt.stackplot(range(len(values[0])), values, labels=labels)
    plt.title(plot_title, fontsize=22, pad=15)
    ax.legend(loc=2, fontsize='x-large')
    plt.xlabel('Miesiące', fontdict={'size': 18})
    plt.ylabel('Wydatki[zł]', fontdict={'size': 18})

    plt.xticks(range(len(values[0])), [str(i+1) for i in range(len(values[0]))])
    ax.tick_params(labelsize=16)

    return fig
