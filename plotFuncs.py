import matplotlib.pyplot as plt
import seaborn as sns

# --------------
# --------------
def plotPie(values, labels, plot_title):

    plt.style.use('default')
    fig = plt.figure(figsize = (12,12))
    plt.pie(x = values, labels = labels, autopct='%1.0f%%',
            textprops = {'fontsize': 20}, labeldistance = 1.15)

    plt.title(plot_title, fontsize = 24, pad = 10)
    plt.subplots_adjust(left = 0.15, bottom = 0.01, right = 0.80, top = 0.95)

    return fig

# --------------
# --------------
def plotBar(values, labels, plot_title):

    plt.style.use('default')
    fig = plt.figure(figsize = (12,12), facecolor = 'white')
    kwargs = {'fillstyle': 'top'}

    sns.barplot(x = values, y = labels, zorder = 2)

    plt.tick_params(axis = 'both', which = 'major', labelsize = 20)
    plt.subplots_adjust(left = 0.26, bottom = 0.1, right = 0.9, top = 0.85)
    plt.title(plot_title, fontsize = 24, pad = 10)
    plt.xlabel('Kwota wydana [zł]', fontsize = 20)
    plt.grid(zorder = 0, axis ='x')

    return fig

# --------------
# --------------
def plotLine(values, plot_title):

    plt.style.use('default')

    fig, ax = plt.subplots(figsize = (9,9))
    fig = sns.lineplot(data = values, markers = True)

    plt.title(plot_title, fontsize = 18, pad = 50)

    plt.ylim([0, 1.1*max(values[max(values)])])
    plt.xticks(list(values.index), [str(i) for i in list(values.index)])
    ax.tick_params(labelsize = 16)
    plt.legend(loc = 0)

    return fig

# --------------
# --------------
def plotHist(values, plot_title):

    plt.style.use('default')

    fig, ax = plt.subplots(figsize = (9, 9))
    bins = np.linspace(min(values), max(values), 6)

    fig = sns.distplot(values, kde = False, bins = bins)
    plt.title(plot_title, fontsize = 22, pad = 50)

    ax.tick_params(labelsize = 16)

    return fig


#
