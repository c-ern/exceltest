import pandas as pd
import datetime
import matplotlib.pyplot as plt
import locale
import numpy as np
from matplotlib.ticker import FuncFormatter

locale.setlocale(locale.LC_ALL, 'German')

jahr = 2017
kalenderwoche = 26

# d = "2017-W24"
d = str(jahr)+'-W'+str(kalenderwoche)
r = datetime.datetime.strptime(d + '-0', "%Y-W%W-%w")
print(r)

startdatum = datetime.datetime(jahr, 1, 1, 0, 0)
enddatum = datetime.datetime(jahr, 12, 31, 0, 0)
print(startdatum)
print(enddatum)
# df['71231'].loc[r]
# https://stackoverflow.com/questions/17087314/get-date-from-week-number

fsk = pd.ExcelFile("Fehlerkennkarte extern NEU.xls")
print(fsk.sheet_names)

df = {sheet_name: fsk.parse(sheet_name) for sheet_name in fsk.sheet_names}

print(df.keys())

summe_io = 'Summe i.O. [Stück]'
# summe_nio = df[sheets].columns[19]
summe_nio = 'Summe\nn.i.O.\n[Stück]'
print("Summe NIO heißt:" + repr(summe_nio))
f_besch_lagerpad = 'Beschädigungam Lagerpad'
f_aufplatt_lagerpad = 'Aufplattierung Lagerpad'
f_besch_aussenkontur = 'Beschädigung Außenkontur'
f_besch_innenkontur = 'Beschädigung Innenkontur'
f_besch_oelablauf = 'Beschädigung Ölablauf'
f_besch_dichtflaeche = 'Beschädigung Dichtfläche'
f_besch_querbohrung = 'Beschädigung Querbohrung'
f_besch_oeltasche = 'Beschädigung Öltasche'
f_oelige_teile = 'Ölige Teile'
f_rueckstaende = 'Rückstände'
f_sonstiges_1 = 'Sonstiges 1'
f_sonstiges_2 = 'Sonstiges 2'
f_sonstiges_3 = 'Sonstiges 3'
f_sonstiges_4 = 'Sonstiges 4'
f_sonstiges_5 = 'Sonstiges 5'

spalten = [summe_io,
           summe_nio,
           f_besch_lagerpad,
           f_aufplatt_lagerpad,
           f_besch_aussenkontur,
           f_besch_innenkontur,
           f_besch_oelablauf,
           f_besch_dichtflaeche,
           f_besch_querbohrung,
           f_besch_oeltasche,
           f_oelige_teile,
           f_rueckstaende,
           f_sonstiges_1,
           f_sonstiges_2,
           f_sonstiges_3,
           f_sonstiges_4,
           f_sonstiges_5]

farbe_io = 'green'
farbe_nio = 'red'
farbe_fehler = ['orange' for i in spalten]
colors = [farbe_io, farbe_nio] + farbe_fehler

for sheets in df:
    # print(df[sheets].head())
    df[sheets] = df[sheets].dropna(subset=['Unnamed: 0'])
    df[sheets].columns = df[sheets].iloc[0]
    df[sheets] = df[sheets].drop(df[sheets].index[0])
    df[sheets]['Datum:'] = pd.to_datetime(df[sheets]['Datum:'], format='%Y-%m-%d %H:%M:%S', errors='coerce')
    df[sheets] = df[sheets].set_index('Datum:').groupby(pd.TimeGrouper('W')).sum()
    df[sheets].columns.name = None

    print("Alle Sheets")
    print(sheets)
    print('------------------')
    # print(df[sheets][spalten].loc[datum])
    # print(df[sheets][spalten])
    if r in df[sheets].index:
        print("Folgendes hat "+str(r)+" im Index:")
        print(sheets)
        print(df[sheets][summe_io].loc[r].sum() > 0)
        print(df[sheets][summe_io].loc[r].sum())
        # print(df[sheets])
        # if df[sheets][summe_io].loc[r].sum() > 0:
        print("Und die Summe ist auch größer Null")
        print(sheets)
        print(df[sheets].index)
        print("")
        print(df[sheets][spalten].loc[r])
        print("")

        summe_io_jahr = df[sheets][summe_io].loc[startdatum:enddatum].mean()
        summe_nio_jahr = df[sheets][summe_nio].loc[startdatum:enddatum].mean()
        anteil_nio_jahr = summe_nio_jahr / (summe_io_jahr + summe_nio_jahr)
        summe_io_kw = df[sheets][summe_io].loc[r].mean()
        summe_nio_kw = df[sheets][summe_nio].loc[r].mean()
        anteil_nio_kw = summe_nio_kw / (summe_io_kw + summe_nio_kw)
        print(anteil_nio_jahr)

        ax = df[sheets][spalten].loc[r].plot(kind='bar', rot=90, color=colors)
        for p in ax.patches:
            # print(p.get_height().type)
            prozent = p.get_height()/(summe_io_kw + summe_nio_kw)
            # prozent = 123
            ax.annotate(locale.format('%.0f', np.round(p.get_height(), decimals=0), True),
                        (p.get_x()+p.get_width()/2.,
                        p.get_height()),
                        ha='center',
                        va='center',
                        xytext=(0, 15),
                        textcoords='offset points',
                        color='black',
                        fontsize='small',
                        weight='heavy')
            ax.annotate(str('{:.1%}'.format(prozent)),
                        (p.get_x()+p.get_width()/2.,
                         p.get_height()),
                        ha='center',
                        va='center',
                        xytext=(0, 6),
                        textcoords='offset points',
                        color='black',
                        fontsize='x-small',
                        weight='normal')
        print(ax.patches)
        plt.title('Technomix: Sichtprüfung t' + str(sheets) + ' in KW ' + str(kalenderwoche) + ' / ' + str(jahr))
        ax.yaxis.set_major_formatter(FuncFormatter(lambda y, _: locale.format('%.0f', y, True)))
        ax.yaxis.grid(True, linestyle='dotted', linewidth=0.3, color='black')
        plt.axvline(x=1.5, color='red')
        # plt.text(5, 1000, 'Der Mittelwert ist ' + str(mittelwert))
        ax.annotate('Anteil NIO in akt. KW:     ' + str('{:.1%}'.format(anteil_nio_kw)), xy=(1, 1),  xycoords='data', xytext=(0.6, -0.8), textcoords='axes fraction')
        ax.annotate('Anteil NIO im Jahr ' + str(jahr) + ': ' + str('{:.1%}'.format(anteil_nio_jahr)), xy=(1, 1),  xycoords='data', xytext=(0.6, -0.9), textcoords='axes fraction')
        plt.tight_layout()

        fig = ax.get_figure()
        # fig.savefig(str(sheets)+"--"+str(jahr)+'-KW' + str(kalenderwoche)+".pdf")
        fig.savefig(str(jahr) + '-KW' + str(kalenderwoche) + '__t' + str(sheets) + '.pdf')
        plt.close()


print("----ENDE----")
