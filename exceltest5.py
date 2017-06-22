import pandas as pd
import datetime
import matplotlib.pyplot as plt

d = "2016-W25"
r = datetime.datetime.strptime(d + '-0', "%Y-W%W-%w")
print(r)
# df['71231'].loc[r]
# https://stackoverflow.com/questions/17087314/get-date-from-week-number

fsk = pd.ExcelFile("Fehlerkennkarte extern NEU.xls")
print(fsk.sheet_names)

df = {sheet_name: fsk.parse(sheet_name) for sheet_name in fsk.sheet_names}

print(df.keys())

datum = '2017-01'
datum1 = pd.Timestamp('2017-01-01')
datum2 = pd.Timestamp('2017-01-31')
print(datum1)
print(datum2)

for sheets in df:
    # print(df[sheets].head())
    df[sheets] = df[sheets].dropna(subset=['Unnamed: 0'])
    df[sheets].columns = df[sheets].iloc[0]
    df[sheets] = df[sheets].drop(df[sheets].index[0])
    df[sheets]['Datum:'] = pd.to_datetime(df[sheets]['Datum:'], format='%Y-%m-%d %H:%M:%S', errors='coerce')
    df[sheets] = df[sheets].set_index('Datum:').groupby(pd.TimeGrouper('W')).sum()
    df[sheets].columns.name = None

    summe_io = df[sheets].columns[21]
    summe_nio = df[sheets].columns[19]

    if sheets == '38947':
        summe_io = df[sheets].columns[22]
        summe_nio = df[sheets].columns[20]
    if sheets == '71231':
        summe_io = df[sheets].columns[20]
        summe_nio = df[sheets].columns[18]

    spalten = [summe_io, summe_nio]
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
        fig = df[sheets][spalten].loc[r].plot(kind='bar').get_figure()
        fig.savefig(str(sheets)+"-"+d+".pdf")


print("----ENDE----")

"""
    if datum in df[sheets].index:
        print("Folgendes hat "+datum+" im Index:")
        print(sheets)
        print(df[sheets].index)
        print("")
        print(df[sheets].loc[datum1:datum2])
        print("")
"""

"""
    if not df[sheets].loc[datum1:datum2].empty:
        print("Folgendes hat "+datum+" im Index:")
        print(sheets)
        print(df[sheets][spalten].loc[datum])
        # fig = df[sheets][spalten].loc[datum].plot(kind='bar').get_figure()
        # fig.savefig(str(sheets)+".pdf")
"""

"""
print("")
print("")
print("Und jetzt nur das 83002")
print(df['83002'].head())
# print(df['83002']["Datum:"])
# print(df['38816']["Datum:"])
print(df['38816'].loc["2017-01"])
"""

# df = fsk.parse("83002")
# print(df)

"""
df = df.dropna(subset=["Unnamed: 0"])
# print(df)

df.columns = df.iloc[0]
df.drop(df.index[0])
df["Datum:"] = pd.to_datetime(df["Datum:"], format='%Y-%m-%d %H:%M:%S', errors='coerce')
print("------------------------------------------------")
print("Hier kommt die Spalte Datum")
print("------------------------------------------------")
print(df["Datum:"])

df.set_index('Datum:', inplace=True)
print("------------------------------------------------")
print("Hier kommt neuer Index")
print("------------------------------------------------")
print(df.head())


df = df.drop(df.index[0])
print("------------------------------------------------")
print("------Hier wurde die erste Zeile gelöscht-------")
print("------------------------------------------------")
print(df.head())

print("------------------------------------------------")
print("------Columns Name-------")
print("------------------------------------------------")
print(df.columns.name)

df.columns.name = None
# df.columns.name = df.index.name


print(df.head())

print(df.index.name)

# df.index.name = None
print(df.index.name)
print(df.head())
print(df.iloc[0])
print(df.index[0])
# df = df.drop(df.index[0])
print(df.index.name)
print("Hallo")
"""
