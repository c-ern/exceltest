import pandas as pd

fsk = pd.ExcelFile("Fehlerkennkarte extern NEU.xls")
print(fsk.sheet_names)

df = {sheet_name: fsk.parse(sheet_name) for sheet_name in fsk.sheet_names}

print(df.keys())

for sheets in df:
    # print(df[sheets].head())
    df[sheets] = df[sheets].dropna(subset=["Unnamed: 0"])
    df[sheets].columns = df[sheets].iloc[0]
    df[sheets] = df[sheets].drop(df[sheets].index[0])
    df[sheets]["Datum:"] = pd.to_datetime(df[sheets]["Datum:"], format='%Y-%m-%d %H:%M:%S', errors='coerce')

print("")
print("")
print("Und jetzt nur das 83002")
print(df['83002'].head())
print(df['83002']["Datum:"])
print(df['38816']["Datum:"])
print(df['38816'].loc["2017-01"])

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
