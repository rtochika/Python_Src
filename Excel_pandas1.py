import pandas as pd
import numpy as np
#pip pandas だけではエラーになるので、同時にpip install xlrd を実行する！
#
df = pd.read_excel("C:/Users\86001/PycharmProjects/HelloTensorFlow/datas/excel-comp-data.xlsx")
df.head()

#df["total"] = df["Jan"] + df["Feb"] + df["Mar"]

#print(df["Jan"].sum(), df["Jan"].mean(),df["Jan"].min(),df["Jan"].max())
print(pd.crosstab(df['account'], df['Mar']))
#print(pd.crosstab(df['Mar'],df['account'] ))
