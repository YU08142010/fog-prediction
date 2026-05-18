import pandas as pd
from sklearn.tree import DecisionTreeClassifier

df = pd.read_csv("data.csv")

X = df[["気温(℃)", "風速(m/s)", "相対湿度(％)"]]
y = df["蒲生水道橋"]

model = DecisionTreeClassifier(max_depth=3)
model.fit(X, y)

data = pd.DataFrame([[10.0, 0, 0]], columns=["気温(℃)", "風速(m/s)", "相対湿度(％)"])

prediction = model.predict(data)

if prediction[0] == 1:
    print("AIの予測：霧が発生する可能性が高いです！")
else:
    print("AIの予測：霧は発生しないでしょう。")
