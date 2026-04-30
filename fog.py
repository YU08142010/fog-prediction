import pandas as pd
from sklearn.tree import DecisionTreeClassifier

df = pd.read_csv("fog_data.csv")

X = df[["気温", "湿度", "風速"]]
y = df["霧発生"]

model = DecisionTreeClassifier(max_depth=3)
model.fit(X, y)

print("学習が完了しました！")

new_data = pd.DataFrame([[10.0, 0, 0]], columns=["気温", "湿度", "風速"])

prediction = model.predict(new_data)

if prediction[0] == 1:
    print("AIの予測：霧が発生する可能性が高いです！")
else:
    print("AIの予測：霧は発生しないでしょう。")
