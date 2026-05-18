import pandas as pd
from sklearn.tree import DecisionTreeClassifier

df = pd.read_csv("data.csv")

X = df[["気温(℃)", "風速(m/s)", "相対湿度(％)"]]
y = df["蒲生水道橋"].fillna(0)

model = DecisionTreeClassifier(class_weight="balanced", random_state=42)
model.fit(X, y)

new_data = pd.DataFrame(
    [[24.5, 0.5, 93.0]], columns=["気温(℃)", "風速(m/s)", "相対湿度(％)"]
)

prediction = model.predict(new_data)
print(prediction[0])

print(pd.read_csv("data.csv")["蒲生水道橋"].unique())
