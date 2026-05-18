import lightgbm as lgb
import pandas as pd

df = pd.read_csv("data.csv")

X = df[["気温(℃)", "風速(m/s)", "相対湿度(％)", "露点温度(℃)"]]
y = df["蒲生水道橋"].fillna(0)

model = lgb.LGBMClassifier(max_depth=3, class_weight="balanced", random_state=42)
model.fit(X, y)

df["予測結果"] = model.predict(X)

probs = model.predict_proba(X)
classes = model.classes_

for i, label in enumerate(classes):
    df[f"レベル{int(label)}確率(%)"] = probs[:, i] * 100

print(
    df[
        [
            "気温(℃)",
            "風速(m/s)",
            "相対湿度(％)",
            "露点温度(℃)",
            "蒲生水道橋",
            "予測結果",
        ]
    ].head()
)
