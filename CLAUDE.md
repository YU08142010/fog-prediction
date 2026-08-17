# CLAUDE.md

このファイルは、Claude Code（claude.ai/code）がこのリポジトリで作業するときの指針です。

## このリポジトリについて

只見川流域の川霧を観測データからグラフ化し、機械学習で予測するプログラムです。

- **実行環境の前提は Google Colab**。ローカルでも動きますが、日本語フォントの自動導入・ファイルのアップロード／ダウンロード支援は Colab に合わせて作られています。
- 実体は **単一ファイル `weather_visualizer.py`（約2400行）** です。パッケージ化・モジュール分割はしていません。ファイル冒頭のモジュール docstring に全体仕様が書かれています。
- ユーザーは Colab のセルに README のコードを貼って実行します。**README の手順がそのまま動くことが最優先の受け入れ条件**です。

## ファイル構成

| ファイル | 役割 |
|---|---|
| `weather_visualizer.py` | 本体（読み込み・グラフ・学習・予測・CLI・Colab支援のすべて） |
| `test_weather_visualizer.py` | unittest（63件）。ネットワークには接続しない |
| `data.csv` | 動作確認用のサンプル（只見川流域4地点・2024年6〜9月） |
| `README.md` | Colab利用者向けの手順書 |
| `requirements.txt` | pandas / numpy / matplotlib / openpyxl / scikit-learn / xgboost / requests |
| `review_point.md` | 利用者からの改善要望メモ |

`output_graphs/` は出力先で `.gitignore` 済み。生成物はコミットしません。

## `weather_visualizer.py` の内部構成

コード内のコメントで 0〜7 のセクションに区切られています。編集するときはこの区切りを維持してください。

| セクション | 内容 |
|---|---|
| 0 | 日本語フォントの自動検出・自動インストール（豆腐文字対策） |
| 1 | 列レイアウト定義、現象コードの色・ラベル定義 |
| 2 | セル値の正規化、日時パース、現象コード／風向の解釈 |
| 3 | Excel/CSV読み込みと列レイアウトの自動検出 |
| 4 | グラフ生成（上下2段構成・タイムライン同期） |
| 5 | 霧予測モデル（地点ごとの多クラス分類。RandomForest+XGBoostのハイブリッド） |
| 6 | 予測結果のグラフ・CSV出力 |
| 7 | メイン実行部（`run()` / `main()`）と Colab 支援 |

入口は2つ、どちらも最終的に `run()` を通ります。

- `run(input_file, output_dir, ...)` — Colab のセルから呼ぶ想定のエントリーポイント
- `main(argv)` → `_parse_args()` → `run()` — CLI（`python weather_visualizer.py ...`）

## 開発コマンド

```bash
# セットアップ（ローカル）
python -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt

# テスト（ネットワーク不要・約12秒）
python -m unittest test_weather_visualizer

# サンプルデータで動作確認（予測なし＝ネットワーク不要・数秒）
python -c "from weather_visualizer import run; run('data.csv', forecast=False)"

# 予測まで含めた通し確認（Open-Meteo APIに接続・数十秒）
python weather_visualizer.py data.csv

# 日本語フォントの確認だけ
python weather_visualizer.py --check-font
```

このリポジトリには `.venv/` が置かれています。ローカル確認は `.venv/bin/python` を使ってください。

## 変更するときの注意点

### Colab 前提を壊さない

- `matplotlib.use("Agg")` はインポート直後に呼ばれています。画面のない環境で描画するために必要なので外さないでください。
- Colab 依存の機能（`google.colab` の import）は必ず遅延 import と try/except で囲み、**ローカルやテストでも動くようにする**こと。既存の `_in_notebook()` / `_try_colab_upload()` / `zip_outputs()` のパターンに合わせます。
- `_parse_args()` はノートブック環境では `sys.argv` を読みません（Colab のカーネル引数を出力フォルダと誤認する不具合があったため）。この分岐を消さないでください。
- 追加ライブラリを増やさないこと。Colab に標準で入っていないものを入れると、利用者に `pip install` の手間が増えます。増やす場合は `requirements.txt` と README の両方を更新します。

### データの扱いで壊しやすいところ

- 現象コードの `/`（現象なし＝コード0）と**空欄（未入力＝欠測 NaN）は必ず区別**します。混同すると学習データが汚染されます（`encode_phenomena_cell()`）。
- 気象庁形式の `24時` は**翌日の0時**として扱います（`parse_datetime_value()`）。
- 降水量は列が無い／空欄なら **0mm** として学習に使います（無降水時に空欄という記録が多いため）。他の要素と同じ扱いにしないこと。
- Open-Meteo の取得では `wind_speed_unit=ms` を必ず指定します。既定は km/h で、指定を落とすと学習データ（m/s）と単位が食い違い、予測が静かに壊れます。
- 学習の train/test 分割は**時系列順を保った後ろ20%**です。シャッフル分割にすると未来のデータが学習に混ざります（`_temporal_train_test_split()`）。
- ハイブリッドの混合比は**学習データのさらに後ろ20%（検証期間）だけ**で決めます（`_choose_blend_weight()`）。テスト期間で選ぶとスコアが甘くなります。
- `xgboost` の import は try/except で囲み、無い環境では RandomForest 単独に降格します（`_HAS_XGBOOST` / `_resolve_model_kind()`）。この分岐を消さないこと。
- XGBoost はラベルが 0,1,2,… と連続している必要があります。現象コードは歯抜けになるため、`_XGBLabelSafeClassifier` が符号化を担っています。XGBoost を直接使う書き方に戻さないこと。

### 出力ファイル名

`{接頭辞}_①…` の接頭辞は**入力ファイル名から生成**されます（`run()` 内、`_` か半角スペースの手前まで）。観測地点名ではありません。README を書き換えるときに混同しないこと。

### 表示メッセージ

「Excelから正しく読めているか」「Open-Meteoから取れているか」を利用者が毎回確認できるよう、読み込み内容・検出地点・取得URL・件数・期間を実行時に表示しています（`_print_load_report()` / `report_phenomena_quality()`）。**この確認出力は利用者の要望で入っている機能なので、静かにしないでください。**

## 変更後に必ず行うこと

1. `python -m unittest test_weather_visualizer` が `OK` になること
2. `python weather_visualizer.py data.csv` が最後まで通ること（グラフ・予測を変更した場合）
3. 利用手順に影響する変更をしたら **README.md も同時に更新**すること
4. グラフの見た目を変えた場合は、生成された PNG の**日本語が □ になっていないか**を目視確認すること

## コミット規約

- コミットメッセージは英語。
- `git push` はユーザーの明示的な確認を得てから行うこと。
- コミット前に差分の要約（何を・なぜ）をユーザーに提示すること。
- 応答は日本語。絵文字は使わない。
