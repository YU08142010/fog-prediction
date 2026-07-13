# -*- coding: utf-8 -*-
"""
気象庁ダウンロード形式データ（只見_2026年1_5月_気象データ.xlsx など）を
自動で読み込んでグラフ化するプログラム。

【読み込む気象データ】
  B列: 気温(℃)
  C列: 降水量(mm)
  D列: 風速(m/s)
  G列: 相対湿度(％)
  H列: 露点温度(℃)

【J〜AP列（現象コード）】
  "/"      = その時刻に現象なし
  空白     = まだデータ未入力
  1〜10    = 現象発生コード（下記の意味）
      1  薄い川霧
      2  川霧
      3  濃い川霧
      4  薄い全体霧
      5  全体霧
      6  全体濃い霧
      7  薄い層雲
      8  濃い層雲
      9  霧雨
      10 雨

【出力グラフ（5項目 × 月ごとに分割）】
  ① 気温 × 現象コード（「/」・1〜10）
  ② 降水量 × 現象コード
  ③ 風速 × 現象コード
  ④ 相対湿度 × 現象コード
  ⑤ 露点温度 × 現象コード

  各グラフは、気象データの時系列（線 or 棒グラフ）と、
  その時刻に記録された現象コード（「/」・1〜10）を“レーン”形式で
  分けて表示する2段構成です。データ期間が長い場合に1枚の画像へ
  詰め込みすぎて見づらくならないよう、「年-月」ごとにファイルを
  分割して出力します（例: 5ヶ月分のデータなら、1項目につき5枚、
  合計25枚のPNGが生成されます）。

使い方:
    python3 weather_visualizer.py 入力ファイル.xlsx [出力フォルダ]

Jupyter / Google Colab で直接セルに貼って実行する場合は、
下の DEFAULT_INPUT_FILE / DEFAULT_OUTPUT_DIR を編集してください。
（ファイルが見つからない場合、Colabなら自動でアップロード画面が出ます）
"""

import sys
import os
import re
import glob
import subprocess
import datetime

import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")  # 画面のないサーバー環境（Colab等）でも背景で安全にグラフを描画するための設定
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import matplotlib.font_manager as fm
from matplotlib.patches import Patch
from matplotlib.collections import LineCollection
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, column_index_from_string

# --- 霧予測アドオン用の追加インポート ---
import requests
from sklearn.model_selection import train_test_split, RandomizedSearchCV
from sklearn.pipeline import Pipeline
from sklearn.preprocessing import StandardScaler
from sklearn.ensemble import RandomForestClassifier
from sklearn.linear_model import LogisticRegression
from sklearn.metrics import classification_report, confusion_matrix, roc_auc_score, accuracy_score, f1_score


# ===========================================================================
# 0. 日本語フォント自動設定・インストール機能（文字化け・豆腐文字対策）
# ===========================================================================

_CJK_FONT_KEYWORDS = [
    "noto sans cjk", "noto serif cjk", "ipaex", "ipagothic", "ipa gothic",
    "takao", "vl gothic", "yu gothic", "ms gothic", "hiragino",
    "source han sans", "droid sans fallback",
]

def _find_cjk_font():
    for f in fm.fontManager.ttflist:
        if any(k in f.name.lower() for k in _CJK_FONT_KEYWORDS):
            return f.name
    return None

def _install_noto_cjk_font():
    try:
        subprocess.run(["apt-get", "update"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=120, check=False)
        result = subprocess.run(
            ["apt-get", "install", "-y", "fonts-noto-cjk"],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=300, check=False,
        )
        return result.returncode == 0
    except Exception:
        return False

def setup_japanese_font():
    font_name = _find_cjk_font()
    if font_name is None:
        print("日本語フォントが見つからないため、自動インストールを試みます…")
        if _install_noto_cjk_font():
            for fp in glob.glob("/usr/share/fonts/**/*.[ot]t[fc]", recursive=True):
                try:
                    fm.fontManager.addfont(fp)
                except Exception:
                    pass
            font_name = _find_cjk_font()
    if font_name:
        plt.rcParams["font.family"] = font_name
        print(f"日本語フォントを設定しました: {font_name}")
    else:
        print("【警告】日本語フォントの自動設定に失敗しました。グラフ内の文字が『□』になる可能性があります。")
    plt.rcParams["axes.unicode_minus"] = False
    return font_name

setup_japanese_font()


# ===========================================================================
# 1. Excelの列レイアウト設定と現象コード（色・ラベル）の定義
# ===========================================================================

COL_DATETIME = "A"
COL_TEMP = "B"
COL_PRECIP = "E"
COL_WIND = "H"
COL_DEWPOINT = "S"
COL_HUMID = "V"

PHENOMENA_RANGE = ("AC", "BH")

MAIN_COLUMNS = {
    COL_TEMP: "気温(℃)",
    COL_PRECIP: "降水量(mm)",
    COL_WIND: "風速(m/s)",
    COL_DEWPOINT: "露点温度(℃)",
    COL_HUMID: "相対湿度(％)",
}

PHENOM_LABELS = {
    1: "薄い川霧", 2: "川霧", 3: "濃い川霧", 4: "薄い全体霧", 5: "全体霧",
    6: "全体濃い霧", 7: "薄い層雲", 8: "濃い層雲", 9: "霧雨", 10: "雨",
}

PHENOM_COLORS = {
    1: "#aed6f1", 2: "#3498db", 3: "#1a5276", 4: "#dcdde1", 5: "#909497",
    6: "#2c3e50", 7: "#f8c471", 8: "#d35400", 9: "#a9dfbf", 10: "#196f3d",
}

SLASH_COLOR = "#f9e79f"


# ===========================================================================
# 2. データ読み込みおよびデータクレンジング（品質情報等の除外）
# ===========================================================================

def find_header_row(ws, search_col="A", keyword="年月日時", max_search_rows=15):
    col_idx = column_index_from_string(search_col)
    for r in range(1, max_search_rows + 1):
        v = ws.cell(row=r, column=col_idx).value
        if v == keyword:
            return r
    raise ValueError(f"ヘッダー行（{search_col}列に「{keyword}」）が見つかりませんでした。気象庁のデータか確認してください。")


def load_weather_data(filepath, sheet_name=None):
    wb = load_workbook(filepath, data_only=True)
    ws = wb[sheet_name] if sheet_name else wb[wb.sheetnames[0]]

    header_row = find_header_row(ws, COL_DATETIME)
    data_start_row = header_row + 1
    dt_col_idx = column_index_from_string(COL_DATETIME)

    start_idx = column_index_from_string(PHENOMENA_RANGE[0])
    end_idx = column_index_from_string(PHENOMENA_RANGE[1])
    phenom_cols = [get_column_letter(c) for c in range(start_idx, end_idx + 1)]

    location_mapping = {}
    for col_letter in phenom_cols:
        cidx = column_index_from_string(col_letter)
        loc_name = ws.cell(row=3, column=cidx).value
        if loc_name:
            loc_name = str(loc_name).strip().replace("\n", "").replace("\r", "")
        else:
            loc_name = f"地点_{col_letter}"
        location_mapping[col_letter] = loc_name

    main_rows = []
    phenom_rows = []

    for r in range(data_start_row, ws.max_row + 1):
        date_val = ws.cell(row=r, column=dt_col_idx).value
        if date_val is None:
            continue

        mrow = {"datetime": date_val}
        for col_letter, label in MAIN_COLUMNS.items():
            cidx = column_index_from_string(col_letter)
            mrow[label] = ws.cell(row=r, column=cidx).value
        main_rows.append(mrow)

        prow = {"datetime": date_val}
        for col_letter in phenom_cols:
            cidx = column_index_from_string(col_letter)
            prow[col_letter] = ws.cell(row=r, column=cidx).value
        phenom_rows.append(prow)

    main_df = pd.DataFrame(main_rows)
    main_df["datetime"] = pd.to_datetime(main_df["datetime"], errors="coerce")
    main_df = main_df.dropna(subset=["datetime"])
    for label in MAIN_COLUMNS.values():
        main_df[label] = pd.to_numeric(main_df[label], errors="coerce")
    main_df = main_df.sort_values("datetime").reset_index(drop=True)

    phenom_df = pd.DataFrame(phenom_rows)
    phenom_df["datetime"] = pd.to_datetime(phenom_df["datetime"], errors="coerce")
    phenom_df = phenom_df.dropna(subset=["datetime"])
    phenom_df = phenom_df.sort_values("datetime").reset_index(drop=True)

    return main_df, phenom_df, phenom_cols, location_mapping


def encode_phenomena_cell(value):
    if value is None:
        return np.nan
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).strip()
    if s == "/":
        return 0.0
    nums = re.findall(r"\d+", s)
    if nums:
        return float(max(int(n) for n in nums))
    return np.nan


# ===========================================================================
# 3. グラフ生成・描画処理（上下2段構成・タイムライン同期システム）
# ===========================================================================

DAYS_PER_INCH = 0.55
MIN_FIG_WIDTH = 16
MAX_FIG_WIDTH = 60
SAVE_DPI = 150


def compute_fig_width(times):
    total_days = (times.max() - times.min()).total_seconds() / 86400.0
    width = max(total_days, 1.0) / DAYS_PER_INCH
    return float(np.clip(width, MIN_FIG_WIDTH, MAX_FIG_WIDTH))


def split_by_month_two(main_df, phenom_df):
    m = main_df.copy()
    p = phenom_df.copy()
    m["__ym"] = m["datetime"].dt.strftime("%Y-%m")
    p["__ym"] = p["datetime"].dt.strftime("%Y-%m")
    groups = []
    for ym in sorted(set(m["__ym"]) | set(p["__ym"])):
        msub = m[m["__ym"] == ym].drop(columns="__ym").reset_index(drop=True)
        psub = p[p["__ym"] == ym].drop(columns="__ym").reset_index(drop=True)
        groups.append((ym, msub, psub))
    return groups


def _draw_lane_panel(ax2, phenom_df, phenom_cols, location_mapping):
    ptimes = phenom_df["datetime"]
    n_lanes = len(phenom_cols)

    for row, col_letter in enumerate(phenom_cols):
        col_values = phenom_df[col_letter].map(encode_phenomena_cell).to_numpy()

        slash_mask = col_values == 0
        if slash_mask.any():
            x_vals = mdates.date2num(ptimes[slash_mask])
            ax2.vlines(x_vals, row - 0.35, row + 0.35,
                       color=SLASH_COLOR, linewidth=1.0, alpha=0.9, zorder=2)

        for code in range(1, 11):
            code_mask = col_values == code
            if code_mask.any():
                x_vals = mdates.date2num(ptimes[code_mask])
                ax2.vlines(x_vals, row - 0.42, row + 0.42,
                           color=PHENOM_COLORS[code], linewidth=2.4, zorder=3)

    for row in range(n_lanes + 1):
        ax2.axhline(row - 0.5, color="#ececec", linewidth=0.5, zorder=1)

    ax2.set_yticks(range(n_lanes))
    ax2.set_yticklabels([location_mapping[c] for c in phenom_cols], fontsize=8)
    ax2.set_ylim(n_lanes - 0.5, -0.5)
    ax2.set_xlabel("日時")
    ax2.set_ylabel("現象記録地点", fontsize=10)

    legend_handles = [Patch(facecolor=SLASH_COLOR, label="「/」現象なし")]
    for code in range(1, 11):
        legend_handles.append(Patch(facecolor=PHENOM_COLORS[code], label=f"{code}: {PHENOM_LABELS[code]}"))
    ax2.legend(handles=legend_handles, bbox_to_anchor=(1.0, 1.0), loc="upper left",
               fontsize=8, borderaxespad=0., title="現象コード", title_fontsize=9)


def _finalize_figure(fig, ax1, ax2, times, fig_width):
    ax2.xaxis_date()
    ax2.xaxis.set_major_formatter(mdates.DateFormatter("%m/%d"))
    n_ticks_target = max(10, int(fig_width / 1.3))
    ax2.xaxis.set_major_locator(
        mdates.AutoDateLocator(minticks=n_ticks_target, maxticks=n_ticks_target * 2)
    )
    ax1.set_xlim(times.min(), times.max())
    fig.autofmt_xdate()


def _make_fig(times, phenom_cols):
    n_lanes = len(phenom_cols)
    fig_width = compute_fig_width(times)
    lane_panel_height = max(6.0, n_lanes * 0.24)
    fig_height = 5.5 + lane_panel_height
    fig, (ax1, ax2) = plt.subplots(
        2, 1, figsize=(fig_width, fig_height), sharex=True,
        gridspec_kw={"height_ratios": [5.5, lane_panel_height], "hspace": 0.04},
    )
    return fig, ax1, ax2, fig_width


def plot_temp_humid_dew(main_df, phenom_df, phenom_cols, location_mapping, location_name, out_path):
    times = main_df["datetime"]
    fig, ax1, ax2, fig_width = _make_fig(times, phenom_cols)

    l1, = ax1.plot(times, main_df["気温(℃)"],    color="#e74c3c", linewidth=1.1, label="気温(℃)", zorder=3)
    l2, = ax1.plot(times, main_df["露点温度(℃)"], color="#16a085", linewidth=1.1, label="露点温度(℃)", zorder=3)
    ax1.set_ylabel("気温・露点温度（℃）", fontsize=10)
    ax1.grid(True, alpha=0.25)

    ax1r = ax1.twinx()
    l3, = ax1r.plot(times, main_df["相対湿度(％)"], color="#8e44ad", linewidth=0.9, alpha=0.65, label="相対湿度(％)", zorder=2)
    ax1r.set_ylabel("相対湿度（％）", fontsize=10)
    ax1r.set_ylim(0, 115)

    ax1.set_title(f"【{location_name}】気温・露点温度（左軸℃）・相対湿度（右軸%） と 各地点の現象コードの関係", fontsize=13)
    ax1.legend(handles=[l1, l2, l3], loc="upper left", fontsize=9)

    _draw_lane_panel(ax2, phenom_df, phenom_cols, location_mapping)
    _finalize_figure(fig, ax1, ax2, times, fig_width)

    fig.savefig(out_path, dpi=SAVE_DPI, bbox_inches="tight")
    plt.close(fig)


def plot_wind_precip(main_df, phenom_df, phenom_cols, location_mapping, location_name, out_path):
    times = main_df["datetime"]
    fig, ax1, ax2, fig_width = _make_fig(times, phenom_cols)

    l1, = ax1.plot(times, main_df["風速(m/s)"], color="#27ae60", linewidth=1.1, label="風速(m/s)", zorder=3)
    ax1.set_ylabel("風速（m/s）", fontsize=10)
    ax1.grid(True, alpha=0.25)

    ax1r = ax1.twinx()
    ax1r.bar(times, main_df["降水量(mm)"], width=0.03, color="#2980b9", alpha=0.55, label="降水量(mm)", zorder=2)
    pmax = main_df["降水量(mm)"].dropna().max() if not main_df["降水量(mm)"].dropna().empty else 1.0
    ax1r.set_ylim(0, max(pmax * 3.5, 2.0))
    ax1r.set_ylabel("降水量（mm）", fontsize=10)

    ax1.set_title(f"【{location_name}】風速（左軸m/s）・降水量（右軸mm） と 各地点の現象コードの関係", fontsize=13)
    ax1.legend(handles=[l1, Patch(color="#2980b9", alpha=0.55, label="降水量(mm)")], loc="upper left", fontsize=9)

    _draw_lane_panel(ax2, phenom_df, phenom_cols, location_mapping)
    _finalize_figure(fig, ax1, ax2, times, fig_width)

    fig.savefig(out_path, dpi=SAVE_DPI, bbox_inches="tight")
    plt.close(fig)


def plot_combo_by_month(main_df, phenom_df, phenom_cols, location_mapping, location_name, out_dir):
    for ym, msub, psub in split_by_month_two(main_df, phenom_df):
        if msub["datetime"].nunique() < 2:
            continue
        loc = f"{location_name}（{ym}）"
        plot_temp_humid_dew(
            msub, psub, phenom_cols, location_mapping, loc,
            os.path.join(out_dir, f"{location_name}_①気温・湿度・露点×現象コード_{ym}.png"),
        )
        plot_wind_precip(
            msub, psub, phenom_cols, location_mapping, loc,
            os.path.join(out_dir, f"{location_name}_②風速・降水量×現象コード_{ym}.png"),
        )


# ===========================================================================
# 4. メイン実行部および Google Colab 支援機能
# ===========================================================================

DEFAULT_INPUT_FILE = "気象データ.xlsx"
DEFAULT_OUTPUT_DIR = "./output_graphs"


def _try_colab_upload():
    if "google.colab" not in sys.modules:
        return None
    try:
        from google.colab import files as colab_files
    except ImportError:
        return None
    print("【確認】入力ファイルが指定のパスに見つかりません。")
    print("アップロード画面を表示しますので、解析したい気象データのExcelファイルを選択してください…")
    uploaded = colab_files.upload()
    for name in uploaded.keys():
        if name.lower().endswith((".xlsx", ".xls")):
            return os.path.abspath(name)
    return None


def _resolve_input_file(filepath):
    if filepath and os.path.isfile(filepath):
        return filepath
    if "google.colab" in sys.modules:
        uploaded_path = _try_colab_upload()
        if uploaded_path and os.path.isfile(uploaded_path):
            return uploaded_path
    if "ipykernel" in sys.modules or "IPython" in sys.modules:
        try:
            entered = input("入力ファイルのパスを入力してください（Enterでキャンセル): ").strip()
            if entered and os.path.isfile(entered):
                return entered
        except Exception:
            pass
    return None


def _parse_args():
    args = [a for a in sys.argv[1:] if not a.startswith("-")]
    args = [a for a in args if a.lower().endswith((".xlsx", ".xls")) or os.path.isdir(a)]
    filepath = None
    out_dir = None
    for a in args:
        if a.lower().endswith((".xlsx", ".xls")): filepath = a
        else: out_dir = a
    if filepath is None: filepath = DEFAULT_INPUT_FILE
    if out_dir is None: out_dir = DEFAULT_OUTPUT_DIR
    return filepath, out_dir


def main():
    filepath, out_dir = _parse_args()
    filepath = _resolve_input_file(filepath)

    if filepath is None:
        print("【エラー】入力ファイルが指定されなかったか、見つかりませんでした。")
        print("ファイル名が正しく『気象データ.xlsx』になっているか、またはフォルダーに正しく配置されているか確認してください。")
        sys.exit(1)

    os.makedirs(out_dir, exist_ok=True)

    print(f"Excelファイルを確認中: {filepath}")
    main_df, phenom_df, phenom_cols, location_mapping = load_weather_data(filepath)

    base = os.path.splitext(os.path.basename(filepath))[0]
    location_name = base.split("_")[0].split(" ")[0] if ("_" in base or " " in base) else base

    print(f"データ解析成功: {location_name} (計 {len(phenom_cols)} 地点の現象レーンを検出しました)")
    print("月ごと・全地点統合グラフの自動生成を開始します...")

    plot_combo_by_month(main_df, phenom_df, phenom_cols, location_mapping, location_name, out_dir)

    print(f"\nすべてのグラフが正常に生成されました！")
    print(f"出力先フォルダ: {out_dir} の中を確認してください。")

    # === 【追加機能】霧予測モデルの学習・評価・16日間予測グラフ ===
    run_fog_prediction_addon(main_df, phenom_df, phenom_cols, location_name, out_dir, location_mapping)


#　=======================================================================
#　グラフに関しての説明
#　=======================================================================

"""
生成されたグラフでは、結合された気象要素（気温、相対湿度など）と日付、気象コード（霧発生など）すべてが適応しているかどうかの整合性チェックはしました。
グラフの見方は、日付が書いてあるメモリの位置はその日の正午となっています。
使用できるファイスの形式は、teamsの投稿にあがっているような形式のみです。
現象コードの色一覧

　＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
　 /　薄い黄色　「現象なし」
 　1　薄い水色　「薄い川霧」
 　2　水色　　　「川霧」
 　3　濃い水色　「濃い川霧」
 　4　薄い灰色　「薄い全体霧」
 　5　灰色　　　「全体霧」
 　6　紺色　　　「全体濃い霧」
 　7　オレンジ色「薄い層雲」
 　8　赤色　　　「濃い層雲」
 　9　薄い緑色　「霧雨」
 　10 緑色　　　「雨」
　＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
"""




# ===========================================================================
# 5. 霧予測モデル（追加機能・地点ごと個別予測版）
# ---------------------------------------------------------------------------
# 既存の可視化処理（1〜4）は一切変更していない。
#
# ・「32地点のどこかで霧」ではなく、【地点ごとに個別】の予測モデル
#   （scikit-learn RandomForestClassifier）を学習する
# ・元データと同じ粒度で予測する: 「/」(現象なし)+ コード1〜10 の
#   【11クラスの多クラス分類】として学習する（霧かどうかの二値化はしない）
# ・Open-Meteoから今後16日分の気象予報を取得し、地点ごとに現象コードを予測
# ・その結果を、既存の「①気温・湿度・露点×現象コード」グラフと同じ見た目・
#   同じ現象レーンパネル形式・同じ配色で、直近の実測データの続きに
#   16日分追加した専用グラフとして出力する
# ===========================================================================

FORECAST_LAT = 37.4936   # 予報取得地点の緯度（会津若松付近の目安値。分かれば書き換え可）
FORECAST_LON = 139.9298  # 予報取得地点の経度
FORECAST_DAYS = 16       # Open-Meteo無料予報で取得できる最大日数

ML_FEATURES = [
    "気温(℃)", "降水量(mm)", "風速(m/s)", "相対湿度(％)", "露点温度(℃)",
    "気温露点差", "時_sin", "時_cos", "月",
]

# 地点ごとにモデルを学習するための最低条件（これを満たさない地点はスキップする）
MIN_ROWS_PER_LOCATION = 200
MIN_CLASSES_PER_LOCATION = 2  # 最低2種類以上の現象コードが記録されている必要がある

# RandomizedSearchCVで探索するハイパーパラメータの範囲
PARAM_DIST = {
    "model__n_estimators": [200, 300, 500],
    "model__max_depth": [6, 8, 10, 12, 14, None],
    "model__min_samples_leaf": [1, 2, 3, 5],
    "model__max_features": ["sqrt", "log2", None],
}


def build_location_class_targets(phenom_df, phenom_cols):
    """32地点それぞれについて『現象コード（0="/"、1〜10、欠測はNaN）』を作る。
    地点ごとに独立したラベルであり、他地点の値には一切影響されない。
    以前のバージョンは霧(1-6)か否かの二値だったが、元データと同じ粒度で
    扱えるよう、11クラス（0〜10）の多クラス分類用ターゲットに変更した。
    """
    codes = phenom_df[phenom_cols].apply(lambda col: col.map(encode_phenomena_cell))
    targets = codes.copy()
    targets.insert(0, "datetime", phenom_df["datetime"].values)
    return targets


def _add_time_features(df):
    """main_df/予報データフレームに、モデルが使う時間特徴量を追加する共通処理"""
    df = df.copy()
    df["時"] = df["datetime"].dt.hour
    df["月"] = df["datetime"].dt.month
    df["気温露点差"] = df["気温(℃)"] - df["露点温度(℃)"]
    df["時_sin"] = np.sin(2 * np.pi * df["時"] / 24)
    df["時_cos"] = np.cos(2 * np.pi * df["時"] / 24)
    return df


def train_location_models(main_df, phenom_df, phenom_cols, location_mapping):
    """32地点それぞれについて、独立したRandomForestClassifier（多クラス分類）を学習する。
    予測対象は「/」(0) + コード1〜10 の11クラス（元データと同じ粒度）。
    戻り値: {col_letter: {"pipe": Pipeline, "accuracy": float, "f1": float,
                          "n": int, "n_classes": int, "classes": list}}
    （学習条件を満たさない地点は辞書に含めない＝その地点は予測グラフでも空欄になる）
    """
    targets = build_location_class_targets(phenom_df, phenom_cols)
    merged = pd.merge(main_df, targets, on="datetime", how="inner")
    merged = _add_time_features(merged)
    merged = merged.dropna(subset=ML_FEATURES)

    print("\n" + "=" * 60)
    print("■ 地点ごとの現象コード予測モデルを学習（scikit-learn RandomForestClassifier・多クラス分類）")
    print("=" * 60)

    dummy_blocks_by_col = {col: find_dummy_blocks(phenom_df, col) for col in phenom_cols}
    all_warnings = []
    for col, blocks in dummy_blocks_by_col.items():
        loc_name = location_mapping.get(col, col)
        for start, end, ratio, n in blocks:
            all_warnings.append(
                f"{loc_name}: {start.strftime('%Y/%m/%d')}〜{end.strftime('%Y/%m/%d')}（{n}件）は"
                f"「/」の割合が{ratio:.1%}と極端に低く、ダミーデータの可能性があるため学習から除外します。"
            )
    if all_warnings:
        print("【データ品質チェック】以下の期間はダミーデータの可能性があるため、学習から除外します:")
        for w in all_warnings:
            print(f"  ⚠ {w}")
        print()

    print(f"{'地点名':<14} {'件数':>7} {'クラス数':>8} {'正解率':>7} {'F1':>6}  結果")
    print("-" * 60)

    models = {}
    for col in phenom_cols:
        loc_name = location_mapping.get(col, col)
        sub = merged.dropna(subset=[col])

        # ダミーと判定されたブロックの期間を学習データから除外する
        for start, end, _, _ in dummy_blocks_by_col.get(col, []):
            sub = sub[(sub["datetime"] < start) | (sub["datetime"] > end)]

        if len(sub) < MIN_ROWS_PER_LOCATION:
            print(f"{loc_name:<14} {len(sub):>7} {'-':>8} {'-':>7} {'-':>6}  データ不足でスキップ")
            continue

        y = sub[col].astype(int)
        class_counts = y.value_counts()
        n_classes = len(class_counts)
        if n_classes < MIN_CLASSES_PER_LOCATION:
            print(f"{loc_name:<14} {len(sub):>7} {n_classes:>8} {'-':>7} {'-':>6}  現象の種類が少なくスキップ")
            continue

        X = sub[ML_FEATURES]
        # サンプル数1件しかないクラスがあると層化分割できないため、その場合のみ層化なしにする
        stratify_arg = y if class_counts.min() >= 2 else None
        X_train, X_test, y_train, y_test = train_test_split(
            X, y, test_size=0.2, stratify=stratify_arg, random_state=42
        )

        if y_train.nunique() < 2:
            print(f"{loc_name:<14} {len(sub):>7} {n_classes:>8} {'-':>7} {'-':>6}  分割後クラス不足でスキップ")
            continue

        # 「より正確に」との要望に対応: RandomizedSearchCVで地点ごとにハイパーパラメータを
        # 自動探索する。ただし最小クラスの件数がCV分割数より少ない場合は探索できないため、
        # その場合は既定値にフォールバックする。
        base_pipe = Pipeline([
            ("scaler", StandardScaler()),
            ("model", RandomForestClassifier(
                class_weight="balanced_subsample", random_state=42, n_jobs=-1,
            )),
        ])

        n_splits = min(3, int(class_counts.min()))
        if n_splits >= 2 and len(X_train) >= 30:
            try:
                search = RandomizedSearchCV(
                    base_pipe, PARAM_DIST, n_iter=10, cv=n_splits,
                    scoring="accuracy", random_state=42, n_jobs=-1, error_score="raise",
                )
                search.fit(X_train, y_train)
                pipe = search.best_estimator_
                tuned = True
            except Exception:
                pipe = base_pipe
                pipe.set_params(model__n_estimators=500, model__max_depth=14, model__min_samples_leaf=2)
                pipe.fit(X_train, y_train)
                tuned = False
        else:
            pipe = base_pipe
            pipe.set_params(model__n_estimators=500, model__max_depth=14, model__min_samples_leaf=2)
            pipe.fit(X_train, y_train)
            tuned = False

        y_pred = pipe.predict(X_test)
        acc = accuracy_score(y_test, y_pred)
        f1 = f1_score(y_test, y_pred, average="weighted", zero_division=0)

        tag = "学習完了(調整済)" if tuned else "学習完了(既定値)"
        print(f"{loc_name:<14} {len(sub):>7} {n_classes:>8} {acc:>7.1%} {f1:>6.3f}  {tag}")
        models[col] = {
            "pipe": pipe, "accuracy": acc, "f1": f1,
            "n": len(sub), "n_classes": n_classes, "classes": sorted(y.unique().tolist()),
        }

    if models:
        accs = [m["accuracy"] for m in models.values()]
        print("-" * 60)
        print(f"学習できた地点数: {len(models)} / {len(phenom_cols)}　平均正解率: {np.mean(accs):.1%}")
    else:
        print("\n【警告】どの地点も学習条件を満たさず、モデルを1つも作れませんでした。")

    return models


def fetch_forecast_range(lat: float = FORECAST_LAT, lon: float = FORECAST_LON, forecast_days: int = None):
    """Open-Meteo（無料・APIキー不要）から今日から forecast_days 日分の
    時間別予報をまとめて1回のリクエストで取得する。
    """
    if forecast_days is None:
        forecast_days = FORECAST_DAYS
    url = "https://api.open-meteo.com/v1/forecast"
    params = {
        "latitude": lat,
        "longitude": lon,
        "hourly": "temperature_2m,precipitation,relative_humidity_2m,dew_point_2m,wind_speed_10m",
        "timezone": "Asia/Tokyo",
        "forecast_days": forecast_days,
    }
    resp = requests.get(url, params=params, timeout=20)
    resp.raise_for_status()
    data = resp.json()["hourly"]

    return pd.DataFrame({
        "datetime": pd.to_datetime(data["time"]),
        "気温(℃)": data["temperature_2m"],
        "降水量(mm)": data["precipitation"],
        "相対湿度(％)": data["relative_humidity_2m"],
        "露点温度(℃)": data["dew_point_2m"],
        "風速(m/s)": data["wind_speed_10m"],
    })


def build_location_forecast_codes(models, forecast_raw):
    """予報データフレームに対して、地点ごとに学習済みモデルで
    現象コード（0="/"、1〜10）を予測する（多クラス分類のpredict()を使用）。
    戻り値: datetime列 + 各地点(col_letter)の予測コード列 を持つDataFrame
    """
    df = _add_time_features(forecast_raw)
    X = df[ML_FEATURES]

    pred_df = pd.DataFrame({"datetime": df["datetime"]})
    for col, info in models.items():
        pred_df[col] = info["pipe"].predict(X)
    return pred_df


def get_last_observed_datetime(phenom_df, col):
    """その地点(col)について、実際に結果が入力されている最後の日時を返す。
    Excelに未来の日付の行が空欄のまま存在していても、それは無視して
    「本当に観測結果がある最後の時刻」を基準にする。
    """
    codes = phenom_df[col].map(encode_phenomena_cell)
    valid_dates = phenom_df.loc[codes.notna(), "datetime"]
    if valid_dates.empty:
        return phenom_df["datetime"].max()
    return valid_dates.max()


def find_dummy_blocks(phenom_df, col, min_block_rows: int = 50, gap_hours: float = 24 * 7,
                       max_slash_ratio: float = 0.02):
    """観測データを時間的な『まとまり(ブロック)』に分割し、各ブロックの中で
    「/」(現象なし)が極端に少ない（＝ダミーデータの疑いがある）ブロックを検出する。

    本物の観測なら大半は「/」のはずなので、まとまった件数があるのに「/」が
    ほぼ皆無なブロックは、テスト用の仮の値である可能性が高いと判断する。
    戻り値: [(開始日時, 終了日時, 「/」の割合, 件数), ...] の疑わしいブロックのリスト
    """
    codes = phenom_df[col].map(encode_phenomena_cell)
    valid_mask = codes.notna()
    valid_df = pd.DataFrame({
        "datetime": phenom_df.loc[valid_mask, "datetime"].values,
        "code": codes[valid_mask].values,
    })
    if valid_df.empty:
        return []

    # 時間差がgap_hours以上空いたら別ブロックとみなす
    time_diff = valid_df["datetime"].diff().dt.total_seconds() / 3600
    block_id = (time_diff > gap_hours).cumsum()

    suspicious = []
    for _, block in valid_df.groupby(block_id):
        if len(block) < min_block_rows:
            continue
        slash_ratio = (block["code"] == 0).mean()
        if slash_ratio <= max_slash_ratio:
            suspicious.append((block["datetime"].min(), block["datetime"].max(), slash_ratio, len(block)))
    return suspicious


def check_dummy_data_warning(phenom_df, col, location_mapping):
    """この地点にダミーデータらしきブロックが含まれていないかチェックし、
    見つかった場合は警告メッセージ（複数件あれば複数行）のリストを返す。
    """
    blocks = find_dummy_blocks(phenom_df, col)
    if not blocks:
        return []
    loc_name = location_mapping.get(col, col)
    messages = []
    for start, end, ratio, n in blocks:
        messages.append(
            f"{loc_name}: {start.strftime('%Y/%m/%d')}〜{end.strftime('%Y/%m/%d')}（{n}件）は"
            f"「/」の割合が{ratio:.1%}と極端に低く、ダミーデータの可能性があります。"
        )
    return messages


def plot_single_location_forecast(main_df, phenom_df, col, location_mapping,
                                   forecast_raw, pred_series, location_name, out_dir,
                                   history_days: int = 5):
    """1地点だけの「実測(直近history_days日) + 今後FORECAST_DAYS日分の予測」グラフを作る。
    元コードの①グラフと同じ配色（現象コード1〜10 + 「/」）を使う。

    実測の終点は「Excel上の最後の行」ではなく、【実際に結果が入力されている最後の時刻】
    を使う（未来日付の空欄行がExcelに存在していても無視する）。予報はOpen-Meteoから
    「実行した日（＝今日）」以降を取得するので、実測データが今日に近ければ両者はほぼ
    切れ目なく繋がる。
    """
    loc_name = location_mapping.get(col, col)

    last_observed = get_last_observed_datetime(phenom_df, col)
    hist_cutoff = last_observed - pd.Timedelta(days=history_days)
    hist_main = main_df[(main_df["datetime"] >= hist_cutoff) & (main_df["datetime"] <= last_observed)].reset_index(drop=True)
    hist_phenom = phenom_df[(phenom_df["datetime"] >= hist_cutoff) & (phenom_df["datetime"] <= last_observed)].reset_index(drop=True)
    hist_codes = hist_phenom[col].map(encode_phenomena_cell).to_numpy()

    all_times = pd.concat([hist_main["datetime"], forecast_raw["datetime"]])
    fig_width = compute_fig_width(all_times)
    fig, (ax1, ax2) = plt.subplots(
        2, 1, figsize=(fig_width, 6.5), sharex=True,
        gridspec_kw={"height_ratios": [5, 1.4], "hspace": 0.10},
    )

    # ーーー 上段：気温・露点温度（左軸）・相対湿度（右軸） ーーー
    # 【改善】以前は予報部分をすべて赤にしていたため、気温・露点・湿度の見分けが
    # つかなくなっていた。各項目は実測・予報とも同じ色を使い、実線=実測／点線=予報で
    # 区別する。「ここが予測ゾーンだ」という点は、背景の網掛けと境界線だけで示す。
    ax1.axvspan(forecast_raw["datetime"].min(), forecast_raw["datetime"].max(),
                color="gray", alpha=0.08, zorder=0)
    ax1.plot(hist_main["datetime"], hist_main["気温(℃)"], color="#e74c3c", linewidth=1.2, label="気温(℃) [実測]")
    ax1.plot(hist_main["datetime"], hist_main["露点温度(℃)"], color="#16a085", linewidth=1.2, label="露点温度(℃) [実測]")
    ax1.plot(forecast_raw["datetime"], forecast_raw["気温(℃)"], color="#e74c3c", linewidth=1.5,
              linestyle="--", label="気温(℃) [予報]")
    ax1.plot(forecast_raw["datetime"], forecast_raw["露点温度(℃)"], color="#16a085", linewidth=1.5,
              linestyle="--", label="露点温度(℃) [予報]")
    ax1.set_ylabel("気温・露点温度（℃）", fontsize=10)
    ax1.grid(True, axis="y", alpha=0.25)

    ax1r = ax1.twinx()
    ax1r.plot(hist_main["datetime"], hist_main["相対湿度(％)"], color="#8e44ad", linewidth=0.9, alpha=0.6,
               label="相対湿度(％) [実測]")
    ax1r.plot(forecast_raw["datetime"], forecast_raw["相対湿度(％)"], color="#8e44ad", linewidth=1.2, alpha=0.9,
               linestyle="--", label="相対湿度(％) [予報]")
    ax1r.set_ylabel("相対湿度（％）", fontsize=10)
    ax1r.set_ylim(0, 115)

    boundary = forecast_raw["datetime"].min()
    ax1.axvline(boundary, color="dimgray", linewidth=1.5, linestyle=":")
    ax1.text(mdates.date2num(boundary), ax1.get_ylim()[1], " ここから先は予測（背景グレー・点線） → ",
              fontsize=9, color="dimgray", va="top", fontweight="bold")

    n_days = max(1, round(len(forecast_raw) / 24))
    ax1.set_title(
        f"【{location_name}　{loc_name}】直近{history_days}日間の実測 ＋ 今後{n_days}日間の霧予測",
        fontsize=12,
    )
    h1, lbl1 = ax1.get_legend_handles_labels()
    h2, lbl2 = ax1r.get_legend_handles_labels()
    ax1.legend(h1 + h2, lbl1 + lbl2, loc="upper left", fontsize=8)

    # ーーー 下段：この1地点だけの現象コード（実測＋予測、既存と同じ配色） ーーー
    ax2.axvspan(forecast_raw["datetime"].min(), forecast_raw["datetime"].max(), color="gray", alpha=0.08, zorder=0)

    x_hist = mdates.date2num(hist_phenom["datetime"].to_numpy())
    slash_mask = hist_codes == 0
    if slash_mask.any():
        ax2.vlines(x_hist[slash_mask], -0.4, 0.4, color=SLASH_COLOR, linewidth=2.2, zorder=2)
    for code in range(1, 11):
        code_mask = hist_codes == code
        if code_mask.any():
            ax2.vlines(x_hist[code_mask], -0.45, 0.45, color=PHENOM_COLORS[code], linewidth=3.0, zorder=3)

    x_fore = mdates.date2num(forecast_raw["datetime"].to_numpy())
    pred_codes = pred_series.astype(int).to_numpy()
    for i, code in enumerate(pred_codes):
        color = SLASH_COLOR if code == 0 else PHENOM_COLORS.get(code, SLASH_COLOR)
        ax2.vlines(x_fore[i], -0.45, 0.45, color=color, linewidth=3.0, zorder=3)

    ax2.set_yticks([])
    ax2.set_ylabel(loc_name, fontsize=9)
    ax2.axhline(0, color="#ececec", linewidth=0.5, zorder=1)

    legend_handles = [Patch(facecolor=SLASH_COLOR, label="「/」現象なし")]
    for code in range(1, 11):
        legend_handles.append(Patch(facecolor=PHENOM_COLORS[code], label=f"{code}: {PHENOM_LABELS[code]}"))
    ax2.legend(handles=legend_handles, bbox_to_anchor=(1.0, 1.35), loc="upper left",
               fontsize=7, borderaxespad=0., title="現象コード", title_fontsize=8)

    # ーーー 横軸目盛り：年も含めた日付を1日ごとに表示 ーーー
    total_days = max(1, (all_times.max() - all_times.min()).days)
    day_interval = 1 if total_days <= 30 else max(1, total_days // 30)
    for ax in (ax1, ax2):
        ax.xaxis.set_major_locator(mdates.DayLocator(interval=day_interval))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y/%m/%d"))
        ax.grid(True, axis="x", which="major", color="gray", alpha=0.25, linewidth=0.6, zorder=0)
        ax.tick_params(axis="x", labelbottom=True, labelsize=8, labelrotation=45)
    ax1.set_xlim(all_times.min(), all_times.max())
    ax2.set_xlabel("日時（年/月/日）")
    fig.autofmt_xdate()

    out_path = os.path.join(
        out_dir,
        f"{location_name}_④{loc_name}_{n_days}日間予測_{forecast_raw['datetime'].min().strftime('%Y%m%d')}.png",
    )
    fig.savefig(out_path, dpi=SAVE_DPI, bbox_inches="tight")
    plt.close(fig)
    return out_path


SUMMARY_FOG_CODES = {1, 2, 3, 4, 5, 6}  # サマリー集計で「霧」とみなすコード


def plot_all_location_summary(pred_df, models, location_mapping, location_name, out_dir):
    """32地点ぶんの予測結果をまとめ、日ごとに『霧(コード1〜6)が予測された地点数』を
    棒グラフにする。個別グラフを1枚ずつ見なくても全体感がつかめるようにするための追加機能。
    """
    df = pred_df.copy()
    df["date"] = df["datetime"].dt.date
    fog_cols = list(models.keys())
    n_total = len(fog_cols)

    daily_counts = []
    for date, group in df.groupby("date"):
        count = 0
        for col in fog_cols:
            if group[col].astype(int).isin(SUMMARY_FOG_CODES).any():
                count += 1
        daily_counts.append((date, count))

    dates = [d for d, _ in daily_counts]
    counts = [c for _, c in daily_counts]

    fig, ax = plt.subplots(figsize=(max(10, len(dates) * 0.8), 5))
    bars = ax.bar(dates, counts, color="#3498db", width=0.6)
    for bar, c in zip(bars, counts):
        ax.text(bar.get_x() + bar.get_width() / 2, c + 0.3, str(c),
                ha="center", va="bottom", fontsize=9)

    ax.set_ylim(0, n_total + 1)
    ax.set_ylabel(f"霧(コード1〜6)が予測された地点数（全{n_total}地点中）", fontsize=10)
    ax.set_xlabel("日付")
    ax.set_title(f"【{location_name}】日ごとの霧予測サマリー（今後{len(dates)}日間）", fontsize=13)
    ax.grid(True, axis="y", alpha=0.3)
    fig.autofmt_xdate()

    out_path = os.path.join(
        out_dir, f"{location_name}_⑤日別霧予測サマリー_{dates[0].strftime('%Y%m%d') if dates else 'na'}.png"
    )
    fig.savefig(out_path, dpi=SAVE_DPI, bbox_inches="tight")
    plt.close(fig)
    return out_path


def run_fog_prediction_addon(main_df, phenom_df, phenom_cols, location_name, out_dir, location_mapping):
    """main()の最後に呼ばれる追加エントリーポイント。
    地点ごとにモデルを学習し、今後FORECAST_DAYS日分の気象予報を自動取得して、
    【32地点それぞれ個別に】予測グラフを生成する（1枚にまとめない）。

    学習には「実際に結果が入力されている行」だけを使う（未入力の行は自動的に除外
    される＝build_location_class_target/train_location_modelsのdropna処理）。
    予測はOpen-Meteoから「実行した日（今日）」以降の気象予報を取得して行う。
    """
    # ーーー 【確認①】地点の一覧（列 → 地点名の対応）をそのまま表示 ーーー
    # Excelの3行目から読み取った地点名がずれていないか、ここで目視確認できるようにする
    print("\n" + "=" * 60)
    print("■ 【確認】検出された32地点（列 → 地点名）")
    print("=" * 60)
    for col in phenom_cols:
        print(f"  {col}列 → {location_mapping.get(col, '(不明)')}")

    # ーーー 【確認②】Excelから実際に読み込んだ気象データの範囲を表示 ーーー
    print("\n" + "=" * 60)
    print("■ 【確認】Excelから読み込んだ気象データ（main_df）")
    print("=" * 60)
    print(f"  行数: {len(main_df)}件")
    print(f"  期間: {main_df['datetime'].min()} 〜 {main_df['datetime'].max()}")
    print("  先頭5行:")
    print(main_df.head(5).to_string(index=False))
    print("  末尾5行:")
    print(main_df.tail(5).to_string(index=False))

    models = train_location_models(main_df, phenom_df, phenom_cols, location_mapping)
    if not models:
        print("\n【予測モデル】1地点も学習できなかったため、予測グラフの生成をスキップしました。")
        return

    print("\n" + "=" * 60)
    print(f"■ 今後{FORECAST_DAYS}日間の気象予報から地点ごとの現象コード予測グラフを生成")
    print("=" * 60)
    try:
        forecast_raw = fetch_forecast_range()
    except requests.exceptions.RequestException as e:
        print(f"\n(予報取得に失敗しました: {e})")
        print("→ インターネットに出られる環境（ご自身のPCやRaspberry Piなど）で実行してください。")
        return

    # ーーー 【確認③】Open-Meteoから実際に取得した予報データの中身を表示 ーーー
    print(f"  取得元URL: https://api.open-meteo.com/v1/forecast")
    print(f"  取得地点: 緯度{FORECAST_LAT}, 経度{FORECAST_LON}")
    print(f"  取得件数: {len(forecast_raw)}件")
    print(f"  取得期間: {forecast_raw['datetime'].min()} 〜 {forecast_raw['datetime'].max()}")
    print("  先頭3行:")
    print(forecast_raw.head(3).to_string(index=False))

    pred_df = build_location_forecast_codes(models, forecast_raw)

    generated_paths = []
    for col in phenom_cols:
        if col not in models:
            continue

        last_observed = get_last_observed_datetime(phenom_df, col)
        gap_days = (forecast_raw["datetime"].min() - last_observed).days
        if gap_days > 3:
            loc_name = location_mapping.get(col, col)
            print(f"  【注意】{loc_name}: 最後の観測記録（{last_observed.strftime('%Y/%m/%d')}）から"
                  f"予報開始（{forecast_raw['datetime'].min().strftime('%Y/%m/%d')}）まで{gap_days}日の空白があります。")

        out_path = plot_single_location_forecast(
            main_df, phenom_df, col, location_mapping, forecast_raw, pred_df[col],
            location_name, out_dir,
        )
        generated_paths.append(out_path)
        loc_name = location_mapping.get(col, col)
        counts = pred_df[col].astype(int).value_counts().sort_index()
        breakdown = ", ".join(f"{('/' if c == 0 else c)}:{n}時間" for c, n in counts.items())
        print(f"  [{loc_name}] {out_path}")
        print(f"      予測内訳: {breakdown}")

    print(f"\n{len(generated_paths)}地点分の予測グラフを生成しました（{out_dir} 内）。")

    # ーーー 【追加機能】全地点サマリーグラフ ーーー
    summary_path = plot_all_location_summary(pred_df, models, location_mapping, location_name, out_dir)
    print(f"全地点サマリーグラフを生成しました: {summary_path}")


if __name__ == "__main__":
    # このスクリプトが直接実行された場合にmain関数を動かす
    # （main()の最後でグラフ生成 → 地点ごとの霧予測モデルの学習・16日間予測グラフ生成まで実行される）
    main()
