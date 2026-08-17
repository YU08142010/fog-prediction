# -*- coding: utf-8 -*-
"""weather_visualizer.py の自動テスト。

実行方法:
    python3 -m unittest -v test_weather_visualizer

ネットワークには一切アクセスしません（Open-Meteoへの通信はダミーに差し替えます）。
"""

import os
import shutil
import tempfile
import unittest
from datetime import datetime
from unittest import mock

os.environ.setdefault("WEATHER_VIZ_QUIET", "1")
os.environ.setdefault("WEATHER_VIZ_NO_FONT_INSTALL", "1")

import numpy as np
import pandas as pd

import weather_visualizer as wv


# ---------------------------------------------------------------------------
# 日時の解釈
# ---------------------------------------------------------------------------

class TestParseDatetime(unittest.TestCase):
    def test_japanese_format(self):
        self.assertEqual(wv.parse_datetime_value("2024年6月28日1時"),
                         pd.Timestamp("2024-06-28 01:00"))

    def test_hour_24_means_next_day(self):
        # 気象庁形式の「24時」は翌日の0時
        self.assertEqual(wv.parse_datetime_value("2024年9月25日24時"),
                         pd.Timestamp("2024-09-26 00:00"))
        self.assertEqual(wv.parse_datetime_value("2024/9/25 24:00"),
                         pd.Timestamp("2024-09-26 00:00"))

    def test_slash_and_iso_formats(self):
        self.assertEqual(wv.parse_datetime_value("2024/6/28 1:00"),
                         pd.Timestamp("2024-06-28 01:00"))
        self.assertEqual(wv.parse_datetime_value("2024-06-28 01:00:00"),
                         pd.Timestamp("2024-06-28 01:00"))

    def test_datetime_object_passthrough(self):
        self.assertEqual(wv.parse_datetime_value(datetime(2024, 6, 28, 1)),
                         pd.Timestamp("2024-06-28 01:00"))

    def test_excel_serial_number(self):
        # 45471 = 2024-06-28（1900年日付システム）
        self.assertEqual(wv.parse_datetime_value(45471), pd.Timestamp("2024-06-28 00:00"))

    def test_missing_year_uses_previous_row_not_current_year(self):
        """年が省略された行が『実行した年』に化けないこと（旧版の不具合）。"""
        values = ["2024年12月31日23時", "1月1日1時", "1月1日2時"]
        parsed, unparsed = wv.parse_datetime_series(values)
        self.assertEqual(list(parsed), [pd.Timestamp("2024-12-31 23:00"),
                                        pd.Timestamp("2025-01-01 01:00"),
                                        pd.Timestamp("2025-01-01 02:00")])
        self.assertEqual(unparsed, [])

    def test_missing_year_before_any_year_uses_first_found_year(self):
        values = ["6月28日1時", "2024年6月28日2時"]
        parsed, _ = wv.parse_datetime_series(values)
        self.assertEqual(parsed[0], pd.Timestamp("2024-06-28 01:00"))

    def test_unparsable_values_are_reported(self):
        parsed, unparsed = wv.parse_datetime_series(["2024年6月28日1時", "合計", None, ""])
        self.assertTrue(pd.isna(parsed[1]))
        self.assertEqual(unparsed, ["合計"])


# ---------------------------------------------------------------------------
# 現象コード・風向の解釈
# ---------------------------------------------------------------------------

class TestEncoding(unittest.TestCase):
    def test_slash_is_zero(self):
        self.assertEqual(wv.encode_phenomena_cell("/"), 0.0)
        self.assertEqual(wv.encode_phenomena_cell("／"), 0.0)  # 全角スラッシュ

    def test_blank_is_nan(self):
        for v in (None, "", "   ", np.nan):
            self.assertTrue(np.isnan(wv.encode_phenomena_cell(v)))

    def test_codes(self):
        self.assertEqual(wv.encode_phenomena_cell(3), 3.0)
        self.assertEqual(wv.encode_phenomena_cell("10"), 10.0)
        self.assertEqual(wv.encode_phenomena_cell("１０"), 10.0)  # 全角数字
        self.assertEqual(wv.encode_phenomena_cell("0"), 0.0)

    def test_out_of_range_is_nan(self):
        self.assertTrue(np.isnan(wv.encode_phenomena_cell(99)))
        self.assertTrue(np.isnan(wv.encode_phenomena_cell("あ")))

    def test_wind_direction(self):
        self.assertEqual(wv.encode_wind_direction("南"), 180.0)
        self.assertEqual(wv.encode_wind_direction("北北西"), 337.5)
        self.assertTrue(np.isnan(wv.encode_wind_direction("静穏")))
        self.assertEqual(wv.encode_wind_direction(90), 90.0)

    def test_wind_direction_features_are_cyclic(self):
        df = pd.DataFrame({
            "datetime": pd.to_datetime(["2024-06-28 01:00", "2024-06-28 02:00"]),
            "気温(℃)": [16.0, 16.0], "露点温度(℃)": [15.0, 15.0],
            wv.WIND_DIR_LABEL: [0.0, np.nan],  # 北 と 静穏
        })
        out = wv._add_time_features(df)
        self.assertAlmostEqual(out["風向_cos"].iloc[0], 1.0)
        self.assertAlmostEqual(out["風向_sin"].iloc[0], 0.0)
        # 静穏は sin=cos=0（欠測として行が落ちないこと）
        self.assertEqual(out["風向_sin"].iloc[1], 0.0)
        self.assertEqual(out["風向_cos"].iloc[1], 0.0)


# ---------------------------------------------------------------------------
# 列レイアウトの自動検出
# ---------------------------------------------------------------------------

def _jma_like_grid():
    """気象庁ダウンロード形式に似たグリッド（品質情報などの補助列つき）。"""
    header = ["年月日時", "気温(℃)", "品質情報", "均質番号", "降水量(mm)", "現象なし情報",
              "風速(m/s)", "風向", "露点温度(℃)", "相対湿度(％)", None, None, None]
    name_row = [None] * 10 + ["伊南川合流点", "柴倉橋", "堅盤橋"]
    grid = [["只見川 観測データ"] + [None] * 12, name_row, header]
    for i in range(24):
        grid.append([f"2024年7月1日{i + 1}時", 20.0 + i * 0.1, 8, 1, 0.0, None,
                     1.2, "南南西", 18.0, 88, "/", "1", None])
    return grid


class TestLayoutDetection(unittest.TestCase):
    def test_detects_measures_and_skips_metadata_columns(self):
        grid = _jma_like_grid()
        header_row = wv.find_header_row(grid)
        self.assertEqual(header_row, 2)
        info = wv.detect_layout(grid, header_row, layout="auto")
        self.assertEqual(info["measures"]["気温(℃)"], 1)      # 品質情報(2)ではない
        self.assertEqual(info["measures"]["降水量(mm)"], 4)
        self.assertEqual(info["measures"]["風速(m/s)"], 6)
        self.assertEqual(info["measures"]["露点温度(℃)"], 8)
        self.assertEqual(info["measures"]["相対湿度(％)"], 9)
        self.assertEqual(info["wind_dir"], 7)

    def test_detects_location_columns_only(self):
        grid = _jma_like_grid()
        info = wv.detect_layout(grid, wv.find_header_row(grid), layout="auto")
        names = [name for _, _, name in info["phenom"]]
        self.assertEqual(names, ["伊南川合流点", "柴倉橋", "堅盤橋"])
        self.assertEqual([letter for letter, _, _ in info["phenom"]], ["K", "L", "M"])

    def test_fixed_layout_uses_hardcoded_columns(self):
        grid = _jma_like_grid()
        info = wv.detect_layout(grid, wv.find_header_row(grid), layout="fixed")
        self.assertEqual(info["source"], "fixed")
        self.assertEqual(info["measures"]["気温(℃)"], 1)  # B列
        self.assertEqual(info["measures"]["降水量(mm)"], 4)  # E列


# ---------------------------------------------------------------------------
# 実データ（data.csv）の読み込み
# ---------------------------------------------------------------------------

@unittest.skipUnless(os.path.isfile("data.csv"), "data.csv がありません")
class TestLoadRealCsv(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.main_df, cls.phenom_df, cls.cols, cls.mapping = wv.load_weather_data(
            "data.csv", verbose=False)

    def test_all_rows_are_parsed(self):
        # 2160行すべてが日時として解釈できること（旧版はすべてNaTで消えていた）
        self.assertEqual(len(self.main_df), 2160)
        self.assertEqual(self.main_df["datetime"].min(), pd.Timestamp("2024-06-28 01:00"))
        # 最終行は「2024年9月25日24時」＝ 9/26 0時
        self.assertEqual(self.main_df["datetime"].max(), pd.Timestamp("2024-09-26 00:00"))

    def test_locations_detected(self):
        self.assertEqual(list(self.mapping.values()),
                         ["伊南川合流点", "柴倉橋", "堅盤橋", "蒲生水道橋"])

    def test_measures_detected(self):
        self.assertGreater(self.main_df["気温(℃)"].notna().sum(), 2000)
        self.assertGreater(self.main_df["相対湿度(％)"].notna().sum(), 2000)
        self.assertGreater(self.main_df[wv.WIND_DIR_LABEL].notna().sum(), 1900)
        # 降水量の列は無いのでNaN列になる（例外にはならない）
        self.assertEqual(self.main_df["降水量(mm)"].notna().sum(), 0)


# ---------------------------------------------------------------------------
# ダミーデータ検出
# ---------------------------------------------------------------------------

class TestDummyBlocks(unittest.TestCase):
    def _df(self, codes, start="2024-06-01"):
        times = pd.date_range(start, periods=len(codes), freq="h")
        return pd.DataFrame({"datetime": times, "X": codes})

    def test_block_without_slash_is_suspicious(self):
        df = self._df(["1"] * 100)
        blocks = wv.find_dummy_blocks(df, "X")
        self.assertEqual(len(blocks), 1)
        self.assertEqual(blocks[0][3], 100)

    def test_normal_block_is_not_suspicious(self):
        codes = (["/"] * 9 + ["2"]) * 20
        self.assertEqual(wv.find_dummy_blocks(self._df(codes), "X"), [])

    def test_short_block_is_ignored(self):
        self.assertEqual(wv.find_dummy_blocks(self._df(["1"] * 10), "X"), [])


# ---------------------------------------------------------------------------
# 予報取得（Open-Meteo）
# ---------------------------------------------------------------------------

def _fake_forecast_payload(n_hours=48, start="2026-08-17 00:00"):
    times = pd.date_range(start, periods=n_hours, freq="h")
    return {
        "hourly": {
            "time": [t.strftime("%Y-%m-%dT%H:%M") for t in times],
            "temperature_2m": [20.0 + (i % 12) for i in range(n_hours)],
            "precipitation": [0.0] * n_hours,
            "relative_humidity_2m": [95 - (i % 12) for i in range(n_hours)],
            "dew_point_2m": [18.0 + (i % 6) for i in range(n_hours)],
            "wind_speed_10m": [0.5 + (i % 3) * 0.4 for i in range(n_hours)],
            "wind_direction_10m": [(i * 15) % 360 for i in range(n_hours)],
        }
    }


class _FakeResponse:
    def __init__(self, payload):
        self._payload = payload

    def raise_for_status(self):
        pass

    def json(self):
        return self._payload


class TestFetchForecast(unittest.TestCase):
    def test_requests_wind_speed_in_meters_per_second(self):
        """Open-Meteoの既定はkm/h。m/sを明示していないと学習データと単位が食い違う。"""
        captured = {}

        def fake_get(url, params=None, timeout=None):
            captured["url"] = url
            captured["params"] = params
            return _FakeResponse(_fake_forecast_payload())

        with mock.patch.object(wv.requests, "get", fake_get):
            df = wv.fetch_forecast_range(lat=37.3, lon=139.3, forecast_days=2)

        self.assertEqual(captured["params"]["wind_speed_unit"], "ms")
        self.assertEqual(captured["params"]["timezone"], "Asia/Tokyo")
        self.assertIn("wind_direction_10m", captured["params"]["hourly"])
        self.assertEqual(len(df), 48)
        self.assertIn(wv.WIND_DIR_LABEL, df.columns)

    def test_missing_values_are_filled(self):
        payload = _fake_forecast_payload(24)
        payload["hourly"]["temperature_2m"][5] = None
        payload["hourly"]["precipitation"][6] = None

        with mock.patch.object(wv.requests, "get", lambda *a, **k: _FakeResponse(payload)):
            df = wv.fetch_forecast_range()

        self.assertEqual(len(df), 24)
        self.assertFalse(df["気温(℃)"].isna().any())
        self.assertEqual(df["降水量(mm)"].iloc[6], 0.0)


# ---------------------------------------------------------------------------
# 学習・予測・グラフ出力（エンドツーエンド）
# ---------------------------------------------------------------------------

def _synthetic_observations(n_days=60):
    """霧が『湿度が高く気温露点差が小さい早朝』に出る、分かりやすい人工データ。"""
    times = pd.date_range("2026-05-01", periods=n_days * 24, freq="h")
    rng = np.random.default_rng(0)
    hours = times.hour.to_numpy()
    temp = 18 + 8 * np.sin(2 * np.pi * (hours - 6) / 24) + rng.normal(0, 0.5, len(times))
    dew = temp - np.where((hours >= 3) & (hours <= 7), 0.2, 4.0) - rng.normal(0, 0.3, len(times))
    humid = np.clip(100 - (temp - dew) * 6, 30, 100)
    wind = np.abs(rng.normal(1.0, 0.5, len(times)))
    precip = np.zeros(len(times))

    main_df = pd.DataFrame({
        "datetime": times, "気温(℃)": temp, "降水量(mm)": precip, "風速(m/s)": wind,
        "露点温度(℃)": dew, "相対湿度(％)": humid,
    })
    fog = (temp - dew < 1.0) & (hours >= 3) & (hours <= 7)
    codes = np.where(fog, "2", "/")
    phenom_df = pd.DataFrame({"datetime": times, "AC": codes, "AD": codes})
    return main_df, phenom_df, ["AC", "AD"], {"AC": "地点甲", "AD": "地点乙"}


class TestTrainingAndPlots(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.main_df, cls.phenom_df, cls.cols, cls.mapping = _synthetic_observations()
        cls.tmpdir = tempfile.mkdtemp(prefix="wv_test_")

    @classmethod
    def tearDownClass(cls):
        shutil.rmtree(cls.tmpdir, ignore_errors=True)

    def test_train_and_forecast_pipeline(self):
        models = wv.train_location_models(self.main_df, self.phenom_df, self.cols, self.mapping)
        self.assertEqual(set(models), {"AC", "AD"})
        self.assertGreater(models["AC"]["accuracy"], 0.8)
        self.assertGreater(models["AC"]["fog_f1"], 0.5)

        with mock.patch.object(wv.requests, "get",
                               lambda *a, **k: _FakeResponse(_fake_forecast_payload(72))):
            forecast = wv.fetch_forecast_range()

        pred_df, prob_df = wv.build_location_forecast_codes(models, forecast)
        self.assertEqual(len(pred_df), len(forecast))
        self.assertTrue(pred_df["AC"].isin(range(0, 11)).all())
        self.assertTrue(((prob_df["AC"] >= 0) & (prob_df["AC"] <= 1)).all())

        p4 = wv.plot_single_location_forecast(
            self.main_df, self.phenom_df, "AC", self.mapping, forecast, pred_df["AC"],
            "只見", self.tmpdir, prob_series=prob_df["AC"])
        self.assertTrue(os.path.isfile(p4))

        p5 = wv.plot_all_location_summary(pred_df, models, self.mapping, "只見",
                                          self.tmpdir, prob_df=prob_df)
        self.assertTrue(os.path.isfile(p5))

        p6 = wv.export_prediction_csv(pred_df, prob_df, models, self.mapping, "只見", self.tmpdir)
        out = pd.read_csv(p6)
        self.assertIn("地点甲_予測コード", out.columns)
        self.assertIn("地点甲_霧確率(%)", out.columns)
        self.assertEqual(len(out), len(forecast))

    def test_monthly_graphs(self):
        made = wv.plot_combo_by_month(self.main_df.head(200), self.phenom_df.head(200),
                                      self.cols, self.mapping, "只見", self.tmpdir)
        self.assertTrue(made)
        for p in made:
            self.assertTrue(os.path.isfile(p))

    def test_forecast_with_missing_feature_does_not_crash(self):
        """予報側に欠測が残っていてもpredictが例外にならないこと。"""
        models = wv.train_location_models(self.main_df, self.phenom_df, ["AC"], {"AC": "地点甲"})
        with mock.patch.object(wv.requests, "get",
                               lambda *a, **k: _FakeResponse(_fake_forecast_payload(24))):
            forecast = wv.fetch_forecast_range()
        forecast.loc[0, "降水量(mm)"] = np.nan  # 降水量の欠測は0扱いになる
        pred_df, prob_df = wv.build_location_forecast_codes(models, forecast)
        self.assertEqual(len(pred_df), 24)


# ---------------------------------------------------------------------------
# 引数の解釈
# ---------------------------------------------------------------------------

class TestArgParsing(unittest.TestCase):
    def test_input_and_output(self):
        args = wv._parse_args(["データ.xlsx", "出力先"])
        self.assertEqual(args.input, "データ.xlsx")
        self.assertEqual(args.output, "出力先")

    def test_reversed_order(self):
        args = wv._parse_args(["出力先", "データ.xlsx"])
        self.assertEqual(args.input, "データ.xlsx")
        self.assertEqual(args.output, "出力先")

    def test_defaults(self):
        args = wv._parse_args([])
        self.assertEqual(args.input, wv.DEFAULT_INPUT_FILE)
        self.assertEqual(args.output, wv.DEFAULT_OUTPUT_DIR)
        self.assertFalse(args.no_forecast)

    def test_flags(self):
        args = wv._parse_args(["a.csv", "--no-forecast", "--lat", "37.1", "--lon", "139.2",
                               "--forecast-days", "7", "--layout", "fixed"])
        self.assertTrue(args.no_forecast)
        self.assertEqual(args.lat, 37.1)
        self.assertEqual(args.forecast_days, 7)
        self.assertEqual(args.layout, "fixed")

    def test_notebook_kernel_args_are_ignored(self):
        """Jupyterのカーネル引数を出力フォルダと誤認しないこと（旧版の不具合）。"""
        with mock.patch.object(wv.sys, "argv",
                               ["ipykernel_launcher.py", "-f", "/tmp/kernel-1234.json"]), \
             mock.patch.object(wv, "_in_notebook", lambda: True):
            args = wv._parse_args()
        self.assertEqual(args.output, wv.DEFAULT_OUTPUT_DIR)


# ---------------------------------------------------------------------------
# main() のエンドツーエンド（Excel入力・予報はダミー）
# ---------------------------------------------------------------------------

class TestMainEndToEnd(unittest.TestCase):
    def test_main_generates_files_from_xlsx(self):
        from openpyxl import Workbook

        main_df, phenom_df, cols, mapping = _synthetic_observations(n_days=20)
        tmpdir = tempfile.mkdtemp(prefix="wv_e2e_")
        try:
            xlsx_path = os.path.join(tmpdir, "只見_テストデータ.xlsx")
            wb = Workbook()
            ws = wb.active
            ws.append([None] * 6 + ["地点甲", "地点乙"])
            ws.append(["年月日時", "気温(℃)", "降水量(mm)", "風速(m/s)", "風向",
                       "露点温度(℃)", "相対湿度(％)"])
            for i in range(len(main_df)):
                t = main_df["datetime"].iloc[i]
                hour = t.hour if t.hour != 0 else 24
                stamp = (t - pd.Timedelta(hours=1)) if t.hour == 0 else t
                ws.append([f"{stamp.year}年{stamp.month}月{stamp.day}日{hour}時",
                           round(float(main_df['気温(℃)'].iloc[i]), 1),
                           0.0, round(float(main_df['風速(m/s)'].iloc[i]), 1), "南",
                           round(float(main_df['露点温度(℃)'].iloc[i]), 1),
                           int(main_df['相対湿度(％)'].iloc[i]),
                           phenom_df["AC"].iloc[i], phenom_df["AD"].iloc[i]])
            wb.save(xlsx_path)

            out_dir = os.path.join(tmpdir, "out")
            with mock.patch.object(wv.requests, "get",
                                   lambda *a, **k: _FakeResponse(_fake_forecast_payload(48))):
                rc = wv.main([xlsx_path, out_dir])

            self.assertEqual(rc, 0)
            files = os.listdir(out_dir)
            self.assertTrue(any(f.startswith("只見_①") for f in files), files)
            self.assertTrue(any(f.startswith("只見_②") for f in files), files)
            self.assertTrue(any("④" in f for f in files), files)
            self.assertTrue(any("⑤" in f for f in files), files)
            self.assertTrue(any(f.endswith(".csv") for f in files), files)
        finally:
            shutil.rmtree(tmpdir, ignore_errors=True)

    def test_main_returns_error_for_missing_file(self):
        self.assertEqual(wv.main(["存在しないファイル.xlsx", "/tmp/なんとか"]), 1)


if __name__ == "__main__":
    unittest.main()
