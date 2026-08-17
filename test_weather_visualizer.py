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

import matplotlib
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
# 日本語フォント（グラフの文字化け対策）
# ---------------------------------------------------------------------------

class _FakeFontEntry:
    def __init__(self, name, fname):
        self.name = name
        self.fname = fname


class TestJapaneseFont(unittest.TestCase):
    def setUp(self):
        self._rc = dict(matplotlib.rcParams)

    def tearDown(self):
        matplotlib.rcParams.update(self._rc)

    def test_dejavu_is_detected_as_unusable_for_japanese(self):
        """名前だけでなく実際のグリフで判定できること（豆腐文字の検出）。"""
        import matplotlib.font_manager as fm
        dejavu = fm.findfont(fm.FontProperties(family="DejaVu Sans"))
        self.assertFalse(wv._font_file_supports_japanese(dejavu))

    def test_placeholder_fonts_are_rejected(self):
        """Last Resort は全文字のグリフを持つが中身は□なので選んではいけない。"""
        self.assertTrue(wv._is_placeholder_font("Last Resort High-Efficiency", "/x/LastResortHE.ttf"))
        self.assertTrue(wv._is_placeholder_font(
            "DejaVu Sans", "/usr/lib/python3/matplotlib/mpl-data/fonts/ttf/DejaVuSans.ttf"))
        self.assertFalse(wv._is_placeholder_font("IPAexGothic", "/usr/share/fonts/ipaexg.ttf"))

    def test_find_cjk_font_skips_placeholder_fonts(self):
        import matplotlib.font_manager as fm
        last_resort = os.path.join(matplotlib.get_data_path(), "fonts", "ttf",
                                   "LastResortHE-Regular.ttf")
        if not os.path.isfile(last_resort):
            self.skipTest("Last Resort フォントが同梱されていません")
        # Last Resort だけを登録した状態では「日本語フォントなし」と判定されるべき
        with mock.patch.object(wv.fm.fontManager, "ttflist",
                               [_FakeFontEntry("Last Resort High-Efficiency", last_resort)]):
            name, _ = wv._find_cjk_font()
        self.assertIsNone(name)

    def test_ttc_font_is_verified(self):
        """.ttc（複数書体をまとめたファイル）でもグリフ判定ができること。"""
        import matplotlib.font_manager as fm
        ttcs = [f.fname for f in fm.fontManager.ttflist
                if str(f.fname).lower().endswith((".ttc", ".otc"))]
        if not ttcs:
            self.skipTest(".ttc フォントがない環境です")
        results = [wv._font_file_supports_japanese(p) for p in ttcs[:10]]
        self.assertTrue(any(r is not None for r in results),
                        "すべて判定不能＝.ttcの読み込みに失敗している")

    def test_apply_font_keeps_other_fonts_as_fallback(self):
        wv._apply_font("TestFont")
        self.assertEqual(matplotlib.rcParams["font.family"], ["sans-serif"])
        self.assertEqual(matplotlib.rcParams["font.sans-serif"][0], "TestFont")
        self.assertIn("DejaVu Sans", matplotlib.rcParams["font.sans-serif"])
        self.assertFalse(matplotlib.rcParams["axes.unicode_minus"])

    def test_apply_font_does_not_duplicate(self):
        wv._apply_font("TestFont")
        wv._apply_font("TestFont")
        self.assertEqual(matplotlib.rcParams["font.sans-serif"].count("TestFont"), 1)

    def test_find_cjk_font_ignores_fonts_without_japanese_glyphs(self):
        """名前に日本語フォントらしさが無くても、グリフがあれば拾えること。"""
        import matplotlib.font_manager as fm
        dejavu = fm.findfont(fm.FontProperties(family="DejaVu Sans"))
        real = fm.findfont(fm.FontProperties(family=fm.FontProperties().get_family()))
        fake_list = [_FakeFontEntry("DejaVu Sans", str(dejavu)),
                     _FakeFontEntry("謎のフォント", str(real))]
        with mock.patch.object(wv.fm.fontManager, "ttflist", fake_list):
            name, path = wv._find_cjk_font()
        # この環境に日本語フォントがあれば拾え、無ければ何も返さない
        if name is not None:
            self.assertNotEqual(name, "DejaVu Sans")

    def test_install_is_attempted_when_no_font_found(self):
        """フォントが無い環境（Colab等）でインストールが試みられること。"""
        calls = []

        def fake_run(cmd, timeout):
            calls.append(cmd)
            return False, "E: Unable to locate package"

        with mock.patch.object(wv, "_find_cjk_font", lambda *a, **k: (None, None)), \
             mock.patch.object(wv, "_register_font_files", lambda *a, **k: 0), \
             mock.patch.object(wv, "_run_command", fake_run), \
             mock.patch.object(wv.sys, "platform", "linux"):
            name, path = wv._install_japanese_font(verbose=False)

        self.assertIsNone(name)
        joined = [" ".join(c) for c in calls]
        self.assertTrue(any("fonts-ipafont-gothic" in c for c in joined), joined)
        self.assertTrue(any("fonts-noto-cjk" in c for c in joined), joined)
        self.assertTrue(any("japanize-matplotlib" in c for c in joined), joined)
        self.assertTrue(any("apt-get update" in c for c in joined), joined)

    def test_install_is_skipped_when_disabled(self):
        with mock.patch.object(wv, "_find_cjk_font", lambda *a, **k: (None, None)), \
             mock.patch.object(wv, "_register_font_files", lambda *a, **k: 0), \
             mock.patch.object(wv, "_install_japanese_font") as installer:
            wv.setup_japanese_font(verbose=False, allow_install=False)
        installer.assert_not_called()

    def test_setup_reports_failure_instead_of_pretending(self):
        """フォントを設定できない場合はNoneを返すこと（成功したふりをしない）。"""
        with mock.patch.object(wv, "_find_cjk_font", lambda *a, **k: (None, None)), \
             mock.patch.object(wv, "_register_font_files", lambda *a, **k: 0), \
             mock.patch.object(wv, "verify_japanese_font",
                               lambda: (False, "DejaVu Sans", "/x/DejaVuSans.ttf")):
            result = wv.setup_japanese_font(verbose=False, allow_install=False)
        self.assertIsNone(result)

    def test_font_check_image_is_created(self):
        tmpdir = tempfile.mkdtemp(prefix="wv_font_")
        try:
            path = wv.plot_font_check(tmpdir)
            self.assertTrue(os.path.isfile(path))
        finally:
            shutil.rmtree(tmpdir, ignore_errors=True)


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
# 本番と同じ形式のファイル（test.xlsx）の読み込み
# ---------------------------------------------------------------------------

SAMPLE_XLSX = "test.xlsx"


@unittest.skipUnless(os.path.isfile(SAMPLE_XLSX), f"{SAMPLE_XLSX} がありません")
class TestLoadSampleXlsx(unittest.TestCase):
    """見出し行の下に小見出し行（風向・品質情報）が続く、本番と同じ形式。"""

    @classmethod
    def setUpClass(cls):
        cls.main_df, cls.phenom_df, cls.cols, cls.mapping = wv.load_weather_data(
            SAMPLE_XLSX, verbose=False)

    def test_all_data_rows_are_parsed(self):
        # 小見出し行を飛ばし、データ行だけを読むこと（以前は日時なしの2行が混ざっていた）
        self.assertEqual(len(self.main_df), 17520)
        self.assertEqual(self.main_df["datetime"].min(), pd.Timestamp("2024-06-01 01:00"))
        self.assertEqual(self.main_df["datetime"].max(), pd.Timestamp("2026-06-01 00:00"))

    def test_locations_detected(self):
        self.assertEqual(len(self.mapping), 32)
        self.assertEqual(list(self.mapping.values())[:4],
                         ["伊奈川合流地点", "柴倉橋", "堅盤橋", "蒲生水道橋"])
        # 品質情報・均質番号・視程などの補助列が地点として混ざっていないこと
        self.assertFalse([n for n in self.mapping.values()
                          if any(k in n for k in ("品質", "均質", "視程", "水温"))])

    def test_measures_detected(self):
        for label in ("気温(℃)", "降水量(mm)", "風速(m/s)", "露点温度(℃)", "相対湿度(％)"):
            self.assertGreater(self.main_df[label].notna().sum(), 17000, label)

    def test_wind_direction_is_detected_below_header(self):
        # 風向は見出し行の【下】の行に書かれている（以前は見落として特徴量から欠けていた）
        self.assertIn(wv.WIND_DIR_LABEL, self.main_df.columns)
        self.assertGreater(self.main_df[wv.WIND_DIR_LABEL].notna().sum(), 16000)


# ---------------------------------------------------------------------------
# 地点ではない列（水温など）を学習に混ぜない
# ---------------------------------------------------------------------------

def _write_production_style_xlsx(path, extra_headers=(), extra_values=()):
    """本番（test.xlsx）と同じ形の小さなExcelを作る。

    extra_headers / extra_values に列を足すと、地点列の右に別の列がある状態を作れる。
    """
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.append(["ダウンロードした時刻:2026/06/30 17:09:05"])
    ws.append([])
    ws.append(["", "只見", "只見", "只見", "只見", "只見", "只見", "",
               "A橋", "B橋", *extra_headers])
    ws.append(["年月日時", "気温(°C)", "降水量(mm)", "風速(m/s)", "風速(m/s)",
               "露点温度(°C)", "相対湿度(%)"])
    ws.append(["", "", "", "", "風向"])                      # 小見出し行1
    ws.append(["", "品質情報", "", "品質情報", "", "", ""])   # 小見出し行2
    for i in range(300):
        t = datetime(2024, 6, 1) + pd.Timedelta(hours=i)
        ws.append([t, 20.0, 0.0, 1.2, "南南西", 18.0, 90,
                   None, "/" if i % 5 else 2, "/" if i % 7 else 1,
                   *[v(i) for v in extra_values]])
    wb.save(path)


class TestNonLocationColumns(unittest.TestCase):
    def setUp(self):
        self.tmpdir = tempfile.mkdtemp(prefix="wv_cols_")
        self.path = os.path.join(self.tmpdir, "sample.xlsx")

    def tearDown(self):
        shutil.rmtree(self.tmpdir, ignore_errors=True)

    def test_water_temperature_column_is_not_a_location(self):
        """「水温」の列を地点として学習しないこと（8.2→コード8 と誤読させない）。"""
        _write_production_style_xlsx(
            self.path,
            extra_headers=["水温", "水温(℃)"],
            extra_values=[lambda i: round(8.0 + (i % 40) * 0.05, 1),
                          lambda i: round(9.0 + (i % 40) * 0.05, 1)],
        )
        _, _, cols, mapping = wv.load_weather_data(self.path, verbose=False)
        self.assertEqual(list(mapping.values()), ["A橋", "B橋"])
        self.assertEqual(len(cols), 2)

    def test_unnamed_measure_column_is_rejected_by_its_values(self):
        """見出しからは判別できない列でも、中身が現象コードでなければ地点にしない。"""
        _write_production_style_xlsx(
            self.path,
            extra_headers=["取水口"],                       # 地点名に見える見出し
            extra_values=[lambda i: round(12.3 + i * 0.01, 2)],
        )
        grid = wv.read_grid(self.path)
        info = wv.detect_layout(grid, wv.find_header_row(grid))
        names = [name for _, _, name in info["phenom"]]
        self.assertEqual(names, ["A橋", "B橋"])
        excluded = [(letter, name) for letter, name, _ in info["excluded"]]
        self.assertIn("取水口", [name for _, name in excluded])

    def test_location_column_without_data_is_kept(self):
        """まだ記入されていない地点の列は、これから使うので残すこと。"""
        _write_production_style_xlsx(self.path, extra_headers=["C橋"],
                                     extra_values=[lambda i: None])
        _, _, _, mapping = wv.load_weather_data(self.path, verbose=False)
        self.assertEqual(list(mapping.values()), ["A橋", "B橋", "C橋"])

    def test_phenomena_value_check(self):
        for good in ("/", "2", "10", "1 8", 0, 7, 10.0):
            self.assertTrue(wv._is_phenomena_value(good), good)
        for bad in ("8.2", "15", "12.5", 8.2, 15, "晴", -1):
            self.assertFalse(wv._is_phenomena_value(bad), bad)

    def test_measure_header_detects_units(self):
        self.assertTrue(wv._is_measure_header(["水温(℃)"]))
        self.assertTrue(wv._is_measure_header(["積算値(mm)"]))
        # 括弧付きの地点名は地点のまま（単位ではない）
        self.assertFalse(wv._is_measure_header(["(川口橋)"]))
        self.assertFalse(wv._is_measure_header(["蒲生水道橋"]))


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

    def test_optional_measures_are_unit_converted(self):
        """日照時間(秒→時間)・日射量(W/m2→MJ/m2)・視程(m→km)の単位換算を確認する。"""
        payload = _fake_forecast_payload(6)
        n = len(payload["hourly"]["time"])
        payload["hourly"]["sunshine_duration"] = [1800.0] * n     # 1800秒=0.5時間
        payload["hourly"]["shortwave_radiation"] = [100.0] * n    # 100W/m2 -> 0.36MJ/m2
        payload["hourly"]["visibility"] = [5000.0] * n            # 5000m -> 5km
        payload["hourly"]["surface_pressure"] = [1013.0] * n

        with mock.patch.object(wv.requests, "get", lambda *a, **k: _FakeResponse(payload)):
            df = wv.fetch_forecast_range()

        self.assertTrue(np.allclose(df["日照時間(時間)"], 0.5))
        self.assertTrue(np.allclose(df["日射量(MJ/m2)"], 0.36))
        self.assertTrue(np.allclose(df["視程(km)"], 5.0))
        self.assertTrue(np.allclose(df["気圧(hPa)"], 1013.0))

    def test_fetch_forecast_multi_maps_by_location_name(self):
        """複数地点の緯度経度をまとめて渡すと、地点名ごとのDataFrameが返ること。"""
        payloads = [_fake_forecast_payload(6), _fake_forecast_payload(6)]
        payloads[1]["hourly"]["temperature_2m"] = [99.0] * len(payloads[1]["hourly"]["time"])

        class _MultiResponse:
            def __init__(self, data):
                self._data = data

            def raise_for_status(self):
                pass

            def json(self):
                return self._data

        with mock.patch.object(wv.requests, "get", lambda *a, **k: _MultiResponse(payloads)):
            result = wv.fetch_forecast_multi({"地点A": (37.0, 139.0), "地点B": (37.1, 139.1)})

        self.assertEqual(set(result.keys()), {"地点A", "地点B"})
        self.assertTrue((result["地点B"]["気温(℃)"] == 99.0).all())
        self.assertFalse((result["地点A"]["気温(℃)"] == 99.0).all())


class TestLocationCoordinates(unittest.TestCase):
    def test_load_location_coordinates(self):
        tmpdir = tempfile.mkdtemp(prefix="wv_coords_")
        try:
            path = os.path.join(tmpdir, "coords.csv")
            with open(path, "w", encoding="utf-8") as f:
                f.write("地点名,緯度,経度\n只見,37.3,139.3\n田子倉,37.4,139.2\n")
            coords = wv.load_location_coordinates(path)
            self.assertEqual(coords, {"只見": (37.3, 139.3), "田子倉": (37.4, 139.2)})
        finally:
            shutil.rmtree(tmpdir, ignore_errors=True)

    def test_missing_lat_lon_columns_raises(self):
        tmpdir = tempfile.mkdtemp(prefix="wv_coords_")
        try:
            path = os.path.join(tmpdir, "bad.csv")
            with open(path, "w", encoding="utf-8") as f:
                f.write("地点名,x,y\n只見,1,2\n")
            with self.assertRaises(ValueError):
                wv.load_location_coordinates(path)
        finally:
            shutil.rmtree(tmpdir, ignore_errors=True)

    def test_fetch_per_location_forecasts_matches_by_name_and_falls_back(self):
        """座標ファイルにある地点だけ専用の予報を使い、無い地点は既定にフォールバックすること。"""
        tmpdir = tempfile.mkdtemp(prefix="wv_coords_")
        try:
            path = os.path.join(tmpdir, "coords.csv")
            with open(path, "w", encoding="utf-8") as f:
                f.write("地点名,緯度,経度\n地点甲,37.0,139.0\n")

            default_forecast = wv._forecast_payload_to_df(_fake_forecast_payload(6)["hourly"])
            own_payload = _fake_forecast_payload(6)
            own_payload["hourly"]["temperature_2m"] = [55.0] * len(own_payload["hourly"]["time"])

            class _MultiResponse:
                def raise_for_status(self):
                    pass

                def json(self):
                    return [own_payload]

            with mock.patch.object(wv.requests, "get", lambda *a, **k: _MultiResponse()):
                forecast_by_col = wv._fetch_per_location_forecasts(
                    path, ["AC", "AD"], {"AC": "地点甲", "AD": "地点乙"}, 2, default_forecast)

            self.assertIn("AC", forecast_by_col)
            self.assertNotIn("AD", forecast_by_col)  # 座標ファイルに無い地点は既定にフォールバック
            self.assertTrue((forecast_by_col["AC"]["気温(℃)"] == 55.0).all())
        finally:
            shutil.rmtree(tmpdir, ignore_errors=True)


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


class _StubModel:
    """predict_proba が固定の確率を返すだけのダミーモデル（ブレンド計算の検証用）。"""

    def __init__(self, classes, proba):
        self.classes_ = np.array(classes)
        self.proba = np.array(proba, dtype=float)

    def predict_proba(self, X):
        return self.proba


class TestHybridFogClassifier(unittest.TestCase):
    def setUp(self):
        # クラスは 0（現象なし）と 2（川霧＝霧コード）の2つ
        self.rf = _StubModel([0, 2], [[0.8, 0.2]])
        self.xgb = _StubModel([0, 2], [[0.2, 0.8]])
        self.X = np.zeros((1, 3))

    def test_blend_is_weighted_average(self):
        half = wv.HybridFogClassifier(self.rf, self.xgb, weight=0.5)
        np.testing.assert_allclose(half.predict_proba(self.X), [[0.5, 0.5]])
        rf_heavy = wv.HybridFogClassifier(self.rf, self.xgb, weight=0.75)
        np.testing.assert_allclose(rf_heavy.predict_proba(self.X), [[0.65, 0.35]])

    def test_weight_endpoints_are_single_models(self):
        np.testing.assert_allclose(
            wv.HybridFogClassifier(self.rf, self.xgb, weight=1.0).predict_proba(self.X),
            self.rf.proba)
        np.testing.assert_allclose(
            wv.HybridFogClassifier(self.rf, self.xgb, weight=0.0).predict_proba(self.X),
            self.xgb.proba)

    def test_predict_returns_original_phenomenon_codes(self):
        # XGB寄りにすると霧コード2が選ばれる
        self.assertEqual(wv.HybridFogClassifier(self.rf, self.xgb, weight=0.0).predict(self.X)[0], 2)
        self.assertEqual(wv.HybridFogClassifier(self.rf, self.xgb, weight=1.0).predict(self.X)[0], 0)

    def test_single_model_types(self):
        rf_only = wv.HybridFogClassifier(self.rf)
        self.assertEqual(rf_only.model_type, "rf")
        self.assertEqual(rf_only.weight, 1.0)
        xgb_only = wv.HybridFogClassifier(None, self.xgb)
        self.assertEqual(xgb_only.model_type, "xgb")
        self.assertEqual(xgb_only.weight, 0.0)
        with self.assertRaises(ValueError):
            wv.HybridFogClassifier(None, None)

    def test_classes_are_aligned_before_blending(self):
        """片方に無いクラスがあっても、列がずれずに合成されること。"""
        rf = _StubModel([0, 1, 2], [[0.5, 0.3, 0.2]])
        xgb = _StubModel([0, 2], [[0.4, 0.6]])  # クラス1を知らない
        blended = wv.HybridFogClassifier(rf, xgb, weight=0.5).predict_proba(self.X)
        np.testing.assert_allclose(blended, [[0.45, 0.15, 0.40]])

    def test_compute_fog_probability_works_with_hybrid(self):
        hybrid = wv.HybridFogClassifier(self.rf, self.xgb, weight=0.5)
        # クラス2は霧コードなので、霧確率は0.5になる
        np.testing.assert_allclose(wv.compute_fog_probability(hybrid, self.X), [0.5])


class TestModelKindResolution(unittest.TestCase):
    def test_invalid_kind_raises(self):
        with self.assertRaises(ValueError):
            wv._resolve_model_kind("lightgbm")

    def test_falls_back_to_rf_without_xgboost(self):
        with mock.patch.object(wv, "_HAS_XGBOOST", False):
            self.assertEqual(wv._resolve_model_kind("hybrid"), "rf")
            self.assertEqual(wv._resolve_model_kind("xgb"), "rf")
            self.assertEqual(wv._resolve_model_kind("rf"), "rf")

    def test_default_is_hybrid(self):
        self.assertEqual(wv.DEFAULT_MODEL_KIND, "hybrid")
        with mock.patch.object(wv, "_HAS_XGBOOST", True):
            self.assertEqual(wv._resolve_model_kind(None), "hybrid")


@unittest.skipUnless(wv._HAS_XGBOOST, "xgboost が導入されていない環境ではスキップ")
class TestXGBLabelSafeClassifier(unittest.TestCase):
    def test_handles_non_contiguous_labels(self):
        """現象コードが {0, 2, 7} のように歯抜けでも学習・予測できること。"""
        rng = np.random.default_rng(0)
        X = pd.DataFrame({"a": rng.normal(size=120), "b": rng.normal(size=120)})
        y = pd.Series(np.tile([0, 2, 7], 40))
        X["a"] = y.map({0: 0.0, 2: 5.0, 7: 10.0}) + rng.normal(0, 0.1, 120)  # 分離しやすくする

        clf = wv._XGBLabelSafeClassifier(n_estimators=20, max_depth=3).fit(X, y)
        self.assertEqual(list(clf.classes_), [0, 2, 7])
        pred = clf.predict(X)
        self.assertTrue(set(np.unique(pred)).issubset({0, 2, 7}))
        self.assertGreater((pred == y.to_numpy()).mean(), 0.9)
        self.assertEqual(clf.predict_proba(X).shape, (120, 3))


class TestDisplayWidth(unittest.TestCase):
    """表の桁揃え（`_disp_width`/`_pad`）。全角混在でも表示幅で揃うことを確認する。"""

    def test_disp_width_counts_full_width_as_two(self):
        self.assertEqual(wv._disp_width("木賊"), 4)
        self.assertEqual(wv._disp_width("AC12"), 4)
        self.assertEqual(wv._disp_width("伊奈川合流地点"), 14)

    def test_pad_aligns_by_display_width_not_char_count(self):
        short = wv._pad("木賊", 14)
        long_name = wv._pad("伊奈川合流地点", 14)
        self.assertEqual(wv._disp_width(short), 14)
        self.assertEqual(wv._disp_width(long_name), 14)

    def test_pad_right_align(self):
        self.assertEqual(wv._pad("1", 4, ">"), "   1")


class TestFogThreshold(unittest.TestCase):
    def test_threshold_defaults_when_no_fog_in_validation(self):
        model = _StubModel([0, 2], [[0.9, 0.1]] * 5)
        y_val = pd.Series([0, 0, 0, 0, 0])
        self.assertEqual(wv._choose_fog_threshold(model, np.zeros((5, 1)), y_val),
                         wv.DEFAULT_FOG_THRESHOLD)

    def test_threshold_picks_separating_value(self):
        # 前半は霧なし（確率低）、後半は霧あり（確率高）の分かりやすいケース
        proba = np.array([[0.9, 0.1]] * 5 + [[0.1, 0.9]] * 5)
        model = _StubModel([0, 2], proba)
        y_val = pd.Series([0] * 5 + [2] * 5)
        t = wv._choose_fog_threshold(model, np.zeros((10, 1)), y_val)
        self.assertGreater(t, 0.1)
        self.assertLess(t, 0.9)


class TestLagAndSunFeatures(unittest.TestCase):
    def test_lag_series_uses_exact_time_lookup_not_position(self):
        """行が飛んでいる（欠測）場合、位置ベースではなく時刻ベースで過去の値を探すこと。"""
        times = pd.to_datetime(["2026-01-01 00:00", "2026-01-01 01:00",
                                "2026-01-01 03:00"])  # 2時が欠けている
        df = pd.DataFrame({"datetime": times, "気温(℃)": [10.0, 12.0, 20.0]})
        lagged = wv._lag_series(df, "気温(℃)", 1)
        self.assertEqual(lagged.iloc[1], 10.0)   # 1時の1時間前=0時は存在する
        self.assertTrue(np.isnan(lagged.iloc[2]))  # 3時の1時間前=2時は存在しない

    def test_add_lag_features_fills_gaps_after_prepare_features(self):
        main_df, _, _, _ = _synthetic_observations(n_days=3)
        out, feature_names = wv.prepare_features(main_df)
        for name in wv.LAG_FEATURE_NAMES:
            self.assertIn(name, feature_names)
            self.assertFalse(out[name].isna().any(), name)

    def test_sunrise_feature_is_zero_during_broad_daytime_and_bounded(self):
        main_df, _, _, _ = _synthetic_observations(n_days=2)
        out = wv._add_sunrise_feature(main_df)
        noon_mask = main_df["datetime"].dt.hour == 12
        self.assertTrue((out.loc[noon_mask, wv.SUN_FEATURE_NAME] == 0.0).all())
        self.assertTrue((out[wv.SUN_FEATURE_NAME] >= 0).all())
        self.assertTrue((out[wv.SUN_FEATURE_NAME] <= 12).all())


class TestOptionalMeasureFeatures(unittest.TestCase):
    def test_low_coverage_optional_measure_is_excluded(self):
        main_df, _, _, _ = _synthetic_observations(n_days=5)
        main_df["気圧(hPa)"] = np.nan
        main_df.loc[main_df.index[:2], "気圧(hPa)"] = 1013.0  # 有効値がごくわずか
        out, feature_names = wv.prepare_features(main_df)
        self.assertNotIn("気圧(hPa)", feature_names)

    def test_high_coverage_optional_measure_is_included_and_filled(self):
        main_df, _, _, _ = _synthetic_observations(n_days=5)
        main_df["気圧(hPa)"] = 1013.0
        main_df.loc[main_df.index[0], "気圧(hPa)"] = np.nan
        out, feature_names = wv.prepare_features(main_df)
        self.assertIn("気圧(hPa)", feature_names)
        self.assertFalse(out["気圧(hPa)"].isna().any())


class TestPooledFallbackModel(unittest.TestCase):
    def test_data_poor_location_is_rescued_by_pooled_model(self):
        """個別条件を満たさない地点でも、共通モデルでmodelsに追加されること。"""
        main_df, phenom_df, cols, mapping = _synthetic_observations(n_days=60)
        # AD地点はデータ不足になるよう大半を欠測にする
        phenom_df = phenom_df.copy()
        phenom_df.loc[phenom_df.index[50:], "AD"] = ""
        models = wv.train_location_models(main_df, phenom_df, cols, mapping, model="rf")
        self.assertIn("AC", models)
        self.assertIn("AD", models)
        self.assertTrue(models["AD"]["model_type"].startswith("pooled"))
        self.assertTrue(models["AD"]["pooled"])
        self.assertFalse(models["AC"]["pooled"])
        # rf_pipe/xgb_pipeは内部専用のキーなので、戻り値には残らないこと
        self.assertNotIn("rf_pipe", models["AC"])
        self.assertNotIn("xgb_pipe", models["AC"])


class TestModelCache(unittest.TestCase):
    def test_load_or_train_models_round_trips_through_cache(self):
        main_df, phenom_df, cols, mapping = _synthetic_observations(n_days=30)
        tmpdir = tempfile.mkdtemp(prefix="wv_cache_")
        try:
            cache_path = os.path.join(tmpdir, "models.joblib")
            self.assertFalse(os.path.isfile(cache_path))
            models = wv.load_or_train_models(main_df, phenom_df, cols, mapping, "rf", cache_path)
            self.assertTrue(os.path.isfile(cache_path))

            # 2回目は保存済みファイルを読み込むだけで、学習し直さないこと
            with mock.patch.object(wv, "train_location_models") as mocked:
                reloaded = wv.load_or_train_models(main_df, phenom_df, cols, mapping, "rf", cache_path)
                mocked.assert_not_called()
            self.assertEqual(set(reloaded.keys()), set(models.keys()))
        finally:
            shutil.rmtree(tmpdir, ignore_errors=True)


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

        p6, p6_wide = wv.export_prediction_csv(pred_df, prob_df, models, self.mapping, "只見", self.tmpdir)
        out = pd.read_csv(p6)
        self.assertEqual(list(out.columns), ["日時", "地点", "予測コード", "現象名", "霧確率(%)", "判定"])
        self.assertEqual(set(out["地点"]), {"地点甲", "地点乙"})
        self.assertEqual(len(out), len(forecast) * 2)  # 縦持ち＝地点数ぶん行が増える
        self.assertTrue(out["判定"].isin(["霧", "-"]).all())

        out_wide = pd.read_csv(p6_wide)
        self.assertIn("地点甲_予測コード", out_wide.columns)
        self.assertIn("地点甲_霧確率(%)", out_wide.columns)
        self.assertEqual(len(out_wide), len(forecast))

    def test_monthly_graphs(self):
        made = wv.plot_combo_by_month(self.main_df.head(200), self.phenom_df.head(200),
                                      self.cols, self.mapping, "只見", self.tmpdir)
        self.assertTrue(made)
        for p in made:
            self.assertTrue(os.path.isfile(p))

    def test_model_kind_switch(self):
        """--model の指定どおりのモデルが使われること（小さめのデータで確認）。"""
        main_df, phenom_df, _, _ = _synthetic_observations(n_days=20)
        args = (main_df, phenom_df, ["AC"], {"AC": "地点甲"})

        rf_only = wv.train_location_models(*args, model="rf")
        self.assertEqual(rf_only["AC"]["model_type"], "rf")
        self.assertGreater(rf_only["AC"]["accuracy"], 0.8)
        self.assertGreater(rf_only["AC"]["fog_f1"], 0.5)

        # xgboostが無い環境では hybrid 指定でも落ちずに rf へ降格する
        with mock.patch.object(wv, "_HAS_XGBOOST", False):
            fallback = wv.train_location_models(*args, model="hybrid")
        self.assertEqual(fallback["AC"]["model_type"], "rf")

    @unittest.skipUnless(wv._HAS_XGBOOST, "xgboost が導入されていない環境ではスキップ")
    def test_hybrid_training(self):
        main_df, phenom_df, _, _ = _synthetic_observations(n_days=20)
        models = wv.train_location_models(main_df, phenom_df, ["AC"], {"AC": "地点甲"},
                                          model="hybrid")
        info = models["AC"]
        self.assertEqual(info["model_type"], "hybrid")
        self.assertIn(info["blend_weight"], wv.BLEND_WEIGHTS)
        self.assertGreater(info["fog_f1"], 0.5)
        # 混ぜる前の単独モデルの成績も記録されていること（効果の確認用）
        self.assertFalse(np.isnan(info["rf_fog_f1"]))
        self.assertFalse(np.isnan(info["xgb_fog_f1"]))
        # ハイブリッドは単独モデルの良い方から大きく劣化しないこと
        self.assertGreaterEqual(info["fog_f1"],
                                min(info["rf_fog_f1"], info["xgb_fog_f1"]) - 0.05)

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

    def test_single_non_data_argument_is_the_output_dir(self):
        args = wv._parse_args(["--check-font", "出力先"])
        self.assertEqual(args.output, "出力先")
        self.assertEqual(args.input, wv.DEFAULT_INPUT_FILE)

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


# ---------------------------------------------------------------------------
# Colab向けのエントリーポイント run()
# ---------------------------------------------------------------------------

@unittest.skipUnless(os.path.isfile(SAMPLE_XLSX), f"{SAMPLE_XLSX} がありません")
class TestRunApi(unittest.TestCase):
    def test_run_generates_files(self):
        tmpdir = tempfile.mkdtemp(prefix="wv_run_")
        try:
            with mock.patch.object(wv.requests, "get",
                                   lambda *a, **k: _FakeResponse(_fake_forecast_payload(48))):
                # 学習方法の切り替えは別のテストで見るので、ここは速いrfで通す
                made = wv.run(SAMPLE_XLSX, tmpdir, monthly=False, show=False, model="rf")
            self.assertTrue(made)
            for p in made:
                self.assertTrue(os.path.isfile(p), p)
            self.assertTrue(any(str(p).endswith(".csv") for p in made))
        finally:
            shutil.rmtree(tmpdir, ignore_errors=True)

    def test_run_raises_clear_error_for_missing_file(self):
        with self.assertRaises(FileNotFoundError) as ctx:
            wv.run("存在しないファイル.xlsx", "/tmp/なんとか")
        self.assertIn("入力ファイルが見つかりません", str(ctx.exception))

    def test_zip_outputs(self):
        tmpdir = tempfile.mkdtemp(prefix="wv_zip_")
        try:
            out_dir = os.path.join(tmpdir, "out")
            os.makedirs(out_dir)
            with open(os.path.join(out_dir, "dummy.txt"), "w") as f:
                f.write("x")
            zip_path = wv.zip_outputs(out_dir)
            self.assertTrue(os.path.isfile(zip_path))
        finally:
            shutil.rmtree(tmpdir, ignore_errors=True)

    def test_display_is_noop_outside_notebook(self):
        self.assertEqual(wv.display_in_notebook(["a.png"]), [])


if __name__ == "__main__":
    unittest.main()
