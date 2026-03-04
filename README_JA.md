# RC梁最適化ツール 使用説明書（日本語）

## 1. 概要
`rc_beam_optimizer.py` は、RC梁の候補断面を評価して最適案を選定するツールです。  
Excel入力から、曲げ・せん断・たわみ・配筋条件をチェックし、`COST / CO2 / HYBRID` の目的関数で選定します。

## 2. 単位系
- 長さ: `mm`, `m`
- 力: `N`, `kN`
- 応力: `N/mm2`
- 重量: `kg`

補足:
- `N/mm2` は数値として `MPa` と同じです。
- 出力の鉄筋量は `kg` で表示されます。

## 3. 必要環境
- Python 3.10 以上（推奨）
- `openpyxl`

インストール例:
```powershell
py -m pip install openpyxl
```

## 4. ファイル構成
- 実行スクリプト: `rc_beam_optimizer.py`
- 入力テンプレート: `input_rc_beam.xlsx`
- 検証スクリプト: `validate_internet_cases.py`

## 5. 基本的な使い方
### 5.1 テンプレート作成
```powershell
py rc_beam_optimizer.py --make-template input_rc_beam.xlsx
```

### 5.2 最適化実行
```powershell
py rc_beam_optimizer.py input_rc_beam.xlsx output_rc_beam.xlsx
```

### 5.3 数値積分分割数を指定
```powershell
py rc_beam_optimizer.py input_rc_beam.xlsx output_rc_beam.xlsx --n-div 800
```

## 6. 入力Excel（必須シート）
必須シート:
- `MATERIALS`
- `SETTINGS`
- `COST`
- `BEAMS`
- `CANDIDATES`

### 6.1 MATERIALS（主なキー）
- `fc_n_mm2`: コンクリート強度
- `fy_main_n_mm2`: 主筋降伏強度
- `fy_shear_n_mm2`: せん断補強筋降伏強度
- `cover_mm`: かぶり
- `concrete_unit_weight_kn_m3`: コンクリート単位体積重量
- `steel_stress_limit_n_mm2`: 鉄筋応力度制限

### 6.2 SETTINGS（主なキー）
- `load_factor_dead`
- `load_factor_live`
- `objective_mode` (`COST`, `CO2`, `HYBRID`)
- `cost_weight`, `co2_weight`
- `output_all_checks`

### 6.3 COST（主なキー）
- `concrete_jpy_m3`
- `rebar_jpy_kg`
- `formwork_jpy_m2`
- `concrete_co2_kg_m3`
- `rebar_co2_kg_kg`
- `formwork_co2_kg_m2`

互換入力（旧キー）:
- `rebar_jpy_kn`
- `rebar_co2_kgco2_kn`

### 6.4 BEAMS
- `Use=TRUE` の行が計算対象です。
- 代表列: `BeamID`, `Span_m`, `TributaryWidth_m`, `DeadLoad_kN_m2`, `LiveLoad_kN_m2`, `PointDead_kN`, `PointLive_kN`, `PointPosRatio`

### 6.5 CANDIDATES
- `Use=TRUE` の行が候補です。
- 代表列: `Width_mm`, `Depth_mm`, `BottomBars_n`, `BottomBar`, `TopBars_n`, `TopBar`, `StirrupLegs_n`, `StirrupBar`, `StirrupSpacing_mm`

無補強筋（スターラップ無し）で扱う場合:
- `StirrupLegs_n = 0`
- `StirrupBar` は空欄で可

## 7. 出力Excel
出力ファイル `output_rc_beam.xlsx` に以下を作成します。
- `SUMMARY`: 梁ごとの採用断面と主要結果
- `CHECKS`: 全候補（または選抜）の照査結果
- `QUANTITY`: 数量・コスト・CO2集計
- `WARNINGS`: 入力警告、成立しない梁の警告

## 8. 妥当性確認（任意）
公開検証値との照合:
```powershell
py validate_internet_cases.py
```
`PASS` になれば照合は正常です。

## 9. よくあるエラー
- `Missing sheet: ...`
  - 入力Excelに必須シートがありません。テンプレートを再生成してください。
- `Unknown bar: ...`
  - 鉄筋径ラベルが未対応です。`D10, D13, D16, D19, D22, D25, D29, D32, D35` を使用してください。
- `No active rows in BEAMS / CANDIDATES`
  - `Use` 列が `TRUE` の行を設定してください。

