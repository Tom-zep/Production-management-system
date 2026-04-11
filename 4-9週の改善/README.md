# Factory Throughput Analyzer — Cell Production Edition 🏭

**EN:** A data analysis tool that visualises hidden idle time and bottleneck structure in cell production lines. Built from real Gemba observations, not theory.

**JP:** セル生産ラインの「隠れアイドル時間」とボトルネック構造を可視化するデータ分析ツール。現場（Gemba）での気づきから生まれたコードです。

> *"Quota management hides the true potential of every worker on the floor."*
> *「出来高管理は、現場の真のポテンシャルを隠蔽する。」*

---

## Background / 背景

**EN:** In a 4-person cell line with fixed process assignment, hitting the daily target means the fastest workers simply wait. Their skill never becomes the organisation's asset. This tool makes that invisible waste visible — and points to where to act.

**JP:** 工程固定の4人セルでは、目標台数に達した時点で速い作業者がただ待つだけになる。個人のスキルが組織の資産にならない。このツールはその「見えない損失」を定量化し、次の打ち手を示す。

---

## Cell Configuration / セル構成

| Role / 役割 | Worker / 担当者 | Avg Cycle / 平均サイクル | Idle / アイドル |
|---|---|---|---|
| 工程1 | 田中 | 93.6秒/台 | 約17分 |
| 工程2 | 鈴木 ⚠️ Bottleneck | 95.2秒/台 | 約10分 |
| 工程3 | 佐藤 | 92.4秒/台 | 約23分 |
| 工程4 | 山田 | 84.3秒/台 | **約62分 ← 要工程設計見直し** |
| 外段取り | 伊藤 | — | SMED改善対象 |

**Target / 目標：290台／シフト（8時間）**

---

## Dashboard / 分析ダッシュボード

![Cell Production Analysis](outputs/cell_analysis.png)

| # | Chart / チャート | Key insight / 読み取りポイント |
|---|---|---|
| ① | 工程別サイクルタイム | 工程2（鈴木）がボトルネック。チャンピオンとの乖離が改善余地 |
| ② | 日別サイクルタイム推移 | 工程2の赤線がセル全体の出来高の天井を決める |
| ③ | 1シフト480分の内訳 | 山田さんのアイドル62分は鈴木さんの6倍 → 工程設計の問題 |
| ④ | 外段取り時間推移 | チャンピオン8分との乖離 = SMED改善余地 |

---

## Key Finding / 主な発見

**EN:** Process 4 (Yamada) has 62 min/shift of idle time — 6× more than the bottleneck Process 2 (Suzuki, 10 min). This is not a people problem. It is a process design problem. The fastest worker on the cell is also the most idle.

**JP:** 工程4（山田）のアイドル時間は62分／シフト — ボトルネックの工程2（鈴木、10分）の6倍。これは個人の問題ではなく、工程設計の問題。セルで最も速い作業者が、最もヒマになっている。

---

## Quickstart / クイックスタート

```bash
git clone https://github.com/YOUR_USERNAME/Factory-Throughput-Analyzer-Python.git
cd Factory-Throughput-Analyzer-Python
pip install -r requirements.txt

# Generate sample data / サンプルデータ生成
python data/generate_cell_data.py

# Generate dashboard / ダッシュボード生成
python visualize_cell.py
# → outputs/cell_analysis.png
```

---

## Adapt to Your Line / 自分の現場に合わせるには

`generate_cell_data.py` の上部の定数を書き換えるだけです：

```python
TARGET_UNITS = 290        # 目標台数/シフト
MEAN_CYCLE   = { ... }   # 各工程の平均サイクルタイム（秒）
CELL_WORKERS = { ... }   # 工程と担当者名
```

---

## File Structure / ファイル構成

```
├── data/
│   ├── generate_cell_data.py   # サンプルデータ生成
│   └── cell_production.csv     # 生成データ
├── visualize_cell.py           # ダッシュボード（4チャート）
├── outputs/
│   └── cell_analysis.png
└── requirements.txt
```

---

*Born from Gemba observation. / 机上の空論ではなく、現場の気づきから生まれたコード。*

**License:** MIT
