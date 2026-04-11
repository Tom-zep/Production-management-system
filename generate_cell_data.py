"""
generate_cell_data.py
---------------------
セル生産ライン用サンプルデータ生成

設計条件:
  - 8時間シフト（480分）、外段取り約11分
  - 目標290台/シフト、全工程残業なし
  - アイドル: 田中19分 / 鈴木10分(BN) / 佐藤19分 / 山田57分
"""

import numpy as np
import pandas as pd

np.random.seed(42)

SHIFT_MIN    = 480
SHIFT_SEC    = 28800
TARGET_UNITS = 290
DAYS         = 20

CELL_WORKERS = {
    "工程1": "田中",
    "工程2": "鈴木",   # ボトルネック
    "工程3": "佐藤",
    "工程4": "山田",
}

# 平均サイクルタイム（秒/台）
MEAN_CYCLE = {
    "工程1": 93.2,
    "工程2": 95.0,   # ボトルネック（最も遅い）
    "工程3": 93.2,
    "工程4": 85.3,
}

# チャンピオンタイム = 平均の約95%
CHAMPION_TIME = {p: round(v * 0.95, 1) for p, v in MEAN_CYCLE.items()}

SETUP_CHAMPION = 480  # 外段取りチャンピオン（秒）

records = []

for day in range(1, DAYS + 1):

    setup_sec = round(np.random.lognormal(
        mean=np.log(SETUP_CHAMPION * 1.3), sigma=0.20), 1)
    setup_min = round(setup_sec / 60, 1)

    for process, worker in CELL_WORKERS.items():
        mean  = MEAN_CYCLE[process]
        champ = CHAMPION_TIME[process]

        actual_cycle = round(np.random.normal(loc=mean, scale=mean * 0.03), 1)
        actual_cycle = max(champ, actual_cycle)

        productive_min = round((TARGET_UNITS * actual_cycle) / 60, 1)
        idle_min       = round(SHIFT_MIN - productive_min - setup_min, 1)

        records.append({
            "日付":              f"Day{day:02d}",
            "作業者":            worker,
            "工程":              process,
            "チャンピオンタイム": champ,
            "実サイクルタイム":   actual_cycle,
            "外段取り時間_秒":    setup_sec,
            "外段取り時間_分":    setup_min,
            "目標台数":           TARGET_UNITS,
            "稼働時間_分":        productive_min,
            "アイドル時間_分":    max(0.0, idle_min),
        })

df = pd.DataFrame(records)
df.to_csv("data/cell_production.csv", index=False, encoding="utf-8-sig")

print("=" * 55)
print("  セル生産サンプルデータ生成完了")
print("=" * 55)
print(f"  目標台数     : {TARGET_UNITS}台/シフト")
print(f"  外段取り平均 : {df['外段取り時間_分'].mean():.1f}分（チャンピオン{SETUP_CHAMPION//60}分）")
print()
summary = df.groupby("工程")[["実サイクルタイム","稼働時間_分","アイドル時間_分"]].mean().round(1)
print(summary)
