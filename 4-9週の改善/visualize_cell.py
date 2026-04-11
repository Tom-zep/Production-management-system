"""
visualize_cell.py
-----------------
セル生産ライン専用グラフ（日本語）
4工程固定分業 + 外段取り1人 構成
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import matplotlib.font_manager as fm
import seaborn as sns

# ── フォント ──────────────────────────────────────────────────────────────────
_jp = next((f.name for f in fm.fontManager.ttflist
            if "Noto Sans CJK JP" in f.name), "IPAGothic")
plt.rcParams.update({
    "font.family": _jp,
    "axes.spines.top": False,
    "axes.spines.right": False,
    "figure.facecolor": "white",
})
sns.set_theme(style="whitegrid", font_scale=1.05)
plt.rcParams["font.family"] = _jp   # seabornが上書きするので再設定

C = {"champion": "#2ECC71", "actual": "#E74C3C",
     "idle": "#F39C12", "setup": "#9B59B6", "dark": "#2C3E50"}

CHAMPION = {"工程1": 88.5, "工程2": 90.3, "工程3": 88.5, "工程4": 81.0}

# ── データ読み込み ─────────────────────────────────────────────────────────────
df = pd.read_csv("data/cell_production.csv", encoding="utf-8-sig")


# ① 工程別サイクルタイム分布（箱ひげ＋チャンピオンライン）──────────────────────
def plot_cycle_time(ax):
    order  = ["工程1", "工程2", "工程3", "工程4"]
    labels = ["工程1\n田中", "工程2\n鈴木", "工程3\n佐藤", "工程4\n山田"]
    x = np.arange(len(order))
    means = [df[df["工程"] == p]["実サイクルタイム"].mean() for p in order]

    # 平均棒グラフ
    ax.bar(x, means, color="#5B9BD5", width=0.5, label="20日間の平均", zorder=2)

    # 全日の実測値を点でプロット
    for i, proc in enumerate(order):
        vals = df[df["工程"] == proc]["実サイクルタイム"].values
        jitter = np.random.uniform(-0.12, 0.12, size=len(vals))
        ax.scatter(i + jitter, vals, color=C["dark"], s=30,
                   alpha=0.6, zorder=3, label="各日の実測値" if i == 0 else "")

    # チャンピオンタイム横線
    for i, proc in enumerate(order):
        ax.hlines(CHAMPION[proc], i - 0.35, i + 0.35,
                  colors=C["champion"], linewidths=2.5, linestyles="--", zorder=5)
        ax.text(i + 0.38, CHAMPION[proc], f"最速{CHAMPION[proc]}秒",
                va="center", fontsize=8, color=C["champion"])

    # 平均値ラベル
    for i, m in enumerate(means):
        ax.text(i, m / 2, f"{m:.1f}秒", ha="center", va="center",
                fontsize=9, fontweight="bold", color="white")

    champ_patch = mpatches.Patch(color=C["champion"], label="最速記録（チャンピオン）")
    dot_patch   = mpatches.Patch(color=C["dark"],     label="各日の実測値")
    bar_patch   = mpatches.Patch(color="#5B9BD5",     label="20日間の平均")
    ax.legend(handles=[bar_patch, dot_patch, champ_patch], framealpha=0.9, fontsize=8)

    ax.set_xticks(x)
    ax.set_xticklabels(labels)
    ax.set_title("① 工程別サイクルタイム（20日間の平均と実測値）",
                 fontweight="bold", pad=10)
    ax.set_xlabel("工程（担当者固定）")
    ax.set_ylabel("サイクルタイム（秒/台）")
    ax.set_ylim(0, max(means) * 1.4)


# ② 日別ボトルネック工程の推移 ────────────────────────────────────────────────
def plot_bottleneck_trend(ax):
    pivot = df.pivot_table(
        index="日付", columns="工程", values="実サイクルタイム")
    days = pivot.index.tolist()
    x = np.arange(len(days))

    colors = ["#3498DB", "#E74C3C", "#2ECC71", "#F39C12"]
    for (col, color) in zip(pivot.columns, colors):
        ax.plot(x, pivot[col], marker="o", markersize=4,
                label=col, color=color, linewidth=1.8)

    # ボトルネック（最大）を強調
    bottleneck = pivot.max(axis=1)
    ax.fill_between(x, bottleneck, pivot.min(axis=1),
                    alpha=0.07, color=C["actual"], label="ばらつき幅")

    ax.set_xticks(x[::2])
    ax.set_xticklabels(days[::2], rotation=30, fontsize=8)
    ax.set_title("② 日別サイクルタイム推移（ボトルネック工程の把握）",
                 fontweight="bold", pad=10)
    ax.set_ylabel("サイクルタイム（秒/台）")
    ax.legend(ncol=4, fontsize=8, framealpha=0.9)


# ③ 480分の内訳：工程別の稼働・アイドル時間 ──────────────────────────────────
def plot_hidden_idle(ax):
    SHIFT_MIN = 480
    TARGET    = 290
    SETUP_MIN = 638 / 60

    order  = ["工程1", "工程2", "工程3", "工程4"]
    labels = ["工程1\n田中", "工程2\n鈴木\n(BN)", "工程3\n佐藤", "工程4\n山田"]
    x = np.arange(4)
    w = 0.5

    mean_cycle = df.groupby("工程")["実サイクルタイム"].mean()
    productive = [(TARGET * mean_cycle[p]) / 60 for p in order]
    idle       = [max(0, SHIFT_MIN - productive[i] - SETUP_MIN) for i in range(4)]

    ax.bar(x, productive,       width=w, color="#5B9BD5", label="稼働時間（分）",   zorder=3)
    ax.bar(x, [SETUP_MIN]*4,    width=w, bottom=productive,
           color=C["setup"], alpha=0.8, label=f"外段取り（{SETUP_MIN:.0f}分）", zorder=3)
    bottom2 = [p + SETUP_MIN for p in productive]
    ax.bar(x, idle,             width=w, bottom=bottom2,
           color=C["idle"], alpha=0.85, label="アイドル時間（分）", zorder=3)

    for i, (b, idl) in enumerate(zip(bottom2, idle)):
        if idl > 3:
            ax.text(x[i], b + idl / 2, f"{idl:.0f}分",
                    ha="center", va="center", fontsize=10,
                    color="white", fontweight="bold")

    ax.axhline(SHIFT_MIN, color=C["dark"], linestyle=":",
               linewidth=1.5, label="シフト時間 480分")
    ax.set_xticks(x)
    ax.set_xticklabels(labels)
    ax.set_ylabel("時間（分）")
    ax.set_title("③ 1シフト480分の内訳：稼働 vs アイドル（290台達成後）\n"
                 "アイドルが長い工程 = 工程設計の見直し対象",
                 fontweight="bold", pad=10)
    ax.legend(framealpha=0.9, fontsize=8)
    ax.set_ylim(0, 520)


# ④ 外段取り時間の推移 ────────────────────────────────────────────────────────
def plot_setup_trend(ax):
    setup = df.groupby("日付")["外段取り時間_秒"].first().reset_index()
    days  = setup["日付"].tolist()
    x     = np.arange(len(days))
    vals  = setup["外段取り時間_秒"].values

    ax.bar(x, vals / 60, color=C["setup"], alpha=0.7,
           label="実段取り時間（分）", width=0.6)
    ax.axhline(480 / 60, color=C["champion"], linestyle="--",
               linewidth=2, label=f"チャンピオン（{480//60}分）")
    ax.axhline(vals.mean() / 60, color=C["actual"], linestyle="-.",
               linewidth=1.8, label=f"平均（{vals.mean()/60:.1f}分）")

    ax.fill_between(x, 480 / 60, vals / 60,
                    where=(vals > 480), alpha=0.2, color=C["actual"],
                    label="チャンピオン超過")

    ax.set_xticks(x[::2])
    ax.set_xticklabels(days[::2], rotation=30, fontsize=8)
    ax.set_ylabel("外段取り時間（分）")
    ax.set_title("④ 外段取り時間の日別推移（担当：伊藤）\n"
                 "チャンピオン比との乖離 = SMED改善余地",
                 fontweight="bold", pad=10)
    ax.legend(framealpha=0.9, fontsize=8)


# ── 描画 ──────────────────────────────────────────────────────────────────────
fig, axes = plt.subplots(2, 2, figsize=(16, 11))
fig.suptitle(
    "セル生産ライン スループット分析\n"
    "4人セル（工程固定）＋外段取り1人 | 目標290台/シフト",
    fontsize=14, fontweight="bold", y=1.01,
)

plot_cycle_time(axes[0, 0])
plot_bottleneck_trend(axes[0, 1])
plot_hidden_idle(axes[1, 0])
plot_setup_trend(axes[1, 1])

plt.tight_layout(pad=2.5)
fig.savefig("outputs/cell_analysis.png", dpi=150,
            bbox_inches="tight", facecolor="white")
print("保存完了 → outputs/cell_analysis.png")
plt.close()
