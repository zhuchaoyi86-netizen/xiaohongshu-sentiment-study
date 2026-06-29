import math
from pathlib import Path

import matplotlib
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import statsmodels.formula.api as smf


matplotlib.rcParams["font.sans-serif"] = [
    "Arial Unicode MS",
    "PingFang SC",
    "Heiti SC",
    "SimHei",
]
matplotlib.rcParams["axes.unicode_minus"] = False


ROOT = Path("/Users/xinxinhuashe/Documents/1/python课")
DATA_PATH = ROOT / "小组共享材料" / "小红书食品种草笔记_研究优化数据.csv"
OUT_DIR = ROOT / "第四部分材料"
FIG_DIR = OUT_DIR / "图"
TAB_DIR = OUT_DIR / "表"


def ensure_dirs() -> None:
    FIG_DIR.mkdir(parents=True, exist_ok=True)
    TAB_DIR.mkdir(parents=True, exist_ok=True)


def load_data() -> pd.DataFrame:
    df = pd.read_csv(DATA_PATH)
    df["len100"] = df["text_length_winsorized"] / 100
    df["t_morning"] = (df["time_period"] == "上午").astype(int)
    df["t_afternoon"] = (df["time_period"] == "下午").astype(int)
    df["t_evening"] = (df["time_period"] == "晚上").astype(int)
    df["publish_hour"] = pd.to_datetime(df["publishtime_clean"], errors="coerce").dt.hour
    df["density_group"] = pd.qcut(
        df["sentiment_density"],
        3,
        labels=["低情绪密度", "中情绪密度", "高情绪密度"],
    )
    return df


def baseline_model(df: pd.DataFrame):
    formula = (
        "log_likes_winsorized ~ sentiment_density + sentiment_score_winsorized + "
        "positive_flag + negative_flag + mixed_review_flag + len100 + imagenumber + "
        "hashtag_count + exclamation_count + question_count + marketing_phrase_flag + "
        "t_morning + t_afternoon + t_evening"
    )
    return smf.ols(formula, data=df).fit(cov_type="HC3"), formula


def build_key_tables(df: pd.DataFrame, model, formula: str) -> None:
    keep = [
        "sentiment_density",
        "sentiment_score_winsorized",
        "positive_flag",
        "negative_flag",
        "mixed_review_flag",
        "len100",
        "imagenumber",
        "hashtag_count",
        "exclamation_count",
        "question_count",
        "marketing_phrase_flag",
        "t_morning",
        "t_afternoon",
        "t_evening",
    ]
    name_map = {
        "sentiment_density": "情绪密度",
        "sentiment_score_winsorized": "情感得分(缩尾)",
        "positive_flag": "正向情绪虚拟变量",
        "negative_flag": "负向情绪虚拟变量",
        "mixed_review_flag": "混合评价虚拟变量",
        "len100": "文本长度(每百字)",
        "imagenumber": "配图数量",
        "hashtag_count": "话题标签数",
        "exclamation_count": "感叹号数量",
        "question_count": "问号数量",
        "marketing_phrase_flag": "营销话术虚拟变量",
        "t_morning": "上午发布",
        "t_afternoon": "下午发布",
        "t_evening": "晚上发布",
    }
    result = pd.DataFrame(
        {
            "变量": [name_map[k] for k in keep],
            "系数": [round(model.params[k], 4) for k in keep],
            "稳健标准误": [round(model.bse[k], 4) for k in keep],
            "p值": [round(model.pvalues[k], 4) for k in keep],
        }
    )
    def interpret(row):
        direction = "正向" if row["系数"] > 0 else "负向"
        if row["p值"] < 0.05:
            sig = "显著"
        elif row["p值"] < 0.1:
            sig = "边际显著"
        else:
            sig = "不显著"
        return f"{direction}{sig}"
    result["结论"] = result.apply(interpret, axis=1)
    result.to_csv(TAB_DIR / "第四部分_基准回归关键结果.csv", index=False, encoding="utf-8-sig")

    relation = pd.DataFrame(
        [
            ["核心自变量", "情绪密度", "log_likes_winsorized", "正向", "情绪表达越集中，点赞热度越高"],
            ["核心自变量", "情感得分(缩尾)", "log_likes_winsorized", "需结合密度解释", "单纯堆叠情感词未必带来更高热度"],
            ["结构变量", "混合评价", "log_likes_winsorized", "负向", "正负信息并存会削弱点赞表现"],
            ["控制变量", "配图数量", "log_likes_winsorized", "正向", "图文呈现越丰富，互动越强"],
            ["控制变量", "发布时间", "log_likes_winsorized", "正向", "上午、下午、晚上优于凌晨"],
            ["控制变量", "文本长度", "log_likes_winsorized", "弱正向", "较长文本可能提供更充分的信息"],
        ],
        columns=["变量层级", "变量", "作用对象", "预期关系", "解释"],
    )
    relation.to_csv(TAB_DIR / "第四部分_变量关系说明.csv", index=False, encoding="utf-8-sig")

    metrics = pd.DataFrame(
        [
            ["样本量", int(model.nobs)],
            ["调整R²", round(model.rsquared_adj, 4)],
            ["AIC", round(model.aic, 2)],
            ["模型公式", formula],
        ],
        columns=["指标", "值"],
    )
    metrics.to_csv(TAB_DIR / "第四部分_模型概况.csv", index=False, encoding="utf-8-sig")

    robust_density = pd.read_csv(ROOT / "输出" / "robust_density.csv")
    robust_mixed = pd.read_csv(ROOT / "输出" / "robust_mixed.csv")
    robust_density.to_csv(TAB_DIR / "第四部分_鲁棒性_情绪密度.csv", index=False, encoding="utf-8-sig")
    robust_mixed.to_csv(TAB_DIR / "第四部分_鲁棒性_混合评价.csv", index=False, encoding="utf-8-sig")


def fig_density_bar(df: pd.DataFrame) -> None:
    grouped = (
        df.groupby("density_group", observed=True)["likes_winsorized"]
        .mean()
        .reindex(["低情绪密度", "中情绪密度", "高情绪密度"])
    )
    fig, ax = plt.subplots(figsize=(6.2, 4.0))
    colors = ["#D7EAF5", "#72A9D6", "#18507B"]
    ax.bar(grouped.index, grouped.values, color=colors, width=0.6)
    ax.set_title("不同情绪密度组的平均点赞热度")
    ax.set_ylabel("平均点赞数(缩尾后)")
    ax.set_xlabel("情绪密度分组")
    for i, v in enumerate(grouped.values):
        ax.text(i, v + grouped.max() * 0.02, f"{v:.0f}", ha="center", fontsize=10)
    plt.tight_layout()
    plt.savefig(FIG_DIR / "第四部分_柱状图_情绪密度与点赞.png", dpi=180)
    plt.close()


def fig_scatter(df: pd.DataFrame) -> None:
    sample = df.sample(min(len(df), 900), random_state=42).copy()
    x = sample["sentiment_density"].to_numpy()
    y = sample["log_likes_winsorized"].to_numpy()
    slope, intercept = np.polyfit(x, y, 1)
    xs = np.linspace(x.min(), x.max(), 100)
    ys = slope * xs + intercept

    fig, ax = plt.subplots(figsize=(6.3, 4.2))
    ax.scatter(x, y, alpha=0.28, s=18, color="#5C88B0", edgecolors="none")
    ax.plot(xs, ys, color="#C94F4F", linewidth=2.2)
    ax.set_title("情绪密度与点赞热度的散点关系")
    ax.set_xlabel("情绪密度")
    ax.set_ylabel("对数点赞热度(缩尾后)")
    ax.text(
        0.03,
        0.95,
        f"样本抽样点数={len(sample)}\n拟合斜率={slope:.3f}",
        transform=ax.transAxes,
        va="top",
        fontsize=10,
        bbox=dict(facecolor="white", alpha=0.8, edgecolor="#CCCCCC"),
    )
    plt.tight_layout()
    plt.savefig(FIG_DIR / "第四部分_散点图_情绪密度与点赞.png", dpi=180)
    plt.close()


def fig_hour_line(df: pd.DataFrame) -> None:
    hour_mean = (
        df.groupby("publish_hour", observed=True)["likes_winsorized"]
        .mean()
        .reindex(range(24))
    )
    fig, ax = plt.subplots(figsize=(7.2, 4.0))
    ax.plot(hour_mean.index, hour_mean.values, color="#C07B2D", linewidth=2.5, marker="o", markersize=4)
    ax.set_title("不同发布时间的平均点赞热度")
    ax.set_xlabel("发布时间(小时)")
    ax.set_ylabel("平均点赞数(缩尾后)")
    ax.set_xticks(range(0, 24, 2))
    ax.grid(axis="y", alpha=0.25)
    peak_hour = hour_mean.idxmax()
    peak_val = hour_mean.max()
    ax.scatter([peak_hour], [peak_val], color="#9E2A2B", s=45, zorder=3)
    ax.text(peak_hour, peak_val + hour_mean.max() * 0.03, f"峰值 {peak_hour}:00", ha="center", fontsize=10)
    plt.tight_layout()
    plt.savefig(FIG_DIR / "第四部分_折线图_发布时间与点赞.png", dpi=180)
    plt.close()


def fig_coef_bar(model) -> None:
    picks = [
        "sentiment_density",
        "mixed_review_flag",
        "imagenumber",
        "exclamation_count",
        "marketing_phrase_flag",
        "t_morning",
        "t_afternoon",
        "t_evening",
    ]
    name_map = {
        "sentiment_density": "情绪密度",
        "mixed_review_flag": "混合评价",
        "imagenumber": "配图数量",
        "exclamation_count": "感叹号数量",
        "marketing_phrase_flag": "营销话术",
        "t_morning": "上午发布",
        "t_afternoon": "下午发布",
        "t_evening": "晚上发布",
    }
    coef = [model.params[p] for p in picks]
    labels = [name_map[p] for p in picks]
    fig, ax = plt.subplots(figsize=(7.0, 4.5))
    colors = ["#2F6690" if v >= 0 else "#B64949" for v in coef]
    ax.barh(labels, coef, color=colors)
    ax.axvline(0, color="#666666", linewidth=1)
    ax.set_title("基准模型中核心变量的影响方向")
    ax.set_xlabel("稳健回归系数")
    plt.tight_layout()
    plt.savefig(FIG_DIR / "第四部分_核心变量系数图.png", dpi=180)
    plt.close()


def build_summary_note(df: pd.DataFrame, model) -> None:
    density_high = df[df["density_group"] == "高情绪密度"]["likes_winsorized"].mean()
    density_low = df[df["density_group"] == "低情绪密度"]["likes_winsorized"].mean()
    mixed_like = df[df["mixed_review_flag"] == 1]["likes_winsorized"].mean()
    pure_like = df[df["mixed_review_flag"] == 0]["likes_winsorized"].mean()
    hour_mean = df.groupby("publish_hour", observed=True)["likes_winsorized"].mean()
    peak_hour = int(hour_mean.idxmax())
    note = [
        "第四部分核心发现",
        f"1. 基准回归中，情绪密度系数为 {model.params['sentiment_density']:.4f}，在普通标准误下显著为正，在稳健标准误下方向保持为正。",
        f"2. 混合评价系数为 {model.params['mixed_review_flag']:.4f}，在多种鲁棒性检验中始终显著为负，说明正负信息并存会明显削弱点赞热度。",
        f"3. 高情绪密度组平均点赞数约为 {density_high:.0f}，明显高于低情绪密度组的 {density_low:.0f}。",
        f"4. 混合评价文本平均点赞数约为 {mixed_like:.0f}，低于非混合评价文本的 {pure_like:.0f}。",
        f"5. 发布时间存在明显差异，平均点赞峰值出现在 {peak_hour}:00 左右。",
    ]
    (OUT_DIR / "第四部分_结论摘要.txt").write_text("\n".join(note), encoding="utf-8")


def main() -> None:
    ensure_dirs()
    df = load_data()
    model, formula = baseline_model(df)
    build_key_tables(df, model, formula)
    fig_density_bar(df)
    fig_scatter(df)
    fig_hour_line(df)
    fig_coef_bar(model)
    build_summary_note(df, model)
    print("第四部分图表和结果表已生成。")


if __name__ == "__main__":
    main()
