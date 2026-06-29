from __future__ import annotations

import csv
from pathlib import Path
from statistics import mean


BASE_DIR = Path("/Users/xinxinhuashe/Documents/1/python课")
INPUT_FILE = BASE_DIR / "output/food_seed_notes/小红书食品种草笔记_研究优化数据.csv"
OUTPUT_DIR = BASE_DIR / "output/food_seed_notes"


def write_csv(path: Path, rows: list[list[object]]) -> None:
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        writer.writerows(rows)


def main() -> None:
    with INPUT_FILE.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        rows = list(reader)

    for row in rows:
        for key in [
            "likes_winsorized",
            "log_likes_winsorized",
            "sentiment_density",
            "mixed_review_flag",
            "positive_ratio",
        ]:
            row[key] = float(row[key])

    intensity_rows = [["情绪强度层级", "样本量", "平均点赞量(缩尾后)", "平均对数点赞量(缩尾后)"]]
    for grp in ["negative", "neutral", "low_positive", "medium_positive", "high_positive"]:
        g = [r for r in rows if r["sentiment_intensity_level"] == grp]
        intensity_rows.append([grp, len(g), round(mean(r["likes_winsorized"] for r in g), 2), round(mean(r["log_likes_winsorized"] for r in g), 4)])
    write_csv(OUTPUT_DIR / "小红书食品种草笔记_变量分析_情绪强度.csv", intensity_rows)

    density_vals = sorted(r["sentiment_density"] for r in rows)
    q1 = density_vals[int((len(density_vals) - 1) * 0.33)]
    q2 = density_vals[int((len(density_vals) - 1) * 0.66)]
    density_groups = [
        ("low_density", [r for r in rows if r["sentiment_density"] <= q1]),
        ("mid_density", [r for r in rows if q1 < r["sentiment_density"] <= q2]),
        ("high_density", [r for r in rows if r["sentiment_density"] > q2]),
    ]
    density_rows = [["情绪密度分组", "样本量", "平均点赞量(缩尾后)", "平均对数点赞量(缩尾后)"]]
    for name, g in density_groups:
        density_rows.append([name, len(g), round(mean(r["likes_winsorized"] for r in g), 2), round(mean(r["log_likes_winsorized"] for r in g), 4)])
    write_csv(OUTPUT_DIR / "小红书食品种草笔记_变量分析_情绪密度.csv", density_rows)

    mixed_rows = [["评价类型", "样本量", "平均点赞量(缩尾后)", "平均对数点赞量(缩尾后)"]]
    for label, val in [("单一评价", 0.0), ("混合评价", 1.0)]:
        g = [r for r in rows if r["mixed_review_flag"] == val]
        mixed_rows.append([label, len(g), round(mean(r["likes_winsorized"] for r in g), 2), round(mean(r["log_likes_winsorized"] for r in g), 4)])
    write_csv(OUTPUT_DIR / "小红书食品种草笔记_变量分析_混合评价.csv", mixed_rows)

    direction_rows = [["情感方向变量", "样本量", "平均点赞量(缩尾后)", "平均对数点赞量(缩尾后)"]]
    for label, key in [("positive_flag", "positive_flag"), ("negative_flag", "negative_flag"), ("neutral_flag", "neutral_flag")]:
        g = [r for r in rows if float(r[key]) == 1.0]
        direction_rows.append([label, len(g), round(mean(r["likes_winsorized"] for r in g), 2), round(mean(r["log_likes_winsorized"] for r in g), 4)])
    write_csv(OUTPUT_DIR / "小红书食品种草笔记_变量分析_情感方向.csv", direction_rows)

    structure_rows = [["情感结构分组", "样本量", "平均点赞量(缩尾后)", "平均对数点赞量(缩尾后)"]]
    high_pos = [r for r in rows if float(r["positive_ratio"]) >= 0.8]
    mid_pos = [r for r in rows if 0.5 <= float(r["positive_ratio"]) < 0.8]
    low_pos = [r for r in rows if float(r["positive_ratio"]) < 0.5]
    for label, g in [("高正向占比", high_pos), ("中等正向占比", mid_pos), ("低正向占比", low_pos)]:
        structure_rows.append([label, len(g), round(mean(r["likes_winsorized"] for r in g), 2), round(mean(r["log_likes_winsorized"] for r in g), 4)])
    write_csv(OUTPUT_DIR / "小红书食品种草笔记_变量分析_情感结构.csv", structure_rows)

    image_rows = [["配图数量", "样本量", "平均点赞量(缩尾后)"]]
    for n in range(1, 10):
        g = [r for r in rows if int(float(r["imagenumber"])) == n]
        if g:
            image_rows.append([n, len(g), round(mean(r["likes_winsorized"] for r in g), 2)])
    write_csv(OUTPUT_DIR / "小红书食品种草笔记_变量分析_配图数量.csv", image_rows)

    time_rows = [["发布时间时段", "样本量", "平均点赞量(缩尾后)", "平均对数点赞量(缩尾后)"]]
    for label in ["凌晨", "上午", "下午", "晚上"]:
        g = [r for r in rows if r["time_period"] == label]
        time_rows.append([label, len(g), round(mean(r["likes_winsorized"] for r in g), 2), round(mean(r["log_likes_winsorized"] for r in g), 4)])
    write_csv(OUTPUT_DIR / "小红书食品种草笔记_变量分析_发布时间.csv", time_rows)

    likes_rows = [
        ["点赞热度变量", "说明"],
        ["likes", "原始点赞数，用于保留真实互动规模信息。"],
        ["log_likes", "对原始点赞数进行对数变换后的结果，用于缓解右偏分布。"],
        ["likes_winsorized", "对原始点赞数进行缩尾处理后的结果，用于降低极端爆款样本影响。"],
        ["log_likes_winsorized", "对缩尾后的点赞数再取对数，更适合后续稳健分析。"],
        ["high_like_outlier_flag", "标记高点赞离群样本，便于识别极端观测值。"],
    ]
    write_csv(OUTPUT_DIR / "小红书食品种草笔记_变量分析_点赞热度变量说明.csv", likes_rows)

    conclusion_rows = [
        ["变量", "对应比较", "初步发现"],
        ["positive_flag / negative_flag / neutral_flag", "正向、负向、中性文本的点赞表现", "positive_flag 对应样本平均点赞高于 negative_flag，说明正向情绪文本整体更容易获得点赞。"],
        ["sentiment_intensity_level", "negative、neutral、low_positive、medium_positive、high_positive 的差异", "正向情绪内部不同强度层级存在差异，说明情绪强度并非越高越好。"],
        ["sentiment_density", "低、中、高情绪密度分组比较", "情绪密度越高，平均点赞表现越高，说明情绪表达越集中越容易引发互动。"],
        ["positive_ratio / negative_ratio", "不同情感结构分组比较", "高正向占比文本整体点赞表现更好，说明文本内部情绪结构会影响互动表现。"],
        ["mixed_review_flag", "单一评价与混合评价比较", "混合评价文本平均点赞低于单一评价文本，说明正负并存内容更容易削弱互动表现。"],
        ["imagenumber", "不同配图数量样本比较", "多图内容整体互动表现好于少图内容，说明图文呈现方式仍是重要辅助因素。"],
        ["time_period", "凌晨、上午、下午、晚上比较", "上午、下午和晚上发布内容的点赞表现整体高于凌晨。"],
        ["likes_winsorized / log_likes_winsorized", "原始点赞热度优化处理", "缩尾后点赞指标能够减弱爆款极端值干扰，更适合后续稳健分析。"],
    ]
    write_csv(OUTPUT_DIR / "小红书食品种草笔记_变量分析_变量结论对应表.csv", conclusion_rows)

    print("已生成变量分析表格。")


if __name__ == "__main__":
    main()
