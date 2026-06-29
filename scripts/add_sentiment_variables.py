from __future__ import annotations

import csv
from pathlib import Path


BASE_DIR = Path("/Users/xinxinhuashe/Documents/1/python课")
INPUT_FILE = BASE_DIR / "output/food_seed_notes/小红书食品种草笔记_清洗后数据.csv"
OUTPUT_DIR = BASE_DIR / "output/food_seed_notes"

POSITIVE_WORDS = [
    "好吃",
    "香",
    "喜欢",
    "推荐",
    "安利",
    "值得",
    "回购",
    "宝藏",
    "惊艳",
    "满意",
    "真香",
    "绝了",
    "上头",
    "爱了",
    "不错",
    "方便",
    "平价",
    "便宜",
    "浓郁",
    "细腻",
    "酥脆",
    "顺滑",
    "入味",
    "过瘾",
    "解馋",
    "划算",
    "可以冲",
    "必囤",
    "无限回购",
]

NEGATIVE_WORDS = [
    "难吃",
    "不好吃",
    "一般",
    "失望",
    "踩雷",
    "避雷",
    "难喝",
    "油腻",
    "太甜",
    "太咸",
    "不值",
    "贵",
    "后悔",
    "无语",
    "不推荐",
    "翻车",
    "奇怪",
    "失望透顶",
    "腻",
    "雷",
    "踩坑",
    "不香",
    "齁甜",
    "寡淡",
    "发苦",
    "劝退",
    "不回购",
    "鸡肋",
    "不值这个价",
]


def count_matches(text: str, words: list[str]) -> int:
    return sum(text.count(word) for word in words)


def build_sentiment_fields(row: dict[str, str]) -> dict[str, str]:
    text = row.get("desc_sentiment", "")
    positive_count = count_matches(text, POSITIVE_WORDS)
    negative_count = count_matches(text, NEGATIVE_WORDS)
    sentiment_score = positive_count - negative_count

    row["positive_word_count"] = str(positive_count)
    row["negative_word_count"] = str(negative_count)
    row["sentiment_score"] = str(sentiment_score)
    row["positive_flag"] = "1" if sentiment_score > 0 else "0"
    row["negative_flag"] = "1" if sentiment_score < 0 else "0"
    row["neutral_flag"] = "1" if sentiment_score == 0 else "0"
    return row


def main() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    with INPUT_FILE.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        rows = [build_sentiment_fields(dict(row)) for row in reader]

    output_file = OUTPUT_DIR / "小红书食品种草笔记_情感分析数据.csv"
    with output_file.open("w", encoding="utf-8-sig", newline="") as f:
        fieldnames = list(rows[0].keys())
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)

    total = len(rows)
    positive_n = sum(int(row["positive_flag"]) for row in rows)
    negative_n = sum(int(row["negative_flag"]) for row in rows)
    neutral_n = sum(int(row["neutral_flag"]) for row in rows)
    avg_score = round(sum(int(row["sentiment_score"]) for row in rows) / total, 4) if total else 0

    summary_file = OUTPUT_DIR / "小红书食品种草笔记_情感分布汇总.csv"
    with summary_file.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["指标", "数值"])
        writer.writerow(["样本量", total])
        writer.writerow(["正向情绪样本数", positive_n])
        writer.writerow(["负向情绪样本数", negative_n])
        writer.writerow(["中性情绪样本数", neutral_n])
        writer.writerow(["正向情绪占比", round(positive_n / total, 4) if total else 0])
        writer.writerow(["负向情绪占比", round(negative_n / total, 4) if total else 0])
        writer.writerow(["中性情绪占比", round(neutral_n / total, 4) if total else 0])
        writer.writerow(["平均情感得分", avg_score])

    print(f"已生成情感分析数据：{total} 条")
    print(f"输出文件：{output_file}")


if __name__ == "__main__":
    main()
