from __future__ import annotations

import csv
from pathlib import Path


BASE_DIR = Path("/Users/xinxinhuashe/Documents/1/python课")
INPUT_FILE = BASE_DIR / "output/cleaned_xiaohongshu/小红书商品笔记_清洗后数据.csv"
OUTPUT_DIR = BASE_DIR / "output/food_seed_notes"

SEED_KEYWORDS = [
    "种草",
    "推荐",
    "安利",
    "回购",
    "合集",
    "测评",
    "分享",
    "好吃",
    "必买",
    "囤货",
    "宝藏",
]

NEGATIVE_REVIEW_KEYWORDS = [
    "踩雷",
    "避雷",
]


def is_seed_note(row: dict[str, str]) -> bool:
    title = row.get("title", "")
    text = row.get("desc_sentiment", "")
    merged = f"{title} {text}"
    hit_count = sum(keyword in merged for keyword in SEED_KEYWORDS)
    sentiment_length = int(row.get("sentiment_text_length", "0") or 0)

    # 至少满足“标题/正文有明确种草表达”或“两个及以上种草关键词同时出现”
    if hit_count >= 2:
        return True
    if any(keyword in title for keyword in ["推荐", "种草", "安利", "回购", "合集", "测评", "好吃"]):
        return True
    if hit_count >= 1 and sentiment_length >= 30:
        return True
    return False


def negative_flag(row: dict[str, str]) -> int:
    text = f"{row.get('title', '')} {row.get('desc_sentiment', '')}"
    return int(any(keyword in text for keyword in NEGATIVE_REVIEW_KEYWORDS))


def main() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    with INPUT_FILE.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        rows = list(reader)

    food_rows = [row for row in rows if row.get("category") == "食品"]
    seed_rows = [row for row in food_rows if is_seed_note(row)]

    for row in seed_rows:
        row["is_food_seed_note"] = "1"
        row["negative_review_flag"] = str(negative_flag(row))

    output_file = OUTPUT_DIR / "小红书食品种草笔记_清洗后数据.csv"
    with output_file.open("w", encoding="utf-8-sig", newline="") as f:
        fieldnames = list(seed_rows[0].keys())
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(seed_rows)

    summary_file = OUTPUT_DIR / "小红书食品种草笔记_筛选汇总.csv"
    with summary_file.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["指标", "数值"])
        writer.writerow(["清洗后总样本量", len(rows)])
        writer.writerow(["食品类样本量", len(food_rows)])
        writer.writerow(["食品种草笔记样本量", len(seed_rows)])
        writer.writerow(["种草笔记占食品类比例", round(len(seed_rows) / len(food_rows), 4) if food_rows else 0])
        writer.writerow(
            ["含踩雷/避雷标记样本量", sum(negative_flag(row) for row in seed_rows)]
        )

    print(f"已生成食品种草笔记数据：{len(seed_rows)} 条")
    print(f"输出目录：{OUTPUT_DIR}")


if __name__ == "__main__":
    main()
