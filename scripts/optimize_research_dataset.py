from __future__ import annotations

import csv
import math
from pathlib import Path


BASE_DIR = Path("/Users/xinxinhuashe/Documents/1/python课")
INPUT_FILE = BASE_DIR / "output/food_seed_notes/小红书食品种草笔记_情感分析数据.csv"
OUTPUT_DIR = BASE_DIR / "output/food_seed_notes"

NEGATIVE_MARKERS = ["踩雷", "避雷", "拔草", "红黑榜", "劝退", "不推荐"]
POSITIVE_MARKERS = ["种草", "推荐", "安利", "回购", "宝藏", "好吃"]


def quantile(sorted_vals: list[float], p: float) -> float:
    idx = int((len(sorted_vals) - 1) * p)
    return sorted_vals[idx]


def winsorize(value: float, lower: float, upper: float) -> float:
    return min(max(value, lower), upper)


def intensity_label(score: int) -> str:
    if score < 0:
        return "negative"
    if score == 0:
        return "neutral"
    if score <= 2:
        return "low_positive"
    if score <= 5:
        return "medium_positive"
    return "high_positive"


def mixed_review_flag(text: str) -> int:
    has_pos = any(word in text for word in POSITIVE_MARKERS)
    has_neg = any(word in text for word in NEGATIVE_MARKERS)
    return int(has_pos and has_neg)


def main() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    with INPUT_FILE.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        rows = list(reader)

    likes_vals = sorted(float(r["likes"]) for r in rows)
    text_vals = sorted(float(r["text_length"]) for r in rows)
    senti_vals = sorted(float(r["sentiment_score"]) for r in rows)

    likes_p01 = quantile(likes_vals, 0.01)
    likes_p99 = quantile(likes_vals, 0.99)
    text_p01 = quantile(text_vals, 0.01)
    text_p99 = quantile(text_vals, 0.99)
    senti_p01 = quantile(senti_vals, 0.01)
    senti_p99 = quantile(senti_vals, 0.99)

    out_rows: list[dict[str, str]] = []
    high_like_outlier_n = 0
    high_text_outlier_n = 0
    mixed_review_n = 0

    for row in rows:
        likes = float(row["likes"])
        text_length = float(row["text_length"])
        sentiment_score = int(float(row["sentiment_score"]))
        pos_count = int(float(row["positive_word_count"]))
        neg_count = int(float(row["negative_word_count"]))
        sentiment_len = max(int(float(row["sentiment_text_length"])), 1)
        merged_text = f"{row.get('title', '')} {row.get('desc_sentiment', '')}"

        likes_w = winsorize(likes, likes_p01, likes_p99)
        text_w = winsorize(text_length, text_p01, text_p99)
        senti_w = winsorize(sentiment_score, senti_p01, senti_p99)

        total_emotion_words = pos_count + neg_count
        positive_ratio = pos_count / total_emotion_words if total_emotion_words else 0.0
        negative_ratio = neg_count / total_emotion_words if total_emotion_words else 0.0
        sentiment_density = sentiment_score / sentiment_len * 100
        mixed_flag = mixed_review_flag(merged_text)

        row["likes_winsorized"] = str(round(likes_w, 4))
        row["log_likes_winsorized"] = str(round(math.log1p(likes_w), 6))
        row["high_like_outlier_flag"] = "1" if likes > likes_p99 else "0"
        row["text_length_winsorized"] = str(round(text_w, 4))
        row["high_text_outlier_flag"] = "1" if text_length > text_p99 else "0"
        row["sentiment_score_winsorized"] = str(round(senti_w, 4))
        row["sentiment_abs"] = str(abs(sentiment_score))
        row["sentiment_density"] = str(round(sentiment_density, 6))
        row["positive_ratio"] = str(round(positive_ratio, 6))
        row["negative_ratio"] = str(round(negative_ratio, 6))
        row["emotion_word_total"] = str(total_emotion_words)
        row["sentiment_intensity_level"] = intensity_label(sentiment_score)
        row["mixed_review_flag"] = str(mixed_flag)

        high_like_outlier_n += int(likes > likes_p99)
        high_text_outlier_n += int(text_length > text_p99)
        mixed_review_n += mixed_flag
        out_rows.append(row)

    output_file = OUTPUT_DIR / "小红书食品种草笔记_研究优化数据.csv"
    with output_file.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=list(out_rows[0].keys()))
        writer.writeheader()
        writer.writerows(out_rows)

    summary_file = OUTPUT_DIR / "小红书食品种草笔记_研究优化汇总.csv"
    with summary_file.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["指标", "数值"])
        writer.writerow(["样本量", len(out_rows)])
        writer.writerow(["点赞量P1", round(likes_p01, 4)])
        writer.writerow(["点赞量P99", round(likes_p99, 4)])
        writer.writerow(["文本长度P1", round(text_p01, 4)])
        writer.writerow(["文本长度P99", round(text_p99, 4)])
        writer.writerow(["情感得分P1", round(senti_p01, 4)])
        writer.writerow(["情感得分P99", round(senti_p99, 4)])
        writer.writerow(["高点赞离群样本数", high_like_outlier_n])
        writer.writerow(["超长文本样本数", high_text_outlier_n])
        writer.writerow(["混合评价样本数", mixed_review_n])

    print(f"已生成研究优化数据：{output_file}")


if __name__ == "__main__":
    main()
