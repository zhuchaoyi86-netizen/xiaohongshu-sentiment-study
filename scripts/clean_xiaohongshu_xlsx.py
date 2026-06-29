from __future__ import annotations

import csv
import math
import re
from collections import Counter
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path
from typing import Iterable
from xml.etree import ElementTree as ET
from zipfile import ZipFile


BASE_DIR = Path("/Users/xinxinhuashe/Documents/1/python课")
INPUT_FILE = BASE_DIR / "python/数据集/小红书数据汇总/小红书对应数据.xlsx"
OUTPUT_DIR = BASE_DIR / "output/cleaned_xiaohongshu"

NS = {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
USE_COLUMNS = [
    "id",
    "title",
    "likes",
    "publishtime",
    "collects",
    "shareCount",
    "comments",
    "desc",
    "fans",
    "follows",
    "product",
    "category",
    "imagelist",
    "imagenumber",
]

RE_MULTI_SPACE = re.compile(r"\s+")
RE_INVALID_ONLY = re.compile(r"^[\W_]+$")
RE_FILLER_ONLY = re.compile(r"^(哈|啊|哦|嗯|滴|啦|呀|哇){3,}$")
RE_BAD_CHAR = re.compile(r"�+")
RE_URL = re.compile(r"https?://\S+|www\.\S+")
RE_MENTION = re.compile(r"@[\w\u4e00-\u9fff\-]+")
RE_TOPIC = re.compile(r"#([^#\s]+)")
RE_BRACKET_TAG = re.compile(r"\[[^\[\]]{1,12}\]")
RE_PRICE = re.compile(r"\b\d+(?:\.\d+)?r\b", re.IGNORECASE)
RE_LONG_DIGITS = re.compile(r"\b\d{5,}\b")
RE_REPEAT_PUNC = re.compile(r"([!！?？~～。.,，])\1+")
RE_ENUM_LINE = re.compile(r"^\s*\d+[\.、]\s*")
RE_INDEXED_LINE = re.compile(r"^\s*[0-9一二三四五六七八九十]+⃣️?")
RE_ACTIVITY_PHRASE = re.compile(
    r"(预售开启|支付定金|支付尾款|限量秒杀|每满\d+减\d+|惊喜玩法|图片价格仅供参考|理性种草|图文感谢|先到先得)"
)
RE_BOILERPLATE_PHRASE = re.compile(
    r"(无广|纯分享|仅供参考|根据需要|宝宝们快|冲了啦|大家理性种草|感谢|收藏住|快来看看同款)"
)
RE_NON_CJK_LATIN_DIGIT = re.compile(r"[A-Za-z0-9]")


@dataclass
class CleanSummary:
    raw_rows: int = 0
    missing_desc_rows: int = 0
    duplicate_id_rows: int = 0
    low_quality_rows: int = 0
    sentiment_too_short_rows: int = 0
    overly_marketing_rows: int = 0
    kept_rows: int = 0


def load_shared_strings(zf: ZipFile) -> list[str]:
    shared: list[str] = []
    if "xl/sharedStrings.xml" not in zf.namelist():
        return shared
    root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    for si in root.findall("a:si", NS):
        texts = [t.text or "" for t in si.iterfind(".//a:t", NS)]
        shared.append("".join(texts))
    return shared


def excel_value(cell: ET.Element, shared: list[str]) -> str:
    value_node = cell.find("a:v", NS)
    if value_node is None or value_node.text is None:
        return ""
    value = value_node.text
    if cell.attrib.get("t") == "s":
        return shared[int(value)]
    return value


def iter_xlsx_rows(path: Path) -> Iterable[list[str]]:
    with ZipFile(path) as zf:
        shared = load_shared_strings(zf)
        sheet = ET.fromstring(zf.read("xl/worksheets/sheet1.xml"))
        rows = sheet.findall(".//a:sheetData/a:row", NS)
        for row in rows:
            yield [excel_value(cell, shared) for cell in row.findall("a:c", NS)]


def normalize_text(text: str) -> str:
    text = text.replace("\r", "\n")
    text = RE_BAD_CHAR.sub("", text)
    text = RE_MULTI_SPACE.sub(" ", text)
    return text.strip()


def build_sentiment_text(text: str) -> str:
    raw_text = normalize_text(text)
    lines = [line.strip() for line in raw_text.replace(". ", "\n").split("\n")]
    kept_lines: list[str] = []
    for line in lines:
        if not line:
            continue
        if RE_ACTIVITY_PHRASE.search(line):
            continue
        if line.startswith("⚠️"):
            continue
        if RE_ENUM_LINE.match(line) and len(line) < 16:
            continue
        if RE_INDEXED_LINE.match(line) and len(line) < 10:
            continue
        kept_lines.append(line)

    text = " ".join(kept_lines) if kept_lines else raw_text
    text = RE_URL.sub(" ", text)
    text = RE_MENTION.sub(" ", text)
    text = RE_BRACKET_TAG.sub(" ", text)
    text = RE_TOPIC.sub(" ", text)
    text = RE_PRICE.sub(" ", text)
    text = RE_LONG_DIGITS.sub(" ", text)
    text = RE_ACTIVITY_PHRASE.sub(" ", text)
    text = RE_BOILERPLATE_PHRASE.sub(" ", text)
    text = RE_REPEAT_PUNC.sub(r"\1", text)
    text = RE_MULTI_SPACE.sub(" ", text)
    return text.strip()


def is_low_quality(text: str) -> bool:
    if not text:
        return True
    compact = text.replace(" ", "").replace("\n", "")
    if len(compact) < 5:
        return True
    if RE_INVALID_ONLY.fullmatch(compact):
        return True
    if RE_FILLER_ONLY.fullmatch(compact):
        return True
    unique_chars = set(compact)
    if len(compact) >= 8 and len(unique_chars) <= 2:
        return True
    return False


def sentiment_text_too_short(text: str) -> bool:
    compact = text.replace(" ", "").replace("\n", "")
    if len(compact) < 8:
        return True
    chinese_chars = re.findall(r"[\u4e00-\u9fff]", compact)
    if len(chinese_chars) < 4:
        return True
    return False


def is_overly_marketing_text(desc_clean: str, desc_sentiment: str) -> bool:
    compact_clean = desc_clean.replace(" ", "")
    compact_sentiment = desc_sentiment.replace(" ", "")
    if not compact_sentiment:
        return True
    if len(compact_sentiment) / max(len(compact_clean), 1) < 0.18:
        return True
    if compact_sentiment.count("优惠") + compact_sentiment.count("活动") + compact_sentiment.count("秒杀") >= 2:
        return True
    non_cjk_ratio = len(RE_NON_CJK_LATIN_DIGIT.findall(compact_sentiment)) / max(len(compact_sentiment), 1)
    if len(compact_sentiment) >= 20 and non_cjk_ratio > 0.45:
        return True
    return False


def parse_number(value: str) -> int:
    value = str(value).strip()
    if not value:
        return 0
    try:
        return int(float(value))
    except ValueError:
        digits = re.sub(r"[^\d.-]", "", value)
        return int(float(digits)) if digits else 0


def excel_serial_to_datetime(value: str) -> str:
    value = str(value).strip()
    if not value:
        return ""
    try:
        serial = float(value)
    except ValueError:
        return value
    base = datetime(1899, 12, 30)
    dt = base + timedelta(days=serial)
    return dt.strftime("%Y-%m-%d %H:%M:%S")


def time_period(dt_text: str) -> str:
    if not dt_text:
        return "未知"
    try:
        hour = datetime.strptime(dt_text, "%Y-%m-%d %H:%M:%S").hour
    except ValueError:
        return "未知"
    if 0 <= hour < 6:
        return "凌晨"
    if 6 <= hour < 12:
        return "上午"
    if 12 <= hour < 18:
        return "下午"
    return "晚上"


def build_clean_row(row: dict[str, str]) -> dict[str, object]:
    desc_clean = normalize_text(row["desc"])
    desc_sentiment = build_sentiment_text(row["desc"])
    likes = parse_number(row["likes"])
    imagenumber = parse_number(row["imagenumber"])
    publishtime_clean = excel_serial_to_datetime(row["publishtime"])
    return {
        "id": row["id"],
        "title": normalize_text(row["title"]),
        "desc": row["desc"],
        "desc_clean": desc_clean,
        "desc_sentiment": desc_sentiment,
        "likes": likes,
        "log_likes": round(math.log1p(likes), 6),
        "publishtime_raw": row["publishtime"],
        "publishtime_clean": publishtime_clean,
        "time_period": time_period(publishtime_clean),
        "collects": parse_number(row["collects"]),
        "shareCount": parse_number(row["shareCount"]),
        "comments": parse_number(row["comments"]),
        "fans": parse_number(row["fans"]),
        "follows": parse_number(row["follows"]),
        "product": normalize_text(row["product"]),
        "category": normalize_text(row["category"]),
        "imagenumber": imagenumber,
        "has_image": 1 if imagenumber > 0 else 0,
        "text_length": len(desc_clean),
        "sentiment_text_length": len(desc_sentiment),
        "hashtag_count": len(RE_TOPIC.findall(row["desc"])),
        "mention_count": len(RE_MENTION.findall(row["desc"])),
        "bracket_tag_count": len(RE_BRACKET_TAG.findall(row["desc"])),
        "exclamation_count": row["desc"].count("!") + row["desc"].count("！"),
        "question_count": row["desc"].count("?") + row["desc"].count("？"),
        "marketing_phrase_flag": 1 if RE_ACTIVITY_PHRASE.search(row["desc"]) or RE_BOILERPLATE_PHRASE.search(row["desc"]) else 0,
        "product_missing": 1 if not normalize_text(row["product"]) else 0,
        "category_missing": 1 if not normalize_text(row["category"]) else 0,
    }


def main() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    row_iter = iter_xlsx_rows(INPUT_FILE)
    headers = next(row_iter)
    header_index = {name: idx for idx, name in enumerate(headers)}

    selected_missing = [col for col in USE_COLUMNS if col not in header_index]
    if selected_missing:
        raise KeyError(f"缺少字段: {selected_missing}")

    summary = CleanSummary()
    seen_ids: set[str] = set()
    cleaned_rows: list[dict[str, object]] = []
    low_quality_examples: Counter[str] = Counter()

    for values in row_iter:
        summary.raw_rows += 1
        row = {col: values[header_index[col]] if header_index[col] < len(values) else "" for col in USE_COLUMNS}

        desc_clean = normalize_text(row["desc"])
        if not desc_clean:
            summary.missing_desc_rows += 1
            continue

        row_id = row["id"].strip()
        if row_id in seen_ids:
            summary.duplicate_id_rows += 1
            continue
        seen_ids.add(row_id)

        if is_low_quality(desc_clean):
            summary.low_quality_rows += 1
            low_quality_examples[desc_clean[:30]] += 1
            continue

        desc_sentiment = build_sentiment_text(row["desc"])
        if sentiment_text_too_short(desc_sentiment):
            summary.sentiment_too_short_rows += 1
            low_quality_examples[desc_sentiment[:30] or desc_clean[:30]] += 1
            continue

        if is_overly_marketing_text(desc_clean, desc_sentiment):
            summary.overly_marketing_rows += 1
            low_quality_examples[desc_sentiment[:30] or desc_clean[:30]] += 1
            continue

        cleaned_rows.append(build_clean_row(row))

    summary.kept_rows = len(cleaned_rows)

    cleaned_file = OUTPUT_DIR / "小红书商品笔记_清洗后数据.csv"
    with cleaned_file.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=list(cleaned_rows[0].keys()))
        writer.writeheader()
        writer.writerows(cleaned_rows)

    summary_file = OUTPUT_DIR / "小红书商品笔记_清洗汇总.csv"
    with summary_file.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["指标", "数值"])
        writer.writerow(["原始样本量", summary.raw_rows])
        writer.writerow(["删除正文缺失样本", summary.missing_desc_rows])
        writer.writerow(["删除重复id样本", summary.duplicate_id_rows])
        writer.writerow(["删除低质量文本样本", summary.low_quality_rows])
        writer.writerow(["删除情感文本过短样本", summary.sentiment_too_short_rows])
        writer.writerow(["删除营销噪声过强样本", summary.overly_marketing_rows])
        writer.writerow(["清洗后保留样本", summary.kept_rows])

    example_file = OUTPUT_DIR / "小红书商品笔记_低质量样本示例.csv"
    with example_file.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["文本前30字", "出现次数"])
        for text, count in low_quality_examples.most_common(30):
            writer.writerow([text, count])

    print(f"清洗完成，保留 {summary.kept_rows} / {summary.raw_rows} 条样本。")
    print(f"输出目录：{OUTPUT_DIR}")


if __name__ == "__main__":
    main()
