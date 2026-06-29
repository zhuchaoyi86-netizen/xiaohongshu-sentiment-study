from __future__ import annotations

import csv
import math
from pathlib import Path
from statistics import mean


BASE_DIR = Path("/Users/xinxinhuashe/Documents/1/python课")
INPUT_FILE = BASE_DIR / "output/food_seed_notes/小红书食品种草笔记_情感分析数据.csv"
OUTPUT_DIR = BASE_DIR / "output/food_seed_notes"


def transpose(matrix: list[list[float]]) -> list[list[float]]:
    return [list(row) for row in zip(*matrix)]


def matmul(a: list[list[float]], b: list[list[float]]) -> list[list[float]]:
    result = [[0.0 for _ in range(len(b[0]))] for _ in range(len(a))]
    for i in range(len(a)):
        for k in range(len(b)):
            aik = a[i][k]
            for j in range(len(b[0])):
                result[i][j] += aik * b[k][j]
    return result


def invert_matrix(matrix: list[list[float]]) -> list[list[float]]:
    n = len(matrix)
    aug = [row[:] + [1.0 if i == j else 0.0 for j in range(n)] for i, row in enumerate(matrix)]
    for col in range(n):
        pivot = max(range(col, n), key=lambda r: abs(aug[r][col]))
        if abs(aug[pivot][col]) < 1e-12:
            raise ValueError("矩阵不可逆，变量可能存在共线性")
        aug[col], aug[pivot] = aug[pivot], aug[col]
        pivot_val = aug[col][col]
        aug[col] = [v / pivot_val for v in aug[col]]
        for row in range(n):
            if row == col:
                continue
            factor = aug[row][col]
            aug[row] = [rv - factor * cv for rv, cv in zip(aug[row], aug[col])]
    return [row[n:] for row in aug]


def ols(y: list[float], x: list[list[float]]) -> tuple[list[float], list[float], float]:
    xt = transpose(x)
    xtx = matmul(xt, x)
    xtx_inv = invert_matrix(xtx)
    y_col = [[v] for v in y]
    beta = [row[0] for row in matmul(matmul(xtx_inv, xt), y_col)]
    fitted = [sum(b * xi for b, xi in zip(beta, row)) for row in x]
    resid = [yi - fi for yi, fi in zip(y, fitted)]
    n = len(y)
    k = len(beta)
    sse = sum(r * r for r in resid)
    y_mean = sum(y) / n
    sst = sum((yi - y_mean) ** 2 for yi in y)
    r2 = 1 - sse / sst if sst else 0.0
    sigma2 = sse / (n - k)
    se = [math.sqrt(max(sigma2 * xtx_inv[i][i], 0.0)) for i in range(k)]
    return beta, se, r2


def load_rows() -> list[dict[str, float | str]]:
    rows: list[dict[str, float | str]] = []
    with INPUT_FILE.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        for row in reader:
            parsed: dict[str, float | str] = dict(row)
            for key in [
                "likes",
                "log_likes",
                "imagenumber",
                "text_length",
                "sentiment_score",
                "positive_word_count",
                "negative_word_count",
                "positive_flag",
                "negative_flag",
                "neutral_flag",
                "marketing_phrase_flag",
            ]:
                parsed[key] = float(row[key])
            rows.append(parsed)
    return rows


def build_controls(row: dict[str, float | str]) -> list[float]:
    return [
        float(row["text_length"]),
        float(row["imagenumber"]),
        1.0 if row["time_period"] == "晚上" else 0.0,
        1.0 if row["time_period"] == "下午" else 0.0,
        1.0 if row["time_period"] == "上午" else 0.0,
    ]


def summarize_groups(rows: list[dict[str, float | str]]) -> list[list[object]]:
    groups = {
        "positive": [r for r in rows if float(r["positive_flag"]) == 1.0],
        "negative": [r for r in rows if float(r["negative_flag"]) == 1.0],
        "neutral": [r for r in rows if float(r["neutral_flag"]) == 1.0],
        "high_intensity": [r for r in rows if float(r["sentiment_score"]) >= 6.0],
        "medium_intensity": [r for r in rows if 3.0 <= float(r["sentiment_score"]) <= 5.0],
        "low_positive": [r for r in rows if 1.0 <= float(r["sentiment_score"]) <= 2.0],
    }
    out: list[list[object]] = [["group", "count", "avg_likes", "avg_log_likes", "avg_score"]]
    for name, grp in groups.items():
        if not grp:
            continue
        out.append(
            [
                name,
                len(grp),
                round(mean(float(r["likes"]) for r in grp), 4),
                round(mean(float(r["log_likes"]) for r in grp), 4),
                round(mean(float(r["sentiment_score"]) for r in grp), 4),
            ]
        )
    return out


def main() -> None:
    rows = load_rows()
    y = [float(r["log_likes"]) for r in rows]

    models: list[tuple[str, list[str], list[list[float]]]] = []

    x1 = [[1.0, float(r["sentiment_score"]), *build_controls(r)] for r in rows]
    models.append(("model_1_score", ["const", "sentiment_score", "text_length", "imagenumber", "evening", "afternoon", "morning"], x1))

    x2 = [[1.0, float(r["positive_flag"]), float(r["negative_flag"]), *build_controls(r)] for r in rows]
    models.append(("model_2_direction", ["const", "positive_flag", "negative_flag", "text_length", "imagenumber", "evening", "afternoon", "morning"], x2))

    x3 = [[1.0, float(r["positive_word_count"]), float(r["negative_word_count"]), *build_controls(r)] for r in rows]
    models.append(("model_3_asymmetry", ["const", "positive_word_count", "negative_word_count", "text_length", "imagenumber", "evening", "afternoon", "morning"], x3))

    x4 = [[1.0, 1.0 if float(r["sentiment_score"]) >= 6.0 else 0.0, 1.0 if 3.0 <= float(r["sentiment_score"]) <= 5.0 else 0.0, 1.0 if 1.0 <= float(r["sentiment_score"]) <= 2.0 else 0.0, *build_controls(r)] for r in rows]
    models.append(("model_4_intensity", ["const", "high_intensity", "medium_intensity", "low_positive", "text_length", "imagenumber", "evening", "afternoon", "morning"], x4))

    output_file = OUTPUT_DIR / "小红书食品种草笔记_情感模型结果.csv"
    with output_file.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["model", "variable", "coef", "std_error", "t_value", "r_squared"])
        for model_name, var_names, x in models:
            beta, se, r2 = ols(y, x)
            for name, coef, sei in zip(var_names, beta, se):
                writer.writerow([model_name, name, round(coef, 6), round(sei, 6), round(coef / sei, 6) if sei else "", round(r2, 6)])

    group_file = OUTPUT_DIR / "小红书食品种草笔记_情感分组结果.csv"
    with group_file.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        writer.writerows(summarize_groups(rows))

    note_file = OUTPUT_DIR / "小红书食品种草笔记_情感模型说明.md"
    note_file.write_text(
        "\n".join(
            [
                "# 情感模型说明",
                "",
                "模型1：检验情感得分对点赞热度的总体影响。",
                "模型2：分别检验正向情绪与负向情绪的方向差异。",
                "模型3：分别检验正向词数量与负向词数量的不对称影响。",
                "模型4：按情感强度分层，比较低强度、中强度和高强度正向情绪对点赞热度的差异。",
                "",
                "所有模型的控制变量均包括：文本长度、配图数量和发布时间时段。",
                "时间虚拟变量以“凌晨”为参照组。",
            ]
        ),
        encoding="utf-8",
    )

    print(f"已生成情感模型结果：{output_file}")


if __name__ == "__main__":
    main()
