from __future__ import annotations

import csv
import math
from pathlib import Path


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
            raise ValueError("矩阵不可逆，模型变量可能共线")
        aug[col], aug[pivot] = aug[pivot], aug[col]

        pivot_val = aug[col][col]
        aug[col] = [x / pivot_val for x in aug[col]]

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
    xty = matmul(xt, y_col)
    beta_col = matmul(xtx_inv, xty)
    beta = [row[0] for row in beta_col]

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


def load_rows() -> list[dict[str, str]]:
    with INPUT_FILE.open("r", encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))


def main() -> None:
    rows = load_rows()

    y: list[float] = []
    x1: list[list[float]] = []
    x2: list[list[float]] = []

    for row in rows:
        log_likes = float(row["log_likes"])
        sentiment_score = float(row["sentiment_score"])
        text_length = float(row["text_length"])
        imagenumber = float(row["imagenumber"])
        evening = 1.0 if row["time_period"] == "晚上" else 0.0
        afternoon = 1.0 if row["time_period"] == "下午" else 0.0
        morning = 1.0 if row["time_period"] == "上午" else 0.0
        positive_flag = float(row["positive_flag"])
        negative_flag = float(row["negative_flag"])

        y.append(log_likes)
        x1.append([1.0, sentiment_score, text_length, imagenumber, evening, afternoon, morning])
        x2.append([1.0, positive_flag, negative_flag, text_length, imagenumber, evening, afternoon, morning])

    beta1, se1, r2_1 = ols(y, x1)
    beta2, se2, r2_2 = ols(y, x2)

    model1_names = ["const", "sentiment_score", "text_length", "imagenumber", "evening", "afternoon", "morning"]
    model2_names = ["const", "positive_flag", "negative_flag", "text_length", "imagenumber", "evening", "afternoon", "morning"]

    output_file = OUTPUT_DIR / "小红书食品种草笔记_基础回归结果.csv"
    with output_file.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["model", "variable", "coef", "std_error", "t_value", "r_squared"])
        for name, coef, se in zip(model1_names, beta1, se1):
            writer.writerow(["model_1", name, round(coef, 6), round(se, 6), round(coef / se, 6) if se else "", round(r2_1, 6)])
        for name, coef, se in zip(model2_names, beta2, se2):
            writer.writerow(["model_2", name, round(coef, 6), round(se, 6), round(coef / se, 6) if se else "", round(r2_2, 6)])

    summary_file = OUTPUT_DIR / "小红书食品种草笔记_模型说明.md"
    summary_file.write_text(
        "\n".join(
            [
                "# 基础模型说明",
                "",
                "模型1：",
                "log_likes = b0 + b1*sentiment_score + b2*text_length + b3*imagenumber + b4*evening + b5*afternoon + b6*morning + e",
                "",
                "模型2：",
                "log_likes = b0 + b1*positive_flag + b2*negative_flag + b3*text_length + b4*imagenumber + b5*evening + b6*afternoon + b7*morning + e",
                "",
                "时间虚拟变量以“凌晨”为参照组。",
            ]
        ),
        encoding="utf-8",
    )

    print(f"已生成回归结果：{output_file}")


if __name__ == "__main__":
    main()
