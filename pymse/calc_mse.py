# -*- coding: utf-8 -*-
"""
MSE（均方误差）计算脚本

功能：
    读取真实收率文件 ansA.csv 与
    预测收率文件 submit_v4_testA.csv，
    按“样本id”对齐后计算 MSE，并将结果写入 CSV。

MSE 公式：
    MSE = (1 / n) * Σ (y_pred - y_true)^2      # 值越小越好

输入文件（位于脚本上一级目录）：
    ../data/ansA.csv                     真实值（管理员上传优先；无表头：样本id,收率）
    ../ansA.csv                          真实值（项目默认，未上传时使用）
    ../submit_v4_testA.csv               预测值（UTF-8，无表头：样本id,收率）

输出文件（与本脚本同目录 pymse/）：
    mse_result.csv
"""

import io
import os
from datetime import datetime

import numpy as np
import pandas as pd

# ---------------------------------------------------------------- 路径配置
SCRIPT_DIR = os.path.abspath(os.path.dirname(__file__))          # pymse 文件夹
BASE_DIR = os.path.abspath(os.path.join(SCRIPT_DIR, os.pardir))  # 上一级 web2 文件夹

# 真实值：优先管理员上传的实际数据（data/ansA.csv），否则用项目根目录 ansA.csv
_upload_ans = os.path.join(BASE_DIR, "data", "ansA.csv")
TRUE_FILE = _upload_ans if os.path.exists(_upload_ans) else os.path.join(BASE_DIR, "ansA.csv")
PRED_FILE = os.path.join(BASE_DIR, "submit_v4_testA.csv")              # 预测值
OUT_FILE = os.path.join(SCRIPT_DIR, "mse_result.csv")                   # 结果输出


def read_two_col_csv(path):
    """
    读取“样本id、收率”两列文件，返回 DataFrame[sample_id, y]。
    编码自动识别（UTF-8-SIG / UTF-8 / GBK），
    自动识别并丢弃表头（第一行第二列不是数字时）。
    """
    with open(path, "rb") as f:
        raw = f.read()
    text = None
    for enc in ("utf-8-sig", "utf-8", "gbk"):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    if text is None:
        raise ValueError(f"{os.path.basename(path)} 编码无法识别（需 UTF-8 或 GBK）")

    df = pd.read_csv(io.StringIO(text), header=None,
                     dtype=str, skip_blank_lines=True)
    if df.shape[1] != 2:
        raise ValueError(f"{os.path.basename(path)} 应为两列，实际 {df.shape[1]} 列")

    # 首行第二列不是数字 -> 视为表头
    try:
        float(str(df.iloc[0, 1]).strip())
    except (ValueError, TypeError):
        df = df.iloc[1:].reset_index(drop=True)

    df.columns = ["sample_id", "y"]
    df["sample_id"] = df["sample_id"].astype(str).str.strip()
    df["y"] = pd.to_numeric(df["y"], errors="raise")
    return df


def main():
    # 1. 读取数据（编码自动识别：UTF-8-SIG / UTF-8 / GBK，自动识别表头）
    true_df = read_two_col_csv(TRUE_FILE)
    pred_df = read_two_col_csv(PRED_FILE)

    print(f"真实值样本数：{len(true_df)}（{os.path.basename(TRUE_FILE)}）")
    print(f"预测值样本数：{len(pred_df)}（{os.path.basename(PRED_FILE)}）")

    # 2. 按样本 id 内连接对齐，保证一一对应
    merged = true_df.merge(pred_df, on="sample_id", how="inner",
                           suffixes=("_true", "_pred"))
    missing = set(true_df["sample_id"]) - set(pred_df["sample_id"])
    extra = set(pred_df["sample_id"]) - set(true_df["sample_id"])
    if missing or extra:
        print(f"警告：缺少批次 {len(missing)} 个，多余批次 {len(extra)} 个；"
              f"MSE 仅基于 {len(merged)} 个共同批次计算")
    else:
        print(f"批次完全匹配，共 {len(merged)} 个样本")

    # 3. 计算 MSE
    diff = merged["y_pred"] - merged["y_true"]
    mse = float(np.mean(diff ** 2))
    print("\n========== 评价指标结果 ==========")
    print(f"MSE（均方误差）= {mse:.6f}")
    print(f"RMSE（均方根误差）= {np.sqrt(mse):.6f}")
    print(f"平均绝对偏差 MAE = {np.mean(np.abs(diff)):.6f}")
    print("===================================")
    print("解读：MSE 越小说明预测收率越接近真实检测收率，模型效果越好。")

    # 4. 结果写入 CSV（UTF-8-SIG，Excel 直接打开不乱码）
    result = pd.DataFrame([{
        "评价指标": "MSE",
        "MSE值": round(mse, 6),
        "真实值文件": os.path.basename(TRUE_FILE),
        "预测值文件": os.path.basename(PRED_FILE),
        "参与计算样本数": len(merged),
        "计算时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }])
    result.to_csv(OUT_FILE, index=False, encoding="utf-8-sig")
    print(f"\nMSE 结果已写入：{OUT_FILE}")


if __name__ == "__main__":
    main()
