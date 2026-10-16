# -*- coding: utf-8 -*-
"""
内置评价指标算法库（仅依赖 numpy / pandas）。

管理员在后台选定某个指标作为启用指标后，评分流程读取该设定，
把标准答案的真实值与学生提交的预测值交给对应算法计算。
"""

import math

import numpy as np
import pandas as pd

# 学习类型
TASK_TYPES = {
    "regression": "回归",
    "classification": "分类",
    "clustering": "聚类",
}


class MetricError(ValueError):
    """指标计算不成立时抛出，消息直接展示给用户。"""


def _numeric(values, field):
    s = pd.Series(list(values), dtype=str).fillna("")
    num = pd.to_numeric(s, errors="coerce")
    bad = s[num.isna()]
    if len(bad) > 0:
        sample = "、".join(bad.astype(str).head(3).tolist())
        raise MetricError(f"{field}存在非数字取值（如 {sample}），无法按该指标计算。")
    return num.to_numpy(dtype=float)


def _labels(values):
    arr = np.array(["" if v is None else str(v).strip() for v in values], dtype=object)
    if (arr == "").any():
        raise MetricError("存在空的标签取值，无法按该指标计算。")
    return arr


def _combine(values, field, numeric):
    return _numeric(values, field) if numeric else _labels(values)


# ---------------------------------------------------------------- 回归
def calc_mse(y_true, y_pred):
    t, p = _numeric(y_true, "标准答案真实值"), _numeric(y_pred, "提交预测值")
    return float(np.mean((t - p) ** 2))


def calc_rmse(y_true, y_pred):
    return math.sqrt(calc_mse(y_true, y_pred))


def calc_mae(y_true, y_pred):
    t, p = _numeric(y_true, "标准答案真实值"), _numeric(y_pred, "提交预测值")
    return float(np.mean(np.abs(t - p)))


def calc_mape(y_true, y_pred):
    t, p = _numeric(y_true, "标准答案真实值"), _numeric(y_pred, "提交预测值")
    mask = t != 0
    if not mask.any():
        raise MetricError("标准答案真实值全部为 0，MAPE 无法计算。")
    return float(np.mean(np.abs((t[mask] - p[mask]) / t[mask])) * 100)


def calc_r2(y_true, y_pred):
    t, p = _numeric(y_true, "标准答案真实值"), _numeric(y_pred, "提交预测值")
    ss_res = float(np.sum((t - p) ** 2))
    ss_tot = float(np.sum((t - np.mean(t)) ** 2))
    if ss_tot == 0:
        raise MetricError("标准答案真实值全部相同，R² 无定义。")
    return 1 - ss_res / ss_tot


# ---------------------------------------------------------------- 分类
def _classification_counts(y_true, y_pred):
    t = _labels(y_true)
    p = _labels(y_pred)
    labels = sorted(set(t.tolist()) | set(p.tolist()))
    if len(labels) < 2:
        raise MetricError("真实值与预测值中只出现 1 个类别，分类指标无意义。")
    return t, p, labels


def calc_accuracy(y_true, y_pred):
    t, p, _ = _classification_counts(y_true, y_pred)
    return float(np.mean(t == p))


def _macro_prf(y_true, y_pred, which):
    """宏平均 P/R/F1：先按类别分别计算，再对类别取算术平均（F1 为各类 F1 的均值）。"""
    t, p, labels = _classification_counts(y_true, y_pred)
    precs, recs, f1s = [], [], []
    for label in labels:
        is_t = t == label
        is_p = p == label
        tp = int(np.sum(is_t & is_p))
        prec = tp / int(np.sum(is_p)) if int(np.sum(is_p)) else 0.0
        rec = tp / int(np.sum(is_t)) if int(np.sum(is_t)) else 0.0
        precs.append(prec)
        recs.append(rec)
        f1s.append(2 * prec * rec / (prec + rec) if prec + rec else 0.0)
    return {"precision": float(np.mean(precs)),
            "recall": float(np.mean(recs)),
            "f1": float(np.mean(f1s))}[which]


def calc_macro_precision(y_true, y_pred):
    return _macro_prf(y_true, y_pred, "precision")


def calc_macro_recall(y_true, y_pred):
    return _macro_prf(y_true, y_pred, "recall")


def calc_macro_f1(y_true, y_pred):
    return _macro_prf(y_true, y_pred, "f1")


# ---------------------------------------------------------------- 聚类
def _contingency(y_true, y_pred):
    t, p = _labels(y_true), _labels(y_pred)
    ct = pd.crosstab(pd.Series(t), pd.Series(p)).to_numpy(dtype=float)
    if min(ct.shape) < 2:
        raise MetricError("真实类别与预测类别均需至少 2 类，聚类一致性指标才有意义。")
    return ct


def _entropy(counts):
    n = float(counts.sum())
    p = counts[counts > 0] / n
    return float(-(p * np.log(p)).sum()) if len(p) else 0.0


def calc_nmi(y_true, y_pred):
    """标准化互信息 NMI（几何平均归一化，自然对数，与熵口径一致）。"""
    ct = _contingency(y_true, y_pred)
    n = float(ct.sum())
    row, col = ct.sum(axis=1), ct.sum(axis=0)
    outer = np.outer(row, col)
    mask = ct > 0
    mi = float(np.sum((ct[mask] / n) * np.log((ct[mask] * n) / outer[mask])))
    h_t, h_p = _entropy(row), _entropy(col)
    if h_t == 0 and h_p == 0:
        return 1.0
    denom = math.sqrt(h_t * h_p)
    return mi / denom if denom else 0.0


def calc_ari(y_true, y_pred):
    """调整兰德指数 ARI。"""
    ct = _contingency(y_true, y_pred)
    n = float(ct.sum())
    comb = lambda x: x * (x - 1) / 2  # noqa: E731
    sum_ij = float(comb(ct).sum())
    sum_a = float(comb(ct.sum(axis=1)).sum())
    sum_b = float(comb(ct.sum(axis=0)).sum())
    total = comb(n)
    expected = sum_a * sum_b / total if total else 0.0
    max_index = (sum_a + sum_b) / 2
    denom = max_index - expected
    return (sum_ij - expected) / denom if denom else 0.0


# ---------------------------------------------------------------- 指标注册表
# needs_numeric: 预测值/真实值必须是数字；否则按类别标签处理
METRIC_LIB = {
    "mse": {
        "label": "均方误差 MSE", "tasks": ("regression",), "direction": "min",
        "needs_numeric": True, "decimals": 6, "fn": calc_mse,
        "desc": "均方误差 MSE = (1/n)·Σ(y_true − y_pred)²，越小越好。",
    },
    "rmse": {
        "label": "均方根误差 RMSE", "tasks": ("regression",), "direction": "min",
        "needs_numeric": True, "decimals": 6, "fn": calc_rmse,
        "desc": "RMSE = √MSE，与原始取值同量纲，越小越好。",
    },
    "mae": {
        "label": "平均绝对误差 MAE", "tasks": ("regression",), "direction": "min",
        "needs_numeric": True, "decimals": 6, "fn": calc_mae,
        "desc": "MAE = (1/n)·Σ|y_true − y_pred|，越小越好。",
    },
    "mape": {
        "label": "平均百分比误差 MAPE", "tasks": ("regression",), "direction": "min",
        "needs_numeric": True, "decimals": 4, "fn": calc_mape,
        "desc": "MAPE = mean(|y_true − y_pred| / |y_true|) × 100%，单位为百分比，越小越好。",
    },
    "r2": {
        "label": "决定系数 R²", "tasks": ("regression",), "direction": "max",
        "needs_numeric": True, "decimals": 6, "fn": calc_r2,
        "desc": "R² = 1 − SS_res / SS_tot，越接近 1 越好。",
    },
    "accuracy": {
        "label": "准确率 Accuracy", "tasks": ("classification",), "direction": "max",
        "needs_numeric": False, "decimals": 4, "fn": calc_accuracy,
        "desc": "Accuracy = 预测正确的样本数 / 总样本数，越大越好。",
    },
    "macro_precision": {
        "label": "宏平均精确率 Macro-Precision", "tasks": ("classification",), "direction": "max",
        "needs_numeric": False, "decimals": 4, "fn": calc_macro_precision,
        "desc": "各类别精确率的算术平均，越大越好。",
    },
    "macro_recall": {
        "label": "宏平均召回率 Macro-Recall", "tasks": ("classification",), "direction": "max",
        "needs_numeric": False, "decimals": 4, "fn": calc_macro_recall,
        "desc": "各类别召回率的算术平均，越大越好。",
    },
    "macro_f1": {
        "label": "宏平均 F1 Macro-F1", "tasks": ("classification",), "direction": "max",
        "needs_numeric": False, "decimals": 4, "fn": calc_macro_f1,
        "desc": "各类别 F1 的算术平均（宏平均 F1），越大越好。",
    },
    "nmi": {
        "label": "标准化互信息 NMI", "tasks": ("clustering",), "direction": "max",
        "needs_numeric": False, "decimals": 4, "fn": calc_nmi,
        "desc": "NMI = MI / √(H(真实)·H(预测))（几何平均归一化），取值 0~1，越大越好。",
    },
    "ari": {
        "label": "调整兰德指数 ARI", "tasks": ("clustering",), "direction": "max",
        "needs_numeric": False, "decimals": 4, "fn": calc_ari,
        "desc": "ARI 校正随机因素后的成对一致性，取值 −1~1，越大越好。",
    },
}

DEFAULT_METHOD = "mse"


def get_method(method):
    """返回指标算法定义；未登记的算法一律拒绝，避免评分口径失控。"""
    meta = METRIC_LIB.get(method)
    if meta is None:
        raise MetricError(f"未知的指标算法：{method}")
    return meta


def methods_for_task(task_type):
    return [(key, meta) for key, meta in METRIC_LIB.items() if task_type in meta["tasks"]]
