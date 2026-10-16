# -*- coding: utf-8 -*-
"""
商务大数据分析课程竞赛 —— 在线评分系统（Web 版）

功能：
1. 用户管理：学生来自《工业24级名单.xlsx》，账号与初始密码均为学号；支持登录、修改密码
   内置系统管理员：admin / admin888
2. 上传提交：预测结果 CSV（两列：样本id、收率）+ 代码文件（可选，支持 py/ipynb/压缩包等）
3. 自动评分：管理员在后台设定评价指标（指标名称、学习类型、说明、计分算法等），
   学生提交后系统读取该设定，与标准答案按样本id对齐后计算指标值（启用指标唯一）
4. 成绩排名：所有学生按历史最佳指标值排名（方向由指标设定决定：越小越好 / 越大越好）
5. 管理员后台：用户管理（重置密码/删除）、查看排名、查询并下载学生的结果与代码文件
"""

import io
import math
import os
import re
import sqlite3
import urllib.parse
import urllib.request
import zipfile
from datetime import datetime, timedelta
from functools import wraps

import pandas as pd
from flask import (Blueprint, Flask, abort, flash, g, redirect,
                   render_template, request, send_file, send_from_directory,
                   session, url_for)
from markupsafe import Markup, escape
from werkzeug.datastructures import FileStorage
from werkzeug.security import check_password_hash, generate_password_hash
from werkzeug.utils import secure_filename

import metrics_lib
from metrics_lib import MetricError, METRIC_LIB, TASK_TYPES

# ---------------------------------------------------------------- 路径配置
BASE_DIR = os.path.abspath(os.path.dirname(__file__))
STUDENT_FILE = os.path.join(BASE_DIR, "工业24级名单.xlsx")   # 学生名单
ANSWER_FILE = os.path.join(BASE_DIR, "ansA.csv")  # 默认真实收率（标准答案）
DB_FILE = os.path.join(BASE_DIR, "scoring.db")
DATA_DIR = os.path.join(BASE_DIR, "data")  # 管理员上传的赛题数据文件目录

# 赛题数据文件：key -> (settings键, data目录下的规范文件名, 界面名称)
DATA_FILE_MAP = {
    "train": ("data_train", "train.csv", "训练数据"),
    "testA": ("data_testA", "testA.csv", "测试数据"),
    "sample": ("data_sample", "submit_testA.csv", "提交样例"),
    "answer": ("data_answer", "ansA.csv", "实际数据"),
}
ANSWER_KEY = "data_answer"  # 管理员上传的实际数据（评分标准答案）

# 赛题内容字段（管理员在后台维护，学生端“赛题与数据”页面据此生成）
TOPIC_KEYS = ("topic_title", "topic_desc", "topic_data", "topic_submit", "topic_metric")
DEFAULT_TOPIC = {
    "topic_title": "工业化工生产预测",
    "topic_desc": (
        "异烟酸用作医药中间体，主要用于制抗结核病药物异烟肼，也用于合成酰胺、酰肼、酯类等衍生物。"
        "烟酰胺生产过程包含水解脱色、结晶甩滤等过程。每个步骤会受到温度、时间、压强等各方面因素的影响，"
        "造成异烟酸收率的不稳定。为保证产品质量和提高生产效率，需要调整和优化生产过程中的参数。"
        "然而，根据传统经验的人工调整工艺参数费时费力。近年来，人工智能在工艺参数优化以及视频检测等领域"
        "取得了突飞猛进的成果。AI技术的发展有望助力原料药制造企业实现工艺生产革新，规范生产操作过程，"
        "从而达到提高产品的收率的目标。\n"
        "本次竞赛要求选手以异烟酸生产过程中的各参数，包括各主要步骤的时间、温度、压强等参数为基础，"
        "设计精确智能的优秀算法，提升异烟酸的收率。"
    ),
    "topic_data": (
        "竞赛包含有2000批次来自实际异烟酸生产中的各参数的监测指标和最终收率的数据。"
        "监测指标由两大工序数十个步骤构成。总生产步骤达30余项。"
        "我们将工序和步骤分别用字母和数字代号表示，比如A2，B5分别表示A工序第二步骤和B工序第五步骤。"
        "样例数据参考训练数据。"
    ),
    "topic_submit": (
        "要求选手通过使用生产过程的各参数，训练算法来预测最终异烟酸的收率。\n"
        "提交结果文件：每批次异烟酸预测收率，选手提交csv格式。该文件由两列组成，"
        "第一列为异烟酸批次id，第二列为预测的异烟酸预测收率，以小数形式表示，"
        "建议保留小数点后三位。参考：\n"
        "sample_1656,0.92\n"
        "sample_1548,0.8801\n"
        "提交代码文件，请将提交文件打包成一个ZIP压缩包，以“学号-姓名”命名，"
        "通过比赛页面“代码提交”入口提交，提交代码文件仅要求提交一次最终版代码（成绩最佳版）即可。"
    ),
    "topic_metric": "提交结果与实际收率结果进行对比，以均方误差（MSE）为评价指标，结果越小越好。",
}

# 上传文件落盘目录
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
RESULT_DIR = os.path.join(UPLOAD_DIR, "results")  # 预测结果 CSV
CODE_DIR = os.path.join(UPLOAD_DIR, "code")       # 代码文件

# 允许上传的代码文件扩展名
CODE_ALLOWED_EXT = {
    ".py", ".ipynb", ".r", ".m", ".java", ".c", ".cpp", ".h",
    ".zip", ".rar", ".7z", ".tar", ".gz",
    ".txt", ".pdf", ".doc", ".docx",
}

# 评价指标（管理员在后台维护，metrics 表；全表有且仅有一条 enabled = 1 作为当前评分口径）
DEFAULT_METRIC = {
    "name": "均方误差 MSE",
    "task_type": "regression",
    "method": metrics_lib.DEFAULT_METHOD,
    "direction": METRIC_LIB[metrics_lib.DEFAULT_METHOD]["direction"],
    "pred_column": "2",       # 提交结果文件第 2 列
    "true_column": "2",       # 标准答案文件第 2 列
    "description": METRIC_LIB[metrics_lib.DEFAULT_METHOD]["desc"],
}
ID_COLUMN = "1"               # 样本 id 固定取第 1 列

app = Flask(__name__)
# 生产部署请通过环境变量 SECRET_KEY 覆盖（会话签名密钥，切勿使用默认值）
app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY", "jinnan-scoring-system-2026-secret-key")
app.config["MAX_CONTENT_LENGTH"] = 32 * 1024 * 1024  # 单次上传上限 32MB


# ---------------------------------------------------------------- 数据库
def get_db():
    """每个请求复用一个 SQLite 连接。"""
    if "db" not in g:
        g.db = sqlite3.connect(DB_FILE)
        g.db.row_factory = sqlite3.Row
    return g.db


@app.teardown_appcontext
def close_db(exc):
    db = g.pop("db", None)
    if db is not None:
        db.close()


def _table_columns(db, table):
    return {r[1] for r in db.execute(f"PRAGMA table_info({table})").fetchall()}


SUBMISSIONS_DDL = """
    CREATE TABLE submissions (
        id              INTEGER PRIMARY KEY AUTOINCREMENT,
        user_id         INTEGER NOT NULL,
        score           REAL NOT NULL,           -- 启用指标下的本次得分
        metric_name     TEXT,                    -- 评分时使用的指标名称（快照）
        result_path     TEXT,                    -- 结果文件落盘文件名（basename）
        result_filename TEXT,                    -- 结果文件原始文件名（下载展示用）
        code_path       TEXT,                    -- 代码文件落盘文件名（basename）
        code_filename   TEXT,                    -- 代码文件原始文件名
        created_at      TEXT NOT NULL,
        FOREIGN KEY (user_id) REFERENCES users (id)
    )
"""

METRICS_DDL = """
    CREATE TABLE IF NOT EXISTS metrics (
        id          INTEGER PRIMARY KEY AUTOINCREMENT,
        name        TEXT NOT NULL,               -- 指标名称
        task_type   TEXT NOT NULL,               -- 学习类型 regression/classification/clustering
        method      TEXT NOT NULL,               -- 计分算法（metrics_lib 内置库 key）
        direction   TEXT NOT NULL,               -- min 越小越好 / max 越大越好
        pred_column TEXT NOT NULL DEFAULT '2',   -- 提交结果中参与计分的列（列序号或列名）
        true_column TEXT NOT NULL DEFAULT '2',   -- 标准答案中参与计分的列（列序号或列名）
        description TEXT,                        -- 指标说明（学生端展示）
        enabled     INTEGER NOT NULL DEFAULT 0,  -- 1 = 当前评分口径（全表仅一条）
        created_at  TEXT NOT NULL,
        updated_at  TEXT
    )
"""


def _migrate_submissions(db):
    """旧库单指标 mse 列 → 通用 score/metric_name；同时兼容最早的 filename 列名。"""
    cols = _table_columns(db, "submissions")
    if not cols or ("score" in cols and "mse" not in cols):
        return False
    if "score" not in cols and "mse" not in cols:
        return False  # 全新库已是新结构

    score_expr = ("COALESCE(score, mse)" if "score" in cols and "mse" in cols
                  else "score" if "score" in cols else "mse")
    result_col = ("result_filename" if "result_filename" in cols
                  else "filename" if "filename" in cols else None)

    def pick(name):
        return name if name in cols else None

    db.execute("ALTER TABLE submissions RENAME TO submissions_old")
    db.execute(SUBMISSIONS_DDL)
    db.execute(
        f"""
        INSERT INTO submissions
            (id, user_id, score, metric_name, result_path, result_filename,
             code_path, code_filename, created_at)
        SELECT id, user_id, {score_expr}, ?, ?,
               ?, ?, ?, ?
        FROM submissions_old
        """,
        (
            DEFAULT_METRIC["name"],
            pick("result_path"),
            result_col,
            pick("code_path"),
            pick("code_filename"),
            pick("created_at"),
        ),
    )
    db.execute("DROP TABLE submissions_old")
    db.commit()
    print("[迁移] submissions 表已升级为通用指标得分结构（mse -> score/metric_name）")
    return True


def init_db():
    """建表；自动迁移旧库；users 为空时从 Excel 导入学生；确保管理员存在。"""
    db = sqlite3.connect(DB_FILE)

    db.execute(
        """
        CREATE TABLE IF NOT EXISTS users (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            student_no    TEXT NOT NULL UNIQUE,   -- 学号（登录账号）
            name          TEXT NOT NULL,          -- 姓名
            gender        TEXT,                   -- 性别
            class_name    TEXT,                   -- 班级
            password_hash TEXT NOT NULL,          -- 密码哈希
            role          TEXT NOT NULL DEFAULT 'student',  -- student / admin
            created_at    TEXT NOT NULL
        )
        """
    )
    if not _table_columns(db, "submissions"):
        db.execute(SUBMISSIONS_DDL)
    db.execute(METRICS_DDL)
    db.execute(
        """
        CREATE TABLE IF NOT EXISTS settings (
            key   TEXT PRIMARY KEY,
            value TEXT
        )
        """
    )
    db.commit()

    _migrate_submissions(db)

    # 兼容旧版数据库：逐列迁移
    user_cols = _table_columns(db, "users")
    if "role" not in user_cols:
        db.execute("ALTER TABLE users ADD COLUMN role TEXT NOT NULL DEFAULT 'student'")
        db.commit()

    # ---- 无任何指标设定 -> 播种默认 MSE，保证系统开箱可用且评分口径不变
    if db.execute("SELECT COUNT(*) FROM metrics").fetchone()[0] == 0:
        now = datetime.now().strftime(DT_FMT)
        db.execute(
            "INSERT INTO metrics (name, task_type, method, direction, pred_column,"
            " true_column, description, enabled, created_at)"
            " VALUES (:name, :task_type, :method, :direction, :pred_column,"
            " :true_column, :description, 1, :created_at)",
            {**DEFAULT_METRIC, "created_at": now},
        )
        db.commit()
        print("[初始化] 已播种默认评价指标：均方误差 MSE（启用中）")

    # ---- 学生为空 -> 从 Excel 初始化
    if db.execute("SELECT COUNT(*) FROM users WHERE role = 'student'").fetchone()[0] == 0:
        if os.path.exists(STUDENT_FILE):
            df = pd.read_excel(STUDENT_FILE, dtype=str)
            now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            for _, row in df.iterrows():
                sno = str(row["学号"]).strip()
                db.execute(
                    "INSERT INTO users (student_no, name, gender, class_name,"
                    " password_hash, role, created_at) VALUES (?, ?, ?, ?, ?, 'student', ?)",
                    (
                        sno,
                        str(row["姓名"]).strip(),
                        str(row.get("性别", "")).strip(),
                        str(row.get("班级", "")).strip(),
                        generate_password_hash(sno),  # 初始密码 = 学号
                        now,
                    ),
                )
            db.commit()
            print(f"[初始化] 已从 Excel 导入 {len(df)} 名学生，账号=密码=学号")

    # ---- 确保管理员存在：admin / admin888
    admin = db.execute("SELECT id FROM users WHERE student_no = 'admin'").fetchone()
    if admin is None:
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        db.execute(
            "INSERT INTO users (student_no, name, gender, class_name,"
            " password_hash, role, created_at) VALUES ('admin', '系统管理员',"
            " '', '', ?, 'admin', ?)",
            (generate_password_hash("admin888"), now),
        )
        db.commit()
        print("[初始化] 已创建管理员账号：admin / admin888")

    db.close()

    # 确保上传/数据目录存在
    for d in (RESULT_DIR, CODE_DIR, DATA_DIR):
        os.makedirs(d, exist_ok=True)


# ---------------------------------------------------------------- 系统设置
def get_setting(key, default=None):
    row = get_db().execute("SELECT value FROM settings WHERE key = ?", (key,)).fetchone()
    return row["value"] if row is not None else default


def set_setting(key, value):
    db = get_db()
    db.execute(
        "INSERT INTO settings (key, value) VALUES (?, ?)"
        " ON CONFLICT(key) DO UPDATE SET value = excluded.value",
        (key, value),
    )
    db.commit()


DEADLINE_KEY = "submit_deadline"
DT_FMT = "%Y-%m-%d %H:%M:%S"


def get_deadline():
    """返回截止时间 datetime；未设置返回 None。"""
    raw = get_setting(DEADLINE_KEY)
    if not raw:
        return None
    try:
        return datetime.strptime(raw, DT_FMT)
    except ValueError:
        return None


def deadline_status(now=None):
    """
    返回截止状态字典：
    set(是否设置)、deadline(datetime或None)、expired(是否已截止)、remain(剩余timedelta或None)
    """
    now = now or datetime.now()
    dl = get_deadline()
    if dl is None:
        return {"set": False, "deadline": None, "expired": False, "remain": None}
    remain = dl - now
    return {"set": True, "deadline": dl, "expired": now >= dl,
            "remain": remain if remain.total_seconds() > 0 else timedelta(0)}


def fmt_remain(td):
    """将剩余时间格式化为 'x天x小时x分'。"""
    total_sec = int(td.total_seconds())
    days, rem = divmod(total_sec, 86400)
    hours, rem = divmod(rem, 3600)
    minutes = rem // 60
    parts = []
    if days:
        parts.append(f"{days} 天")
    if hours:
        parts.append(f"{hours} 小时")
    if minutes or not parts:
        parts.append(f"{minutes} 分")
    return " ".join(parts)


def delete_submission_files(sub):
    """删除一条提交记录对应的结果与代码文件（sqlite3.Row 或兼容映射）。"""
    for directory, key in ((RESULT_DIR, "result_path"), (CODE_DIR, "code_path")):
        name = sub[key]
        if name:
            try:
                os.remove(os.path.join(directory, name))
            except OSError:
                pass


# ---------------------------------------------------------------- 评价指标设定（管理员维护）
def _metric_view(row):
    """metrics 表行 -> 评分流程与模板统一使用的指标字典。"""
    method = row["method"] if row["method"] in METRIC_LIB else DEFAULT_METRIC["method"]
    meta = METRIC_LIB[method]
    direction = row["direction"] if row["direction"] in ("min", "max") else meta["direction"]
    task_type = row["task_type"] if row["task_type"] in TASK_TYPES else "regression"
    return {
        "id": row["id"],
        "name": row["name"],
        "task_type": task_type,
        "task_type_label": TASK_TYPES[task_type],
        "method": method,
        "method_label": meta["label"],
        "direction": direction,
        "better_text": "越小越好" if direction == "min" else "越大越好",
        "needs_numeric": meta["needs_numeric"],
        "decimals": meta["decimals"],
        "pred_column": str(row["pred_column"] or DEFAULT_METRIC["pred_column"]).strip(),
        "true_column": str(row["true_column"] or DEFAULT_METRIC["true_column"]).strip(),
        "description": row["description"] or meta["desc"],
        "enabled": bool(row["enabled"]),
    }


def get_active_metric():
    """当前生效的评价指标；管理员未设定时回退默认 MSE，保证系统开箱可用。"""
    row = get_db().execute(
        "SELECT * FROM metrics WHERE enabled = 1 ORDER BY id LIMIT 1").fetchone()
    if row is not None:
        return _metric_view(row)
    return _metric_view({**DEFAULT_METRIC, "id": None, "enabled": 1})


def list_metrics():
    rows = get_db().execute("SELECT * FROM metrics ORDER BY enabled DESC, id").fetchall()
    return [_metric_view(r) for r in rows]


def best_score_expr(m):
    """历史最佳得分的聚合表达式。"""
    return "MAX(score)" if m["direction"] == "max" else "MIN(score)"


def score_order_sql(m):
    return "DESC" if m["direction"] == "max" else "ASC"


def better_op_sql(m):
    return ">" if m["direction"] == "max" else "<"


def best_sort_key(m):
    """Python 侧排序键：最佳值优先，同分取更早达到者。"""
    sign = -1 if m["direction"] == "max" else 1
    return lambda score, ts: (sign * score, ts)


# ---------------------------------------------------------------- 标准答案
_answer_cache = None  # (文件路径, mtime, DataFrame)


def _decode_bytes(raw):
    """CSV 字节解码，编码自适应 UTF-8-SIG / UTF-8 / GBK。"""
    for enc in ("utf-8-sig", "utf-8", "gbk"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    raise ValueError("文件编码无法识别（需 UTF-8 或 GBK）")


def _generic_columns(ncols):
    """无表头文件的展示列名：第 1 列为样本id，其后为 col2、col3……"""
    return ["sample_id"] + [f"col{i}" for i in range(2, ncols + 1)]


def _named_columns(names):
    """表头列名去空、去重（重复列追加序号，避免按列名取值时歧义）。"""
    fallback = _generic_columns(len(names))
    out, seen = [], {}
    for i, name in enumerate(names):
        col = name if str(name).strip() else fallback[i]
        col = str(col).strip()
        seen[col] = seen.get(col, 0) + 1
        out.append(col if seen[col] == 1 else f"{col}_{seen[col]}")
    return out


def _parse_answer_bytes(raw):
    """
    解析标准答案 CSV（第 1 列样本id，其余列为真实值），编码自适应 UTF-8/GBK。
    首行若被识别为表头则自动剔除，并以其列名作为真实值列名（供指标按列名设定）。
    """
    text = _decode_bytes(raw)
    ans = pd.read_csv(io.StringIO(text), header=None, dtype=str, skip_blank_lines=True)
    if ans.shape[1] < 2:
        raise ValueError(f"标准答案文件至少需要两列（样本id、真实值），当前检测到 {ans.shape[1]} 列")
    has_header, names = looks_like_header(ans)
    if has_header:
        ans = ans.iloc[1:].reset_index(drop=True)
        ans.columns = _named_columns(names)
    else:
        ans.columns = _generic_columns(ans.shape[1])
    ans = ans.fillna("")
    for c in ans.columns:
        ans[c] = ans[c].astype(str).str.strip()
    if (ans.iloc[:, 0] == "").any():
        raise ValueError("标准答案文件存在空的样本id")
    # 允许重复样本id（测试数据本身可能包含重复批次），所有行均保留参与评分
    return ans


def read_table_head(path, nrows=40):
    """读取数据文件前若干行（按无表头方式，全部转字符串），用于识别列结构。"""
    ext = os.path.splitext(path)[1].lower()
    if ext == ".csv":
        with open(path, "rb") as f:
            text = _decode_bytes(f.read())
        df = pd.read_csv(io.StringIO(text), header=None, dtype=str, nrows=nrows)
    else:
        df = pd.read_excel(path, header=None, dtype=str, nrows=nrows)
    return df


def _all_non_numeric(values):
    for s in values:
        try:
            float(s)
        except (ValueError, TypeError):
            continue
        return False
    return True


def looks_like_header(df):
    """
    首行判定为表头的条件：每一列的首行值都不是数字，且该值不在本列下方数据中出现
    （类别标签类答案的首行同样是数据，靠"下方重复"排除）。
    返回（是否表头，列名列表）。
    """
    if len(df) == 0:
        return False, []
    names = ["" if v is None else str(v).strip() for v in df.iloc[0].tolist()]
    if not names or not _all_non_numeric(names):
        return False, names
    if len(df) < 2:
        return True, names
    below = df.iloc[1:].astype(str)
    for j, name in enumerate(names):
        if name in set(below.iloc[:, j].str.strip()):
            return False, names
    return True, names


# 提交样例列结构缓存：(路径, mtime, 布局字典)
_sample_cache = None


def get_sample_layout():
    """
    结果文件的列格式基准：读取管理员配置的「提交样例」文件，确定应有列数与列名。
    样例缺省时回退两列（样本id、预测值）。
    """
    global _sample_cache
    path, display = get_data_file("sample")
    if path is None:
        return {"ncols": 2, "header": None, "columns": ["sample_id", "col2"],
                "source": None, "note": "提交样例未配置，按两列（样本id、预测值）校验"}
    try:
        mtime = os.path.getmtime(path)
    except OSError:
        mtime = None
    if _sample_cache is None or _sample_cache[0] != path or _sample_cache[1] != mtime:
        try:
            df = read_table_head(path)
            has_header, names = looks_like_header(df)
            ncols = df.shape[1]
            layout = {
                "ncols": ncols,
                "header": names if has_header else None,
                "columns": names if has_header else _generic_columns(ncols),
                "source": display,
                "note": "",
            }
        except Exception as e:
            layout = {"ncols": 2, "header": None, "columns": ["sample_id", "col2"],
                      "source": display, "note": f"提交样例解析失败（{e}），按两列校验"}
        _sample_cache = (path, mtime, layout)
    return _sample_cache[2]


def answer_layout():
    """标准答案列结构（供后台展示，便于管理员设定计分列）；与实际解析口径一致。"""
    path = get_answer_path()
    try:
        with open(path, "rb") as f:
            ans = _parse_answer_bytes(f.read())
        return {"ncols": ans.shape[1], "columns": list(ans.columns),
                "source": os.path.basename(path), "n": len(ans)}
    except Exception:
        return {"ncols": 2, "columns": ["sample_id", "col2"],
                "source": os.path.basename(path), "n": 0}


def resolve_column(spec, columns, label):
    """把指标设定的计分列（1 起的序号或列名）解析为（列序号, 列名）。"""
    spec = str(spec or "").strip() or "2"
    if spec.isdigit():
        idx = int(spec)
        if 1 <= idx <= len(columns):
            return idx, columns[idx - 1]
        raise SubmitError(
            f"{label}中设定的计分列序号 {spec} 超出范围（{label}共 {len(columns)} 列）。")
    for i, c in enumerate(columns):
        if str(c).strip().lower() == spec.strip().lower():
            return i + 1, c
    raise SubmitError(
        f"{label}中找不到设定的计分列「{spec}」，可用列：{', '.join(str(c) for c in columns)}。")


def metric_setup_warnings():
    """
    对照当前启用指标检查数据文件的列设定，返回需要提醒管理员的提示文本。
    仅提示、不拦截：切换赛题类型时允许先换数据文件、再到「评价指标」改设定。
    """
    m = get_active_metric()
    warns = []
    try:
        resolve_column(m["pred_column"], get_sample_layout()["columns"], "提交结果")
    except SubmitError as e:
        warns.append(str(e))
    try:
        true_idx, true_name = resolve_column(
            m["true_column"], answer_layout()["columns"], "标准答案")
    except SubmitError as e:
        warns.append(str(e))
        return warns
    if m["needs_numeric"]:
        try:
            with open(get_answer_path(), "rb") as fh:
                ans = _parse_answer_bytes(fh.read())
            if true_idx == 1:
                warns.append(f"启用指标「{m['name']}」把标准答案第 1 列（样本id）设为真实值列。")
            else:
                col = pd.to_numeric(ans[ans.columns[true_idx - 1]], errors="coerce")
                if col.isna().any():
                    warns.append(
                        f"启用指标「{m['name']}」要求真实值为数字，"
                        f"标准答案列「{true_name}」存在非数字取值。")
        except Exception:
            pass
    return warns


def get_answer_path():
    """评分标准答案路径：优先管理员配置的实际数据，否则用项目默认 ansA.csv。"""
    filename = get_setting(ANSWER_KEY)
    if filename:
        p = os.path.join(DATA_DIR, filename)
        if os.path.exists(p):
            return p
    p = os.path.join(DATA_DIR, DATA_FILE_MAP["answer"][1])
    if os.path.exists(p):
        return p
    return ANSWER_FILE


def load_answer():
    """读取标准答案，进程内缓存；管理员重新上传实际数据后自动失效重载。"""
    global _answer_cache
    path = get_answer_path()
    try:
        mtime = os.path.getmtime(path)
    except OSError:
        mtime = None
    if _answer_cache is None or _answer_cache[0] != path or _answer_cache[1] != mtime:
        with open(path, "rb") as f:
            raw = f.read()
        _answer_cache = (path, mtime, _parse_answer_bytes(raw))
    return _answer_cache[2]


# ---------------------------------------------------------------- 上传文件解析与评分
class SubmitError(ValueError):
    """提交文件不合法时抛出，消息直接展示给用户。"""


def parse_submission_bytes(raw, metric=None, layout=None):
    """
    解析学生上传的结果 CSV 字节，返回 DataFrame[sample_id(str), y_pred]。

    列格式以管理员配置的「提交样例」为准（列数、列名），参与计分的预测列
    由当前启用指标的 pred_column 决定；回归类指标额外要求该列取值为数字，
    分类 / 聚类类指标则按类别标签（字符串）处理。
    """
    metric = metric or get_active_metric()
    layout = layout or get_sample_layout()

    if not raw.strip():
        raise SubmitError("结果文件内容为空，请上传有效的 CSV 文件。")
    text = _decode_bytes(raw) if isinstance(raw, bytes) else str(raw)

    df = pd.read_csv(io.StringIO(text), header=None, dtype=str, skip_blank_lines=True)
    ncols = layout["ncols"]
    if df.shape[1] != ncols:
        ref = f"（提交样例 {layout['source']} 为 {ncols} 列）" if layout["source"] else ""
        raise SubmitError(
            f"结果文件应包含 {ncols} 列{ref}，当前检测到 {df.shape[1]} 列。")

    columns = layout["columns"][:ncols]
    pred_idx, pred_name = resolve_column(metric["pred_column"], columns, "提交结果")
    id_idx, id_name = 1, columns[0]

    # 表头行识别：与样例表头逐列一致，或（样例无表头且指标要求数字时）首行预测列非数字
    header_row = False
    first_pred = str(df.iloc[0, pred_idx - 1]).strip()
    if layout["header"]:
        first_cells = [str(v).strip().lower() for v in df.iloc[0].tolist()]
        header_row = first_cells == [str(c).strip().lower() for c in layout["header"]]
    elif metric["needs_numeric"]:
        try:
            float(first_pred)
        except (ValueError, TypeError):
            header_row = True
    if header_row:
        df = df.iloc[1:].reset_index(drop=True)

    if df.empty:
        raise SubmitError("结果文件没有数据行（仅含表头），请检查后重新上传。")

    out = pd.DataFrame({
        "sample_id": df.iloc[:, id_idx - 1].astype(str).str.strip(),
        "y_pred": df.iloc[:, pred_idx - 1].astype(str).str.strip(),
    })
    out["y_pred"] = out["y_pred"].where(
        ~out["y_pred"].str.lower().isin(["nan", "none", "null"]), "")

    if (out["sample_id"] == "").any():
        raise SubmitError(f"存在空的样本id（第 {id_name} 列），请补全后重新上传。")
    empty = out[out["y_pred"].isna() | (out["y_pred"] == "")]
    if len(empty) > 0:
        raise SubmitError(
            f"存在空的预测值（第 {pred_idx} 列「{pred_name}」，"
            f"如批次 {', '.join(empty['sample_id'].head(3).tolist())}），请补全后重新上传。")
    # 允许重复样本id（测试数据本身可能包含重复批次），所有行均保留参与评分

    if metric["needs_numeric"]:
        num = pd.to_numeric(out["y_pred"], errors="coerce")
        if num.isna().any():
            bad = out.loc[num.isna(), "sample_id"].head(3).tolist()
            raise SubmitError(
                f"以下批次的预测值（第 {pred_idx} 列「{pred_name}」）不是合法数字："
                f"{', '.join(bad)}。")
        out["y_pred"] = num

    out.attrs["pred_name"] = pred_name
    out.attrs["pred_index"] = pred_idx
    return out


def parse_submission(file_storage, metric=None, layout=None):
    """FileStorage 版包装：读取字节后按指标与提交样例解析。"""
    return parse_submission_bytes(file_storage.read(), metric, layout)


def align_to_answer(sub_df, metric, ans=None):
    """
    与标准答案对齐：行数一致 → 样本id集合一致 → 各id出现次数一致 →
    按"同id第 k 次出现"逐行配对（避免重复id直接 merge 产生笛卡尔积）。
    返回 (真实值列表, 预测值列表, 配对行数)。
    """
    ans = ans if ans is not None else load_answer()
    n_total = len(ans)
    id_col = ans.columns[0]

    if len(sub_df) != n_total:
        raise SubmitError(
            f"结果文件应包含 {n_total} 行（与测试数据批次一一对应），"
            f"当前检测到 {len(sub_df)} 行。")

    ans_ids = set(ans[id_col])
    sub_ids = set(sub_df["sample_id"])
    missing = ans_ids - sub_ids
    extra = sub_ids - ans_ids
    if missing or extra:
        msgs = []
        if missing:
            show = sorted(missing)[:3]
            msgs.append(f"缺少 {len(missing)} 个批次（如 {', '.join(show)}）")
        if extra:
            show = sorted(extra)[:3]
            msgs.append(f"含 {len(extra)} 个非法批次id（如 {', '.join(show)}）")
        raise SubmitError("批次与标准答案不一致：" + "；".join(msgs) + "。")

    # 各 id 出现次数必须一致（允许重复id的场景）
    ans_cnt = ans[id_col].value_counts()
    sub_cnt = sub_df["sample_id"].value_counts()
    cnt_diff = (ans_cnt - sub_cnt.reindex(ans_cnt.index, fill_value=0)).abs()
    if (cnt_diff > 0).any():
        bad = cnt_diff[cnt_diff > 0].index[:3].tolist()
        raise SubmitError(
            f"以下批次的提交行数与测试数据不一致：{', '.join(bad)}。"
            f"重复批次需按测试数据中的出现次数逐行提交。")

    true_idx, _true_name = resolve_column(metric["true_column"], list(ans.columns), "标准答案")
    # 答案列改用内部列名，避免与提交列（y_pred）或配对辅助列（_occ）重名
    internal = [f"_true{i + 1}" for i in range(ans.shape[1])]
    a = ans.copy()
    a.columns = ["sample_id"] + internal[1:]
    true_key = "sample_id" if true_idx == 1 else internal[true_idx - 1]
    s = sub_df.copy()
    a["_occ"] = a.groupby("sample_id").cumcount()
    s["_occ"] = s.groupby("sample_id").cumcount()
    merged = a.merge(s, on=["sample_id", "_occ"], how="left")
    y_true = merged[true_key].tolist()
    y_pred = merged["y_pred"].tolist()
    return y_true, y_pred, len(merged)


def score_submission(sub_df, metric=None):
    """
    按当前启用的评价指标计算本次提交得分（指标设定由后台读取）。
    返回 {"score", "metric_name", "method", "direction", "n"}。
    """
    metric = metric or get_active_metric()
    y_true, y_pred, n = align_to_answer(sub_df, metric)
    meta = METRIC_LIB[metric["method"]]
    try:
        value = float(meta["fn"](y_true, y_pred))
    except MetricError as e:
        raise SubmitError(f"指标「{metric['name']}」计算失败：{e}")
    except SubmitError:
        raise
    except Exception as e:
        raise SubmitError(f"指标「{metric['name']}」计算失败：{e}")
    if math.isnan(value) or math.isinf(value):
        raise SubmitError(f"指标「{metric['name']}」计算结果无效，请检查提交内容与指标设定。")
    return {"score": value, "metric_name": metric["name"], "method": metric["method"],
            "direction": metric["direction"], "n": n}


def rescore_from_bytes(raw, metric=None, layout=None):
    """解析 + 评分一步到位（学生提交与后台历史重算共用）。"""
    metric = metric or get_active_metric()
    sub_df = parse_submission_bytes(raw, metric, layout)
    result = score_submission(sub_df, metric)
    result["pred_name"] = sub_df.attrs.get("pred_name", "")
    result["pred_index"] = sub_df.attrs.get("pred_index", 0)
    return result


def save_upload(file_storage, target_dir, submission_id):
    """
    将上传文件安全保存到 target_dir，返回（落盘basename, 原始文件名）。
    以 submission_id 作前缀，避免同名覆盖。
    注意：解析阶段可能已读过流（指针在末尾），保存前必须 seek(0)。
    """
    original = file_storage.filename
    safe = secure_filename(original) or "file"
    stored_name = f"{submission_id}_{safe}"
    file_storage.stream.seek(0)
    file_storage.save(os.path.join(target_dir, stored_name))
    return stored_name, original


# ---------------------------------------------------------------- 统一提交处理（PC 端 / 微信端共用）
def handle_submit(user, file_storage, code_file=None):
    """
    校验、按启用指标评分、落盘一条龙。PC 端与微信端共用，避免两端逻辑漂移。

    file_storage : FileStorage（预测结果 CSV）
    code_file    : FileStorage 或 None
    返回 result 字典（score / metric_name / sid / time / result_name / code_name）；
    任何不合法情况抛 SubmitError，消息直接展示给用户。
    """
    has_code = code_file is not None and code_file.filename.strip() != ""
    dstat = deadline_status()

    if dstat["expired"]:
        raise SubmitError(
            f"提交已于 {dstat['deadline'].strftime('%Y-%m-%d %H:%M:%S')} 截止，"
            "无法再提交，请联系任课老师。")
    if file_storage is None or file_storage.filename == "":
        raise SubmitError("请先提供预测结果 CSV 文件。")
    if not file_storage.filename.lower().endswith(".csv"):
        raise SubmitError("预测结果仅支持 CSV 格式文件，请核对后重新上传。")
    if has_code and os.path.splitext(code_file.filename)[1].lower() not in CODE_ALLOWED_EXT:
        ext = os.path.splitext(code_file.filename)[1].lower()
        raise SubmitError(
            f"不支持的代码文件类型 {ext or '（无扩展名）'}，允许："
            ".py/.ipynb/.zip/.rar/.7z/.tar.gz/.txt/.pdf/.doc/.docx 等。")

    metric = get_active_metric()
    layout = get_sample_layout()
    raw = file_storage.read()
    scored = rescore_from_bytes(raw, metric, layout)

    db = get_db()
    now = datetime.now().strftime(DT_FMT)

    # 先建提交记录拿到 id，再以 id 前缀保存文件，最后回写路径
    cur = db.execute(
        "INSERT INTO submissions (user_id, score, metric_name, created_at)"
        " VALUES (?, ?, ?, ?)",
        (user["id"], scored["score"], metric["name"], now),
    )
    db.commit()
    sid = cur.lastrowid

    stored_files = []
    try:
        file_storage.stream.seek(0)
        result_stored, result_orig = save_upload(file_storage, RESULT_DIR, sid)
        stored_files.append(os.path.join(RESULT_DIR, result_stored))

        code_stored = code_orig = None
        if has_code:
            code_stored, code_orig = save_upload(code_file, CODE_DIR, sid)
            stored_files.append(os.path.join(CODE_DIR, code_stored))

        db.execute(
            "UPDATE submissions SET result_path = ?, result_filename = ?,"
            " code_path = ?, code_filename = ? WHERE id = ?",
            (result_stored, result_orig, code_stored, code_orig, sid),
        )
        db.commit()
    except Exception:
        # 文件落盘失败：清理已写文件并回滚提交记录
        for p in stored_files:
            try:
                os.remove(p)
            except OSError:
                pass
        db.execute("DELETE FROM submissions WHERE id = ?", (sid,))
        db.commit()
        raise

    return {"score": scored["score"], "metric_name": metric["name"], "sid": sid,
            "time": now, "result_name": result_orig,
            "code_name": code_orig if has_code else None}


def student_overview(user):
    """返回某学生的个人战绩概览（历史最佳 / 提交次数 / 排名 / 最近 5 次）。"""
    m = get_active_metric()
    db = get_db()
    row = db.execute(
        f"SELECT {best_score_expr(m)} AS best_score, COUNT(*) AS n_submit"
        " FROM submissions WHERE user_id = ?",
        (user["id"],),
    ).fetchone()
    best_score = row["best_score"]
    my_rank = None
    if best_score is not None:
        my_rank = db.execute(
            f"""
            SELECT COUNT(*) + 1 AS rank FROM (
                SELECT user_id, {best_score_expr(m)} AS best FROM submissions GROUP BY user_id
            ) WHERE best {better_op_sql(m)} ?
            """,
            (best_score,),
        ).fetchone()["rank"]

    recent = db.execute(
        "SELECT * FROM submissions WHERE user_id = ? ORDER BY id DESC LIMIT 5",
        (user["id"],),
    ).fetchall()
    return {
        "user": user,
        "best_score": best_score,
        "n_submit": row["n_submit"],
        "my_rank": my_rank,
        "recent": recent,
        "dstat": deadline_status(),
    }


# ---------------------------------------------------------------- 权限装饰器
def login_required(view):
    @wraps(view)
    def wrapped(*args, **kwargs):
        if "user_id" not in session:
            flash("请先登录。", "warn")
            return redirect(url_for("login", next=request.path))
        return view(*args, **kwargs)
    return wrapped


def admin_required(view):
    @wraps(view)
    def wrapped(*args, **kwargs):
        if "user_id" not in session:
            flash("请先登录管理员账号。", "warn")
            return redirect(url_for("login", next=request.path))
        if session.get("role") != "admin":
            abort(403)
        return view(*args, **kwargs)
    return wrapped


# ---------------------------------------------------------------- 路由：认证
@app.route("/login", methods=["GET", "POST"])
def login():
    if "user_id" in session:
        if session.get("role") == "admin":
            return redirect(url_for("admin_index"))
        return redirect(url_for("upload_page"))

    if request.method == "POST":
        sno = request.form.get("student_no", "").strip()
        pwd = request.form.get("password", "")
        user = get_db().execute("SELECT * FROM users WHERE student_no = ?", (sno,)).fetchone()
        if user is None or not check_password_hash(user["password_hash"], pwd):
            flash("账号或密码错误，请重试。", "error")
        else:
            session.clear()
            session["user_id"] = user["id"]
            session["role"] = user["role"]
            session["name"] = user["name"]
            if user["role"] == "admin":
                flash(f"欢迎管理员 {user['name']}！", "success")
                return redirect(url_for("admin_index"))
            flash(f"欢迎你，{user['name']}（{user['student_no']}）！", "success")
            nxt = request.args.get("next")
            return redirect(nxt or url_for("upload_page"))
    return render_template("login.html")


@app.route("/logout")
def logout():
    session.clear()
    flash("已安全退出登录。", "success")
    return redirect(url_for("login"))


@app.route("/change_password", methods=["GET", "POST"])
@login_required
def change_password():
    db = get_db()
    user = db.execute("SELECT * FROM users WHERE id = ?", (session["user_id"],)).fetchone()
    if request.method == "POST":
        old_pwd = request.form.get("old_password", "")
        new_pwd = request.form.get("new_password", "")
        confirm = request.form.get("confirm_password", "")

        if not check_password_hash(user["password_hash"], old_pwd):
            flash("原密码不正确。", "error")
        elif len(new_pwd) < 6:
            flash("新密码长度至少 6 位。", "error")
        elif new_pwd != confirm:
            flash("两次输入的新密码不一致。", "error")
        elif new_pwd == old_pwd:
            flash("新密码不能与原密码相同。", "error")
        else:
            db.execute("UPDATE users SET password_hash = ? WHERE id = ?",
                       (generate_password_hash(new_pwd), user["id"]))
            db.commit()
            flash("密码修改成功，下次登录请使用新密码。", "success")
            return redirect(url_for("upload_page"))
    return render_template("change_password.html", user=user)


# ---------------------------------------------------------------- 路由：学生上传评分
@app.route("/", methods=["GET", "POST"])
@login_required
def upload_page():
    if session.get("role") == "admin":
        return redirect(url_for("admin_index"))

    user = get_db().execute(
        "SELECT * FROM users WHERE id = ?", (session["user_id"],)).fetchone()
    result = None  # 本次评分结果

    if request.method == "POST":
        try:
            result = handle_submit(user, request.files.get("file"),
                                   request.files.get("code_file"))
            flash("提交成功，成绩已评定！", "success")
        except SubmitError as e:
            flash(str(e), "error")
        except Exception as e:
            flash(f"提交失败：{e}", "error")

    ov = student_overview(user)
    return render_template("upload.html", result=result, layout=get_sample_layout(), **ov)


# ---------------------------------------------------------------- 路由：我的提交
@app.route("/my")
@login_required
def my_submissions():
    if session.get("role") == "admin":
        return redirect(url_for("admin_index"))
    db = get_db()
    user = db.execute("SELECT * FROM users WHERE id = ?", (session["user_id"],)).fetchone()
    rows = db.execute(
        "SELECT * FROM submissions WHERE user_id = ? ORDER BY id DESC", (user["id"],)
    ).fetchall()
    scores = [r["score"] for r in rows if r["score"] is not None]
    best_score = None
    if scores:
        best_score = max(scores) if get_active_metric()["direction"] == "max" else min(scores)
    return render_template("my_submissions.html", user=user, rows=rows, best_score=best_score)


# ---------------------------------------------------------------- 路由：排行榜
@app.route("/leaderboard")
@login_required
def leaderboard():
    if session.get("role") == "admin":
        return redirect(url_for("admin_ranking"))
    m = get_active_metric()
    db = get_db()
    rows = db.execute(
        f"""
        SELECT u.student_no, u.name, u.class_name,
               {best_score_expr(m)} AS best_score,
               COUNT(s.id)       AS n_submit,
               MAX(s.created_at) AS last_time
        FROM users u
        JOIN submissions s ON s.user_id = u.id
        GROUP BY u.id
        ORDER BY best_score {score_order_sql(m)}, last_time ASC
        """
    ).fetchall()

    total_users = db.execute(
        "SELECT COUNT(*) AS c FROM users WHERE role = 'student'").fetchone()["c"]
    my_best = db.execute(
        f"SELECT {best_score_expr(m)} AS best FROM submissions WHERE user_id = ?",
        (session["user_id"],),
    ).fetchone()["best"]

    return render_template("leaderboard.html", rows=rows, total_users=total_users,
                           my_best=my_best)


# ---------------------------------------------------------------- 路由：赛题与数据
def get_topic_settings():
    """读取管理员维护的赛题内容（未设置的字段使用默认值）。"""
    return {k: get_setting(k, DEFAULT_TOPIC.get(k, "")) for k in TOPIC_KEYS}


def get_data_file(key):
    """
    返回赛题数据文件（文件路径, 文件名/显示名）。
    优先按管理员配置的文件名在 data 目录查找，不存在时回退默认文件。
    """
    setting_key, default, label = DATA_FILE_MAP[key]
    filename = get_setting(setting_key)
    if filename:
        p = os.path.join(DATA_DIR, filename)
        if os.path.exists(p):
            return p, filename
    # 回退默认（优先 data 目录，再回退项目根目录）
    p = os.path.join(DATA_DIR, default)
    if os.path.exists(p):
        return p, default
    p = os.path.join(BASE_DIR, default)
    if os.path.exists(p):
        return p, default
    return None, None


@app.route("/topic")
@login_required
def topic_page():
    t = get_topic_settings()
    files = {k: get_data_file(k)[1] for k in ("train", "testA", "sample")}
    return render_template("topic.html", t=t, files=files)


@app.route("/download_data/<key>")
@login_required
def download_data(key):
    """下载竞赛数据文件（管理员上传版，未上传时用项目默认文件）。"""
    if key not in ("train", "testA", "sample"):
        abort(404)
    path, display = get_data_file(key)
    if not path:
        abort(404)
    return send_file(path, as_attachment=True, download_name=display)


# ---------------------------------------------------------------- 路由：下载（本人）
@app.route("/download/<int:sid>/<ftype>")
@login_required
def download_file(sid, ftype):
    if ftype not in ("result", "code"):
        abort(404)
    db = get_db()
    sub = db.execute("SELECT * FROM submissions WHERE id = ?", (sid,)).fetchone()
    if sub is None:
        abort(404)
    # 仅文件所有者或管理员可下载
    if sub["user_id"] != session["user_id"] and session.get("role") != "admin":
        abort(403)

    if ftype == "result":
        stored, original, directory = sub["result_path"], sub["result_filename"], RESULT_DIR
    else:
        stored, original, directory = sub["code_path"], sub["code_filename"], CODE_DIR

    if not stored or not os.path.exists(os.path.join(directory, stored)):
        abort(404)
    return send_from_directory(directory, stored, as_attachment=True,
                               download_name=original)


# ================================================================ 管理员后台
@app.route("/admin")
@admin_required
def admin_index():
    db = get_db()
    stats = {
        "students": db.execute(
            "SELECT COUNT(*) AS c FROM users WHERE role = 'student'").fetchone()["c"],
        "submitted": db.execute(
            "SELECT COUNT(DISTINCT user_id) AS c FROM submissions").fetchone()["c"],
        "submissions": db.execute(
            "SELECT COUNT(*) AS c FROM submissions").fetchone()["c"],
        "codes": db.execute(
            "SELECT COUNT(*) AS c FROM submissions WHERE code_path IS NOT NULL"
        ).fetchone()["c"],
    }
    latest = db.execute(
        """
        SELECT s.*, u.student_no, u.name FROM submissions s
        JOIN users u ON u.id = s.user_id
        ORDER BY s.id DESC LIMIT 8
        """
    ).fetchall()
    return render_template("admin/index.html", stats=stats, latest=latest)


@app.route("/admin/users")
@admin_required
def admin_users():
    q = request.args.get("q", "").strip()
    m = get_active_metric()
    sql = (
        f"""
        SELECT u.*,
               (SELECT COUNT(*) FROM submissions s WHERE s.user_id = u.id) AS n_submit,
               (SELECT {best_score_expr(m)}
                  FROM submissions s WHERE s.user_id = u.id) AS best_score
        FROM users u WHERE u.role = 'student'
        """
    )
    params = ()
    if q:
        sql += " AND (u.student_no LIKE ? OR u.name LIKE ? OR u.class_name LIKE ?)"
        kw = f"%{q}%"
        params = (kw, kw, kw)
    sql += " ORDER BY u.student_no"
    rows = get_db().execute(sql, params).fetchall()
    return render_template("admin/users.html", rows=rows, q=q)


@app.route("/admin/users/<int:uid>/reset", methods=["POST"])
@admin_required
def admin_reset_password(uid):
    db = get_db()
    user = db.execute("SELECT * FROM users WHERE id = ? AND role = 'student'",
                      (uid,)).fetchone()
    if user is None:
        abort(404)
    db.execute("UPDATE users SET password_hash = ? WHERE id = ?",
               (generate_password_hash(user["student_no"]), uid))
    db.commit()
    flash(f"已将 {user['name']}（{user['student_no']}）的密码重置为学号。", "success")
    return redirect(url_for("admin_users"))


@app.route("/admin/users/<int:uid>/delete", methods=["POST"])
@admin_required
def admin_delete_user(uid):
    db = get_db()
    user = db.execute("SELECT * FROM users WHERE id = ? AND role = 'student'",
                      (uid,)).fetchone()
    if user is None:
        abort(404)

    # 先删除该用户的全部落盘文件，再删记录与账号
    subs = db.execute("SELECT * FROM submissions WHERE user_id = ?", (uid,)).fetchall()
    for s in subs:
        for directory, key in ((RESULT_DIR, "result_path"), (CODE_DIR, "code_path")):
            name = s[key]
            if name:
                try:
                    os.remove(os.path.join(directory, name))
                except OSError:
                    pass
    db.execute("DELETE FROM submissions WHERE user_id = ?", (uid,))
    db.execute("DELETE FROM users WHERE id = ?", (uid,))
    db.commit()
    flash(f"已删除用户 {user['name']}（{user['student_no']}）及其全部提交记录。", "success")
    return redirect(url_for("admin_users"))


@app.route("/admin/ranking")
@admin_required
def admin_ranking():
    m = get_active_metric()
    rows = get_db().execute(
        f"""
        SELECT u.student_no, u.name, u.class_name,
               {best_score_expr(m)} AS best_score,
               COUNT(s.id)       AS n_submit,
               MAX(s.created_at) AS last_time
        FROM users u
        JOIN submissions s ON s.user_id = u.id
        GROUP BY u.id
        ORDER BY best_score {score_order_sql(m)}, last_time ASC
        """
    ).fetchall()
    total_students = get_db().execute(
        "SELECT COUNT(*) AS c FROM users WHERE role = 'student'").fetchone()["c"]
    return render_template("admin/ranking.html", rows=rows, total_students=total_students)


@app.route("/admin/submissions")
@admin_required
def admin_submissions():
    q = request.args.get("q", "").strip()
    sql = (
        """
        SELECT s.*, u.student_no, u.name, u.class_name
        FROM submissions s
        JOIN users u ON u.id = s.user_id
        WHERE 1 = 1
        """
    )
    params = []
    if q:
        sql += (" AND (u.student_no LIKE ? OR u.name LIKE ?"
                " OR s.result_filename LIKE ? OR s.code_filename LIKE ?)")
        kw = f"%{q}%"
        params.extend([kw, kw, kw, kw])
    sql += " ORDER BY s.id DESC"
    rows = get_db().execute(sql, params).fetchall()
    return render_template("admin/submissions.html", rows=rows, q=q)


@app.route("/admin/submission/<int:sid>/<ftype>")
@admin_required
def admin_download(sid, ftype):
    if ftype not in ("result", "code"):
        abort(404)
    sub = get_db().execute("SELECT * FROM submissions WHERE id = ?", (sid,)).fetchone()
    if sub is None:
        abort(404)
    if ftype == "result":
        stored, original, directory = sub["result_path"], sub["result_filename"], RESULT_DIR
    else:
        stored, original, directory = sub["code_path"], sub["code_filename"], CODE_DIR
    if not stored or not os.path.exists(os.path.join(directory, stored)):
        abort(404)
    return send_from_directory(directory, stored, as_attachment=True,
                               download_name=original)


# ---------------------------------------------------------------- 系统设置（截止时间）
@app.route("/admin/settings", methods=["GET", "POST"])
@admin_required
def admin_settings():
    if request.method == "POST":
        action = request.form.get("action", "set")
        if action == "clear":
            set_setting(DEADLINE_KEY, "")
            flash("已取消提交截止时间，学生现在可以随时提交。", "success")
        else:
            deadline_raw = request.form.get("deadline", "").strip()
            if not deadline_raw:
                flash("请选择截止时间，或点击“取消截止时间”。", "error")
            else:
                try:
                    # datetime-local 提交格式：2026-10-20T18:00
                    dl = datetime.strptime(deadline_raw, "%Y-%m-%dT%H:%M")
                except ValueError:
                    flash("时间格式不正确，请重新选择。", "error")
                else:
                    set_setting(DEADLINE_KEY, dl.strftime(DT_FMT))
                    flash(f"提交截止时间已设置为 {dl.strftime(DT_FMT)}。", "success")
        return redirect(url_for("admin_settings"))

    return render_template("admin/settings.html", dstat=deadline_status())


# ---------------------------------------------------------------- 赛题与数据管理
@app.route("/admin/topic", methods=["GET", "POST"])
@admin_required
def admin_topic():
    if request.method == "POST":
        action = request.form.get("action", "save_text")

        if action == "save_text":
            if not request.form.get("topic_title", "").strip():
                flash("竞赛题目不能为空。", "error")
            else:
                for k in TOPIC_KEYS:
                    set_setting(k, request.form.get(k, "").strip())
                flash("赛题内容已保存，学生端“赛题与数据”页面已同步更新。", "success")
            return redirect(url_for("admin_topic"))

        if action == "save_files":
            saved, errors = [], []
            for dkey, (setting_key, default, label) in DATA_FILE_MAP.items():
                filename = request.form.get(f"file_{dkey}", "").strip()
                filename = os.path.basename(filename)  # 防路径穿越
                if not filename:
                    continue
                ext = os.path.splitext(filename)[1].lower()
                if dkey == "answer":
                    # 实际数据是评分依据：必须为两列（样本id,收率）CSV
                    if ext != ".csv":
                        errors.append(f"{label}：仅支持 CSV 文件（评分标准答案必须为两列 CSV）。")
                        continue
                else:
                    # 训练/测试/提交样例：支持 CSV 或 Excel
                    if ext not in (".csv", ".xlsx", ".xls"):
                        errors.append(f"{label}：仅支持 CSV 或 Excel（.xlsx / .xls）文件。")
                        continue
                p = os.path.join(DATA_DIR, filename)
                if not os.path.exists(p):
                    errors.append(f"{label}：data 目录下不存在文件 {filename}，"
                                  f"请先将文件放入 data 目录。")
                    continue
                if dkey == "answer":
                    # 实际数据是评分依据：须可解析（≥2 列、样本id 非空）
                    try:
                        with open(p, "rb") as fh:
                            ans = _parse_answer_bytes(fh.read())
                    except ValueError as e:
                        errors.append(f"{label}：{e}")
                        continue
                    except Exception:
                        errors.append(f"{label}：解析失败，请确认为（样本id,真实值）CSV。")
                        continue
                    set_setting(ANSWER_KEY, filename)
                    saved.append(f"{label}（{filename}，{len(ans)} 个批次）")
                else:
                    set_setting(setting_key, filename)
                    saved.append(f"{label}（{filename}）")
            for e in errors:
                flash(e, "error")
            if saved:
                flash("已保存数据文件配置：" + "、".join(saved) + "。", "success")
                for w in metric_setup_warnings():
                    flash(w + "（不会拦截保存，请到「评价指标」核对设定）", "warn")
            if not saved and not errors:
                flash("未填写任何文件名，配置未变更。", "error")
            return redirect(url_for("admin_topic"))

    t = get_topic_settings()
    files = {}
    for k, (setting_key, canon, label) in DATA_FILE_MAP.items():
        filename = get_setting(setting_key) or canon
        exists = os.path.exists(os.path.join(DATA_DIR, filename))
        files[k] = {"label": label, "canon": canon,
                    "filename": filename,
                    "display": f"当前：{filename}" if exists else f"当前：{filename}（文件不存在）",
                    "custom": bool(get_setting(setting_key))}
    return render_template("admin/topic_settings.html", t=t, files=files)


# ---------------------------------------------------------------- 评价指标管理
def _metric_form_to_row(form):
    """校验后台提交的指标设定，返回入库字典；不合法抛 ValueError（消息直接展示）。"""
    name = (form.get("name") or "").strip()
    task_type = (form.get("task_type") or "").strip()
    method = (form.get("method") or "").strip()
    pred_column = (form.get("pred_column") or "").strip() or "2"
    true_column = (form.get("true_column") or "").strip() or "2"
    description = (form.get("description") or "").strip()
    direction = (form.get("direction") or "").strip()

    if not name:
        raise ValueError("指标名称不能为空。")
    if task_type not in TASK_TYPES:
        raise ValueError("请选择学习类型（回归 / 分类 / 聚类）。")
    if method not in METRIC_LIB:
        raise ValueError("请选择内置指标算法。")
    meta = METRIC_LIB[method]
    if task_type not in meta["tasks"]:
        usable = "、".join(TASK_TYPES[t] for t in meta["tasks"])
        raise ValueError(f"算法「{meta['label']}」适用于 {usable}，与所选学习类型不一致。")
    if direction not in ("min", "max"):
        direction = meta["direction"]

    # 立即对照当前配置的数据文件校验计分列，避免学生提交时才暴露设定错误
    layout = get_sample_layout()
    try:
        resolve_column(pred_column, layout["columns"], "提交结果")
        ans_info = answer_layout()
        resolve_column(true_column, ans_info["columns"], "标准答案")
    except ValueError as e:
        raise ValueError(str(e))

    return {
        "name": name,
        "task_type": task_type,
        "method": method,
        "direction": direction,
        "pred_column": pred_column,
        "true_column": true_column,
        "description": description or meta["desc"],
    }


def rescore_all_submissions(m):
    """按当前启用指标重算全部历史提交（结果文件仍在磁盘上时）。"""
    db = get_db()
    rows = db.execute(
        "SELECT * FROM submissions WHERE result_path IS NOT NULL ORDER BY id").fetchall()
    layout = get_sample_layout()
    updated, failed = 0, []
    for r in rows:
        path = os.path.join(RESULT_DIR, r["result_path"])
        if not os.path.exists(path):
            failed.append((r["id"], "结果文件已缺失"))
            continue
        try:
            with open(path, "rb") as f:
                scored = rescore_from_bytes(f.read(), m, layout)
        except ValueError as e:
            failed.append((r["id"], str(e)))
            continue
        except Exception as e:
            failed.append((r["id"], f"文件读取失败：{e}"))
            continue
        db.execute("UPDATE submissions SET score = ?, metric_name = ? WHERE id = ?",
                   (scored["score"], m["name"], r["id"]))
        updated += 1
    db.commit()
    return updated, failed


@app.route("/admin/metrics", methods=["GET", "POST"])
@admin_required
def admin_metrics():
    db = get_db()

    if request.method == "POST":
        action = request.form.get("action", "create")
        mid = request.form.get("id", type=int)
        target = db.execute("SELECT * FROM metrics WHERE id = ?", (mid,)).fetchone() \
            if mid else None
        if action in ("update", "enable", "delete") and target is None:
            abort(404)
        try:
            if action in ("create", "update"):
                row = _metric_form_to_row(request.form)
                if action == "create":
                    has_enabled = db.execute(
                        "SELECT COUNT(*) AS c FROM metrics WHERE enabled = 1").fetchone()["c"]
                    db.execute(
                        "INSERT INTO metrics (name, task_type, method, direction, pred_column,"
                        " true_column, description, enabled, created_at)"
                        " VALUES (:name, :task_type, :method, :direction, :pred_column,"
                        " :true_column, :description, :enabled, :created_at)",
                        {**row, "enabled": 0 if has_enabled else 1,
                         "created_at": datetime.now().strftime(DT_FMT)},
                    )
                    db.commit()
                    flash(f"已新增评价指标「{row['name']}」"
                          f"{'' if has_enabled else '并自动启用，当前评分即按该指标计算'}。", "success")
                else:
                    db.execute(
                        "UPDATE metrics SET name = :name, task_type = :task_type,"
                        " method = :method, direction = :direction, pred_column = :pred_column,"
                        " true_column = :true_column, description = :description,"
                        " updated_at = :updated_at WHERE id = :id",
                        {**row, "updated_at": datetime.now().strftime(DT_FMT), "id": mid},
                    )
                    db.commit()
                    tip = "（当前启用指标，历史成绩建议重算）" if target["enabled"] else ""
                    flash(f"指标「{row['name']}」设定已保存{tip}。", "success")

            elif action == "enable":
                db.execute("UPDATE metrics SET enabled = (id = ?)", (mid,))
                db.commit()
                flash(f"已启用评价指标「{target['name']}」，此后提交按该指标评分；"
                      "历史提交可在下方按新指标重算。", "success")

            elif action == "delete":
                if target["enabled"]:
                    flash("启用中的指标不能删除，请先启用其他指标。", "error")
                else:
                    db.execute("DELETE FROM metrics WHERE id = ?", (mid,))
                    db.commit()
                    flash(f"已删除指标「{target['name']}」。", "success")

            elif action == "rescore":
                m = get_active_metric()
                updated, failed = rescore_all_submissions(m)
                msg = f"已按「{m['name']}」重算 {updated} 条历史提交。"
                if failed:
                    msg += f" {len(failed)} 条无法重算（保留原成绩）：" + \
                           "；".join(f"#{i} {r}" for i, r in failed[:3])
                flash(msg, "success" if not failed else "warn")
        except ValueError as e:
            flash(str(e), "error")
        return redirect(url_for("admin_metrics"))

    editing = None
    edit_id = request.args.get("edit", type=int)
    if edit_id:
        row = db.execute("SELECT * FROM metrics WHERE id = ?", (edit_id,)).fetchone()
        if row is not None:
            editing = dict(row)

    return render_template("admin/metrics.html",
                           items=list_metrics(),
                           editing=editing,
                           form=editing or {**DEFAULT_METRIC, "description": ""},
                           layout=get_sample_layout(),
                           ans_info=answer_layout(),
                           task_types=TASK_TYPES,
                           methods=[
                               {
                                   "key": key,
                                   "label": meta["label"],
                                   "tasks": list(meta["tasks"]),
                                   "task_labels": "、".join(TASK_TYPES[t] for t in meta["tasks"]),
                                   "direction": meta["direction"],
                                   "direction_text": "越小越好" if meta["direction"] == "min" else "越大越好",
                                   "desc": meta["desc"],
                               }
                               for key, meta in METRIC_LIB.items()
                           ])


# ---------------------------------------------------------------- 成绩导出 Excel
def _ranking_rows():
    """返回带排名的成绩列表（含未提交学生）与名次映射，用于导出。"""
    m = get_active_metric()
    db = get_db()
    students = db.execute(
        f"""
        SELECT u.student_no, u.name, u.gender, u.class_name,
               (SELECT score FROM submissions s WHERE s.user_id = u.id
                 ORDER BY s.score {score_order_sql(m)} LIMIT 1) AS best_score,
               (SELECT COUNT(*)       FROM submissions s WHERE s.user_id = u.id) AS n_submit,
               (SELECT MAX(created_at) FROM submissions s WHERE s.user_id = u.id) AS last_time
        FROM users u WHERE u.role = 'student'
        ORDER BY u.class_name, u.student_no
        """
    ).fetchall()

    # 按启用指标的历史最佳值计算名次（未提交不参与排名）
    key = best_sort_key(m)
    submitted = sorted([r for r in students if r["best_score"] is not None],
                       key=lambda r: key(r["best_score"], r["last_time"]))
    rank_map = {r["student_no"]: i + 1 for i, r in enumerate(submitted)}
    return students, rank_map


@app.route("/admin/export_excel")
@admin_required
def admin_export_excel():
    m = get_active_metric()
    students, rank_map = _ranking_rows()
    data = []
    for r in students:
        data.append({
            "排名": rank_map.get(r["student_no"], ""),
            "学号": r["student_no"],
            "姓名": r["name"],
            "性别": r["gender"],
            "班级": r["class_name"],
            f"最佳{m['name']}": (round(r["best_score"], m["decimals"])
                                if r["best_score"] is not None else ""),
            "评价指标": m["name"] if r["best_score"] is not None else "",
            "提交次数": r["n_submit"],
            "最近提交时间": r["last_time"] or "",
        })
    df = pd.DataFrame(data)
    # 按排名排序导出（未提交放最后）
    df["_sort"] = df["排名"].apply(lambda x: x if x != "" else 10 ** 6)
    df = df.sort_values("_sort").drop(columns="_sort")

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="成绩排名")
        ws = writer.sheets["成绩排名"]
        # 自适应列宽
        for col_cells in ws.columns:
            max_len = max(len(str(c.value)) if c.value is not None else 0 for c in col_cells)
            ws.column_dimensions[col_cells[0].column_letter].width = min(max_len + 4, 30)
    buf.seek(0)

    fname = f"成绩排名_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return send_file(buf, as_attachment=True, download_name=fname,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ---------------------------------------------------------------- 新增用户
@app.route("/admin/users/add", methods=["POST"])
@admin_required
def admin_add_user():
    db = get_db()
    sno = request.form.get("student_no", "").strip()
    name = request.form.get("name", "").strip()
    gender = request.form.get("gender", "").strip()
    class_name = request.form.get("class_name", "").strip()
    init_pwd = request.form.get("init_password", "").strip()

    if not sno or not name:
        flash("学号和姓名为必填项。", "error")
    elif db.execute("SELECT id FROM users WHERE student_no = ?", (sno,)).fetchone():
        flash(f"学号 {sno} 已存在，不能重复添加。", "error")
    else:
        pwd = init_pwd or sno  # 初始密码默认=学号
        now = datetime.now().strftime(DT_FMT)
        db.execute(
            "INSERT INTO users (student_no, name, gender, class_name,"
            " password_hash, role, created_at) VALUES (?, ?, ?, ?, ?, 'student', ?)",
            (sno, name, gender, class_name, generate_password_hash(pwd), now),
        )
        db.commit()
        tip = f"，初始密码：{pwd}" if init_pwd else "，初始密码为学号"
        flash(f"已新增用户 {name}（{sno}）{tip}。", "success")
    return redirect(url_for("admin_users"))


# ---------------------------------------------------------------- 批量导入用户
def _read_user_table(file_storage):
    """
    读取管理员上传的用户表（CSV 或 Excel），返回 DataFrame（全部按字符串处理）。
    CSV 编码自动识别（UTF-8-SIG / UTF-8 / GBK）。
    """
    filename = (file_storage.filename or "").strip()
    ext = os.path.splitext(filename)[1].lower()
    if ext not in (".csv", ".xlsx", ".xls"):
        raise ValueError("仅支持 CSV 或 Excel（.xlsx / .xls）文件。")
    if ext == ".csv":
        raw = file_storage.read()
        text = None
        for enc in ("utf-8-sig", "utf-8", "gbk"):
            try:
                text = raw.decode(enc)
                break
            except UnicodeDecodeError:
                continue
        if text is None:
            raise ValueError("CSV 编码无法识别（需 UTF-8 或 GBK）。")
        return pd.read_csv(io.StringIO(text), dtype=str)
    return pd.read_excel(file_storage.stream, dtype=str)


@app.route("/admin/users/import", methods=["POST"])
@admin_required
def admin_import_users():
    f = request.files.get("file")
    if f is None or not (f.filename or "").strip():
        flash("请先选择要导入的用户表文件。", "error")
        return redirect(url_for("admin_users"))

    try:
        df = _read_user_table(f)
    except ValueError as e:
        flash(str(e), "error")
        return redirect(url_for("admin_users"))
    except Exception:
        flash("文件解析失败，请确认为有效的 CSV 或 Excel 文件。", "error")
        return redirect(url_for("admin_users"))

    df.columns = [str(c).strip() for c in df.columns]
    if "学号" not in df.columns or "姓名" not in df.columns:
        flash("文件缺少必需列：表头须包含“学号、姓名”（可选列：性别、班级）。", "error")
        return redirect(url_for("admin_users"))

    db = get_db()
    now = datetime.now().strftime(DT_FMT)
    added, skipped = 0, 0
    for _, row in df.iterrows():
        sno = str(row.get("学号", "")).strip()
        name = str(row.get("姓名", "")).strip()
        if not sno or not name or sno.lower() == "nan" or name.lower() == "nan":
            skipped += 1
            continue
        if db.execute("SELECT id FROM users WHERE student_no = ?", (sno,)).fetchone():
            skipped += 1
            continue
        gender = str(row.get("性别", "")).strip()
        class_name = str(row.get("班级", "")).strip()
        db.execute(
            "INSERT INTO users (student_no, name, gender, class_name,"
            " password_hash, role, created_at) VALUES (?, ?, ?, ?, ?, 'student', ?)",
            (
                sno,
                name,
                "" if gender.lower() == "nan" else gender,
                "" if class_name.lower() == "nan" else class_name,
                generate_password_hash(sno),  # 初始密码 = 学号
                now,
            ),
        )
        added += 1
    db.commit()

    if added or skipped:
        flash(f"导入完成：新增 {added} 名学生，跳过 {skipped} 条"
              f"（学号已存在或缺学号/姓名）。新用户初始密码为学号。", "success")
    else:
        flash("文件中没有可导入的数据行。", "error")
    return redirect(url_for("admin_users"))


# ---------------------------------------------------------------- 删除单条提交
@app.route("/admin/submission/<int:sid>/delete", methods=["POST"])
@admin_required
def admin_delete_submission(sid):
    db = get_db()
    sub = db.execute("SELECT * FROM submissions WHERE id = ?", (sid,)).fetchone()
    if sub is None:
        abort(404)
    delete_submission_files(sub)
    db.execute("DELETE FROM submissions WHERE id = ?", (sid,))
    db.commit()
    flash(f"已删除提交记录 #{sid} 及其结果/代码文件。", "success")

    # 回到来源页（默认提交查询页）
    ref = request.form.get("ref", "")
    return redirect(ref or url_for("admin_submissions"))


# ---------------------------------------------------------------- 一键全部下载（zip）
def _build_summary_csv(students, rank_map):
    """生成成绩排名汇总 CSV 字节（gbk，Excel 可直接打开）。"""
    m = get_active_metric()
    rows = []
    for r in students:
        rows.append({
            "排名": rank_map.get(r["student_no"], ""),
            "学号": r["student_no"],
            "姓名": r["name"],
            "班级": r["class_name"],
            f"最佳{m['name']}": (f"{r['best_score']:.{m['decimals']}f}"
                                 if r["best_score"] is not None else ""),
            "提交次数": r["n_submit"],
            "最近提交时间": r["last_time"] or "",
        })
    sdf = pd.DataFrame(rows)
    sdf["_sort"] = sdf["排名"].apply(lambda x: x if x != "" else 10 ** 6)
    sdf = sdf.sort_values("_sort").drop(columns="_sort")
    buf = io.BytesIO()
    sdf.to_csv(buf, index=False, encoding="gbk")
    return buf.getvalue()


@app.route("/admin/download_all")
@admin_required
def admin_download_all():
    mode = request.args.get("mode", "best")
    db = get_db()
    buf = io.BytesIO()
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        if mode == "all":
            # 全部历史结果文件：按学生分目录
            subs = db.execute(
                """
                SELECT s.*, u.student_no, u.name FROM submissions s
                JOIN users u ON u.id = s.user_id
                WHERE s.result_path IS NOT NULL ORDER BY s.id
                """
            ).fetchall()
            for s in subs:
                fp = os.path.join(RESULT_DIR, s["result_path"])
                if os.path.exists(fp):
                    arc = (f"{s['student_no']}_{s['name']}/"
                           f"{s['id']}_{s['result_filename']}")
                    zf.write(fp, arc)
            zip_name = f"全部历史提交结果_{stamp}.zip"
        else:
            # 每人历史最佳（按启用指标的最佳值、并列取最早）一条结果文件
            m = get_active_metric()
            all_subs = db.execute(
                f"""
                SELECT s.*, u.student_no, u.name FROM submissions s
                JOIN users u ON u.id = s.user_id
                WHERE s.result_path IS NOT NULL
                ORDER BY s.score {score_order_sql(m)}, s.id ASC
                """
            ).fetchall()
            best_per_user = {}
            for s in all_subs:
                best_per_user.setdefault(s["user_id"], s)

            key = best_sort_key(m)
            ordered = sorted(best_per_user.values(),
                             key=lambda s: key(s["score"], s["id"]))
            for rank, s in enumerate(ordered, start=1):
                fp = os.path.join(RESULT_DIR, s["result_path"])
                if os.path.exists(fp):
                    arc = (f"第{rank:03d}名_{s['student_no']}_{s['name']}"
                           f"_{s['result_filename']}")
                    zf.write(fp, arc)
            zip_name = f"每人最佳提交结果_{stamp}.zip"

        # 附成绩排名汇总
        students, rank_map = _ranking_rows()
        zf.writestr("成绩排名汇总.csv", _build_summary_csv(students, rank_map))

    buf.seek(0)
    return send_file(buf, as_attachment=True, download_name=zip_name,
                     mimetype="application/zip")


# ---------------------------------------------------------------- 模板过滤器 / 全局变量
@app.template_filter("fmt_score")
def fmt_score(value, decimals=None):
    """指标得分格式化：精度由当前启用指标的算法决定。"""
    if value is None:
        return "-"
    if decimals is None:
        decimals = get_active_metric()["decimals"]
    return f"{value:.{decimals}f}"


@app.context_processor
def inject_metric():
    """全站模板可直接使用 metric（当前启用的评价指标设定）。"""
    return {"metric": get_active_metric()}


@app.template_filter("fmt_td")
def fmt_td_filter(value):
    """timedelta -> 'x天x小时x分'。"""
    if value is None:
        return "-"
    return fmt_remain(value)


@app.template_filter("nl2p")
def nl2p_filter(value):
    """多行文本转 HTML 段落：空行分段、单换行转 <br>（内容自动转义）。"""
    text = str(value or "").strip()
    if not text:
        return Markup("")
    parts = []
    for para in re.split(r"\n\s*\n", text):
        para = para.strip()
        if para:
            body = escape(para).replace("\n", Markup("<br>"))
            parts.append(Markup("<p>{}</p>").format(body))
    return Markup("").join(parts)


init_db()

if __name__ == "__main__":
    # 生产环境由 gunicorn 启动（见 Procfile）；本地直接 python app.py 运行
    app.run(
        host=os.environ.get("HOST", "0.0.0.0"),
        port=int(os.environ.get("PORT", 5000)),
        debug=os.environ.get("FLASK_DEBUG", "1") == "1",
    )
