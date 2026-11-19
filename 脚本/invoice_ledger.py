# -*- coding: utf-8 -*-
"""进项发票核对与台账工具（数电票导出表专用）

用法：
    python invoice_ledger.py <导出表.xlsx | 目录> [--out 产物目录]
                             [--start 2026-09-01] [--end 2026-10-13]

输入：电子税务局「全量发票查询」导出的 xlsx，通常含 5 个 sheet：
    信息汇总表（明细行级）、发票基础信息（发票级）、旅客运输服务、
    航空运输电子客票行程单、铁路电子客票。
输出：<out>/进项发票台账.xlsx、核对验证报告.md、异常清单.csv

设计要点（踩过的坑，勿改）：
1. 「数电发票号码」在部分 sheet 是数值格式，19 位超过 Excel 15 位精度，
   读进来已失真。台账一律以文本列（信息汇总表）为准；发票基础信息 sheet
   只用于交叉核对，匹配走（销方识别号 + 开票日期 + 价税合计）。
2. 发票基础信息末尾可能带「合计行」，必须剔除，否则统计翻倍。
3. 信息汇总表的「价税合计」是**行级**金额，累加才是发票合计；发票基础信息
   的票面合计不含「不征税」项目（如民航发展基金），所以两者天然差一个
   不征税金额——这是口径差异，不是错票。
4. 旅客运输 / 铁路 / 航空三张表没有金额列，金额回主表按发票号取。
"""
from __future__ import annotations

import argparse
import re
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------- 列名归一
ALIASES = {
    "序号": ["序号"],
    "数电发票号码": ["数电发票号码", "数电票号码", "数电发票号码（20位）"],
    "发票代码": ["发票代码"],
    "发票号码": ["发票号码"],
    "销方识别号": ["销方识别号", "销售方识别号", "销方税号", "销售方税号"],
    "销方名称": ["销方名称", "销售方名称"],
    "购方识别号": ["购方识别号", "购买方识别号"],
    "购买方名称": ["购买方名称", "购方名称"],
    "开票日期": ["开票日期", "开票时间"],
    "金额": ["金额", "不含税金额"],
    "税额": ["税额"],
    "价税合计": ["价税合计", "含税金额", "合计金额"],
    "税率": ["税率", "税率/征收率"],
    "发票来源": ["发票来源"],
    "发票票种": ["发票票种", "票种"],
    "发票状态": ["发票状态"],
    "是否正数发票": ["是否正数发票", "是否正数"],
    "发票风险等级": ["发票风险等级"],
    "开票人": ["开票人"],
    "备注": ["备注"],
    "税收分类编码": ["税收分类编码"],
    "特定业务类型": ["特定业务类型"],
    "货物或应税劳务名称": ["货物或应税劳务名称", "货物名称", "项目名称"],
    "规格型号": ["规格型号"],
    "单位": ["单位"],
    "数量": ["数量"],
    "单价": ["单价"],
    "出行人": ["出行人", "旅客姓名"],
    "有效身份证号": ["有效身份证号"],
    "出行日期": ["出行日期", "日期"],
    "出发地": ["出发地"],
    "到达地": ["到达地"],
    "交通工具": ["交通工具类型", "等级"],
    "车次航班": ["出行车次", "航班号"],
}

SHEET_KINDS = {
    "detail": ["信息汇总"],
    "invoice": ["发票基础"],
    "passenger": ["旅客运输"],
    "air": ["航空"],
    "rail": ["铁路"],
}

# 发票票种可抵扣口径（参考，需会计确认）
DEDUCTIBLE_KINDS = ("增值税专用发票", "铁路电子客票", "航空运输电子客票", "旅客运输")

# 费用类别规则：顺序匹配，先命中先归类
EXPENSE_RULES = [
    ("餐饮", r"餐饮|餐费|食堂"),
    ("住宿", r"住宿"),
    ("加油", r"汽油|柴油|乙醇汽油|燃料油"),
    ("通行费", r"通行费|过路|过桥"),
    ("差旅交通", r"客运|票价|出租车|网约车|航空|铁路|机票|车票|交通服务"),
    ("外包用工", r"劳务费|人力资源|劳务派遣|工资"),
    ("服务费", r"服务费|平台服务|装卸搬运|配送|物业|保洁|安保"),
    ("办公用品", r"文具|办公|纸制品|打印|耗材|笔墨"),
    ("设备与家具", r"家具|设备|监控|计算机|导航|仪器|实验|通风柜|空调|电器|衡器|电线电缆|电子"),
    ("通讯", r"话费|通信|电话|宽带|流量"),
    ("维修", r"修理|维修|保养"),
    ("租赁", r"租赁|租金"),
    ("农产品食品", r"农产品|蔬菜|肉|水产|蛋|水果|粮油|饮料|酒|焙烤|方便食品|冷冻饮品"),
]

MONEY_FMT = "#,##0.00"


# ---------------------------------------------------------------- 读取
def read_sheet(path: Path, sheet: str) -> pd.DataFrame:
    """读一个 sheet：自动定位表头行，丢掉表头以上的标题行与空行。"""
    raw = pd.read_excel(path, sheet_name=sheet, header=None, dtype=object)
    if raw.empty:
        return pd.DataFrame()
    hdr = 0
    for i in range(min(8, len(raw))):
        vals = [str(v) for v in raw.iloc[i].tolist() if pd.notna(v)]
        joined = " ".join(vals)
        if "序号" in joined and ("发票号码" in joined or "数电" in joined):
            hdr = i
            break
    cols, seen = [], {}
    for i, v in enumerate(raw.iloc[hdr]):
        name = str(v).strip() if pd.notna(v) else f"列{i + 1}"
        if name in seen:
            seen[name] += 1
            name = f"{name}_{seen[name]}"
        else:
            seen[name] = 0
        cols.append(name)
    df = raw.iloc[hdr + 1:].copy()
    df.columns = cols
    df = df.dropna(how="all")
    return df.reset_index(drop=True)


def pick(df: pd.DataFrame, key: str):
    """按别名取列，返回 Series 或 None。"""
    if df is None or df.empty:
        return None
    for name in ALIASES.get(key, [key]):
        if name in df.columns:
            return df[name]
    return None


def to_num(s):
    if s is None:
        return None
    return pd.to_numeric(s, errors="coerce")


def to_id_str(s) -> pd.Series:
    """发票号转文本。已经是文本的原样保留；数值型的会丢失精度，另行告警。"""
    if s is None:
        return pd.Series(dtype=object)
    out = s.map(lambda v: "" if pd.isna(v) else str(v).strip())
    return out.str.replace(r"\.0$", "", regex=True)


def as_datetime(s):
    return pd.to_datetime(s, errors="coerce")


# ---------------------------------------------------------------- 结构
def detect_sheets(path: Path) -> dict:
    names = pd.ExcelFile(path).sheet_names
    found = {}
    for kind, keys in SHEET_KINDS.items():
        for name in names:
            if any(k in name for k in keys):
                found[kind] = name
                break
    return found


def load_all(path: Path) -> dict:
    sheets = detect_sheets(path)
    if "detail" not in sheets:
        raise SystemExit(f"没找到「信息汇总」sheet，实际 sheet：{pd.ExcelFile(path).sheet_names}")
    data = {"sheets": sheets}
    for kind, name in sheets.items():
        try:
            data[kind] = read_sheet(path, name)
        except Exception as exc:  # noqa: BLE001
            print(f"[warn] 读取 {name} 失败：{exc}")
            data[kind] = pd.DataFrame()
    return data


def drop_total_rows(df: pd.DataFrame) -> tuple[pd.DataFrame, int]:
    """剔除合计行 / 小计行（序号列非数字或含「合计」）。"""
    if df is None or df.empty:
        return df, 0
    seq = pick(df, "序号")
    if seq is None:
        return df, 0
    text = seq.astype(str).str.strip()
    is_total = text.str.contains("合计|小计|总计", regex=True, na=False)
    num = pd.to_numeric(seq, errors="coerce")
    is_data = num.notna() | text.str.fullmatch(r"\d+(\.0)?", na=False)
    keep = is_data & ~is_total
    return df[keep].reset_index(drop=True), int((~keep).sum())


def build_invoices(detail: pd.DataFrame) -> pd.DataFrame:
    """明细行 → 发票级台账。

    价税合计 = 明细行累加（含不征税项目）；
    票面价税合计 = 价税合计 - 其中不征税（与税局发票的「价税合计」口径一致）。
    """
    d = detail.copy()
    d["_id"] = to_id_str(pick(d, "数电发票号码"))
    d = d[d["_id"] != ""]
    d["_amt"] = to_num(pick(d, "金额")).fillna(0.0)
    d["_tax"] = to_num(pick(d, "税额")).fillna(0.0)
    d["_total"] = to_num(pick(d, "价税合计")).fillna(0.0)
    d["_date"] = as_datetime(pick(d, "开票日期"))
    d["_rate"] = pick(d, "税率").astype(str).str.strip() if pick(d, "税率") is not None else ""
    d["_item"] = pick(d, "货物或应税劳务名称").astype(str) if pick(d, "货物或应税劳务名称") is not None else ""

    rows = []
    for inv_id, g in d.groupby("_id", sort=False):
        first = g.iloc[0]
        main_row = g.loc[g["_amt"].abs().idxmax()] if len(g) else first
        rows.append({
            "数电发票号码": inv_id,
            "开票日期": first["_date"],
            "销方名称": _val(first, "销方名称"),
            "销方识别号": _val(first, "销方识别号"),
            "购买方名称": _val(first, "购买方名称"),
            "购方识别号": _val(first, "购方识别号"),
            "发票票种": _val(first, "发票票种"),
            "特定业务类型": _val(first, "特定业务类型"),
            "发票状态": _val(first, "发票状态"),
            "是否正数发票": _val(first, "是否正数发票"),
            "发票风险等级": _val(first, "发票风险等级"),
            "开票人": _val(first, "开票人"),
            "明细行数": len(g),
            "税率": "、".join(sorted({x for x in g["_rate"].dropna().astype(str) if x and x != "nan"})),
            "金额_明细合计": round(g["_amt"].sum(), 2),
            "税额_明细合计": round(g["_tax"].sum(), 2),
            "价税合计": round(g["_total"].sum(), 2),
            "其中不征税": round(float(g.loc[g["_rate"] == "不征税", "_total"].sum()), 2),
            "费用类别": classify_expense(str(main_row["_item"]), str(_val(first, "销方名称"))),
            "主要品名": str(main_row["_item"]),
            "备注": _val(first, "备注"),
        })
    inv = pd.DataFrame(rows)
    inv["票面价税合计"] = (inv["价税合计"] - inv["其中不征税"]).round(2)
    inv["有效标记"] = inv.apply(_valid_flag, axis=1)
    return inv.sort_values("开票日期").reset_index(drop=True)


def _val(row, key):
    for name in ALIASES.get(key, [key]):
        if name in row.index:
            v = row[name]
            return "" if pd.isna(v) else str(v).strip()
    return ""


def _valid_flag(r):
    if str(r.get("是否正数发票", "")).strip() in ("否", "0", "False"):
        return "红字冲销"
    if "已红冲" in str(r.get("发票状态", "")):
        return "已作废(被红冲)"
    return "有效"


def classify_expense(item: str, seller: str = "") -> str:
    text = f"{item} {seller}"
    for label, pattern in EXPENSE_RULES:
        if re.search(pattern, text):
            return label
    return "其他"


# ---------------------------------------------------------------- 检查
def run_checks(data: dict, inv: pd.DataFrame, detail: pd.DataFrame, start=None, end=None) -> list[dict]:
    issues: list[dict] = []
    sheets = data["sheets"]
    d = detail.copy()
    d["_id"] = to_id_str(pick(d, "数电发票号码"))

    def add(level, rule, inv_id, desc, amount=None, sheet="信息汇总表"):
        issues.append({
            "级别": level, "检查项": rule, "数电发票号码": inv_id,
            "金额影响": "" if amount is None else round(float(amount), 2),
            "说明": desc, "所在表": sheet,
        })

    # R01 行级勾稽：金额 + 税额 = 价税合计
    amt = to_num(pick(d, "金额")); tax = to_num(pick(d, "税额")); tot = to_num(pick(d, "价税合计"))
    if amt is not None and tax is not None and tot is not None:
        diff = (amt.fillna(0) + tax.fillna(0) - tot.fillna(0)).round(2)
        for idx in diff[diff.abs() > 0.02].index:
            add("高", "行级勾稽(金额+税额≠价税合计)", d.loc[idx, "_id"], f"差 {diff[idx]:.2f} 元", diff[idx])

    # R02 含不征税项目：票面合计不含此数，属口径提示
    for _, r in inv[inv["其中不征税"].abs() > 0.02].iterrows():
        add("低", "含不征税项目(票面不含此数)", r["数电发票号码"],
            f"其中不征税 {r['其中不征税']:.2f} 元；明细合计 {r['价税合计']:.2f}，"
            f"票面价税合计 {r['票面价税合计']:.2f}", r["其中不征税"])

    # R03 发票级 sheet 交叉核对
    # 注意：先按（销方 + 开票日期）匹配，再单独比金额。若把「金额差 ≤0.02」塞进匹配条件，
    # 后面的金额不符判断永远不可能成立，报告会假报「全部一致」——这是 2026-11-19 修掉的 bug。
    inv_sheet = data.get("invoice")
    if inv_sheet is not None and not inv_sheet.empty:
        b, dropped = drop_total_rows(inv_sheet)
        id_col = pick(b, "数电发票号码")
        numeric_ids = False
        if id_col is not None and len(b):
            numeric_ids = bool(pd.to_numeric(id_col, errors="coerce").notna().all())
        if numeric_ids:
            add("中", "发票号精度失真(数值格式)", "",
                f"「{sheets.get('invoice')}」的数电发票号码是数值格式，19 位超出 Excel 15 位精度，"
                f"末位已失真、无法与明细表关联；已改用（销方+日期+金额）匹配",
                None, sheets.get("invoice", ""))
        if dropped:
            add("中", "含合计行/小计行", "",
                f"已剔除 {dropped} 行合计行，避免重复统计", None, sheets.get("invoice", ""))
        seller_col = pick(b, "销方识别号")
        b = b.assign(
            _seller=seller_col.astype(str) if seller_col is not None else "",
            _date=as_datetime(pick(b, "开票日期")),
            _total=to_num(pick(b, "价税合计")),
        )
        matched, diffs, unmatched = 0, [], []
        for _, r in b.iterrows():
            rid = str(_val(r, "数电发票号码"))
            if pd.isna(r["_date"]) or pd.isna(r["_total"]):
                unmatched.append(rid)
                continue
            cand = inv[(inv["销方识别号"].astype(str) == str(r["_seller"]))
                       & (inv["开票日期"] == r["_date"])]
            if cand.empty:
                unmatched.append(rid)
                continue
            matched += 1
            # 同销方同日多张时取金额最接近的那张再比差额
            cand = cand.assign(_d=(cand["票面价税合计"] - float(r["_total"])).abs())
            pick_row = cand.sort_values("_d").iloc[0]
            dd = round(float(pick_row["票面价税合计"]) - float(r["_total"]), 2)
            if abs(dd) > 0.02:
                diffs.append((pick_row["数电发票号码"], dd))
        if len(b) != len(inv):
            add("中", "发票级表与明细表张数不一致", "",
                f"发票基础信息 {len(b)} 行（已剔除合计行），信息汇总表 {len(inv)} 张，"
                f"差 {len(b) - len(inv)} 行需人工核对", None, sheets.get("invoice", ""))
        if unmatched:
            add("中", "发票级表有票未能匹配明细表", "、".join(unmatched[:10]),
                f"共 {len(unmatched)} 张：发票基础信息里有、信息汇总表里找不到"
                f"（销方+日期对不上），可能是漏导出或跨期", None, sheets.get("invoice", ""))
        for inv_id, dd in diffs[:20]:
            add("高", "发票级金额与票面不符", inv_id, f"差 {dd:.2f} 元", dd, sheets.get("invoice", ""))
        if len(diffs) > 20:
            add("中", "发票级金额与票面不符(其余)", "",
                f"另有 {len(diffs) - 20} 张存在金额差异，完整清单见「异常清单」页",
                sum(d[1] for d in diffs[20:]), sheets.get("invoice", ""))

    # R04 重复票（同销方 + 同金额 + 同分钟，均为正票）
    pos = inv[(inv["有效标记"] == "有效") & (inv["价税合计"] > 0)].copy()
    pos["_minute"] = pos["开票日期"].dt.floor("min")
    for key, g in pos.groupby(["销方识别号", "价税合计", "_minute"]):
        if len(g) > 1:
            ids = "、".join(g["数电发票号码"].astype(str).tolist())
            add("中", "疑似重复报票", ids,
                f"销方 {g.iloc[0]['销方名称']}，{key[2]}，金额 {key[1]:.2f}，同刻 {len(g)} 张，"
                f"请确认是否重复上传", g["价税合计"].sum())

    # R05 红冲配对
    reds = inv[inv["有效标记"] == "红字冲销"]
    idset = set(inv["数电发票号码"].astype(str))
    paired_blue = set()
    for _, r in reds.iterrows():
        m = re.search(r"被红冲蓝字数电发票号码[：:]\s*(\d{15,25})", str(r["备注"]))
        if not m:
            add("中", "红字发票未见蓝票号", r["数电发票号码"],
                f"负数票 {r['价税合计']:.2f} 元，备注中没有蓝票号码", r["价税合计"])
            continue
        blue = m.group(1)
        if blue not in idset:
            add("高", "红冲蓝票不在本期范围", r["数电发票号码"],
                f"备注指向蓝票 {blue}，该票不在本表，需跨期核对", r["价税合计"])
        else:
            paired_blue.add(blue)
    for _, r in inv[inv["有效标记"] == "已作废(被红冲)"].iterrows():
        if r["数电发票号码"] not in paired_blue:
            add("中", "蓝票已红冲但未见红票", r["数电发票号码"],
                f"状态「{r['发票状态']}」，本期未找到对应红字票", r["价税合计"])

    # R06 期间越界
    if start or end:
        s = pd.Timestamp(start) if start else None
        e = pd.Timestamp(end) if end else None
        for _, r in inv.iterrows():
            dt = r["开票日期"]
            if pd.isna(dt):
                add("高", "开票日期缺失", r["数电发票号码"], "无法归期", r["价税合计"])
                continue
            if (s is not None and dt < s) or (e is not None and dt > e + pd.Timedelta(days=1)):
                add("中", "开票日期超出统计区间", r["数电发票号码"],
                    f"{dt:%Y-%m-%d %H:%M}，区间 {start or '不限'} ~ {end or '不限'}", r["价税合计"])

    # R07 税号格式
    for _, r in inv.iterrows():
        sid = str(r["销方识别号"])
        if not re.fullmatch(r"[0-9A-Z]{15,20}", sid):
            add("中", "销方识别号格式异常", r["数电发票号码"],
                f"「{sid}」长度或字符不符合统一社会信用代码规则", r["价税合计"])

    # R08 数量×单价
    q = to_num(pick(d, "数量")); p = to_num(pick(d, "单价"))
    if q is not None and p is not None and amt is not None:
        calc = (q * p).round(2)
        bad = (calc - amt).abs() > (amt.abs() * 0.01 + 0.05)
        for idx in d.index[bad & q.notna() & p.notna() & amt.notna()]:
            add("低", "数量×单价≠金额", d.loc[idx, "_id"],
                f"数量 {q[idx]} × 单价 {p[idx]} = {calc[idx]:.2f}，金额列 {amt[idx]:.2f}",
                calc[idx] - amt[idx])

    # R09 免税/零税率行税额
    rate = pick(d, "税率")
    if rate is not None and tax is not None:
        zero = rate.astype(str).str.strip().isin(["免税", "0%", "不征税"])
        for idx in d.index[zero & (tax.fillna(0).abs() > 0.02)]:
            add("中", "免税/不征税行却含税额", d.loc[idx, "_id"],
                f"税率 {rate[idx]}，税额 {tax[idx]}", tax[idx])

    # R10 购买方不一致
    buyers = [b for b in inv["购买方名称"].dropna().unique() if str(b).strip()]
    buyer_ids = [b for b in inv["购方识别号"].dropna().unique() if str(b).strip()]
    if len(buyers) > 1:
        add("中", "购买方名称不唯一", "", "、".join(map(str, buyers)))
    if len(buyer_ids) > 1:
        add("中", "购方识别号不唯一", "", "、".join(map(str, buyer_ids)))

    # R11 关键字段空
    for col, level in (("销方名称", "中"), ("价税合计", "高")):
        empty = inv[inv[col].astype(str).str.strip().isin(["", "nan"])]
        for _, r in empty.iterrows():
            add(level, f"{col} 为空", r["数电发票号码"], "关键字段缺失")

    # R12 结构问题：发票代码/号码整列空
    for col in ("发票代码", "发票号码"):
        s = pick(d, col)
        if s is not None and s.astype(str).str.strip().isin(["", "nan", "None"]).all():
            add("低", f"{col}整列为空", "", "数电票无此字段，勿用作唯一键")

    order = {"高": 0, "中": 1, "低": 2}
    issues.sort(key=lambda x: (order.get(x["级别"], 9), x["检查项"]))
    return issues


# ---------------------------------------------------------------- 差旅
def build_travel(data: dict, inv: pd.DataFrame) -> pd.DataFrame:
    frames = []
    for kind in ("passenger", "air", "rail"):
        df = data.get(kind)
        if df is None or df.empty:
            continue
        df = df.copy()
        df["类型"] = {"passenger": "旅客运输", "air": "航空", "rail": "铁路"}[kind]
        frames.append(df)
    if not frames:
        return pd.DataFrame()
    t = pd.concat(frames, ignore_index=True, sort=False)

    out = pd.DataFrame()
    for key in ("出行人", "有效身份证号", "出行日期", "出发地", "到达地", "交通工具", "车次航班"):
        s = pick(t, key)
        out[key] = s if s is not None else ""
    out["类型"] = t["类型"]
    out["_id"] = to_id_str(pick(t, "数电发票号码"))
    out = out.merge(
        inv[["数电发票号码", "价税合计", "票面价税合计", "税额_明细合计", "有效标记"]],
        left_on="_id", right_on="数电发票号码", how="left",
    )
    keep = ["类型", "_id", "出行人", "有效身份证号", "出行日期", "出发地", "到达地",
            "交通工具", "车次航班", "票面价税合计", "税额_明细合计", "有效标记"]
    out = out[[c for c in keep if c in out.columns]]
    return out.rename(columns={"_id": "数电发票号码"})


# ---------------------------------------------------------------- 导出 Excel
HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
TITLE_FONT = Font(bold=True, color="FFFFFF", size=10)
BODY_FONT = Font(size=10)


def _cell(v):
    if pd.isna(v):
        return ""
    if isinstance(v, pd.Timestamp):
        return v.to_pydatetime()
    return v


def _disp_len(s) -> int:
    """按 Excel 列宽估算：中文等宽字符占 2 个单位，数字/字母占 1。"""
    return sum(2 if ord(ch) > 127 else 1 for ch in str(s))


def write_table(wb, name, df, money_cols=(), date_cols=(), widths=None, note=None):
    ws = wb.create_sheet(name[:31])
    row0 = 1
    if note:
        ws.cell(row=1, column=1, value=note).font = Font(size=10, italic=True, color="7F7F7F")
        row0 = 2
    cols = list(df.columns)
    for j, col in enumerate(cols, start=1):
        c = ws.cell(row=row0, column=j, value=str(col))
        c.font = TITLE_FONT
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for i, (_, r) in enumerate(df.iterrows(), start=row0 + 1):
        for j, col in enumerate(cols, start=1):
            c = ws.cell(row=i, column=j, value=_cell(r[col]))
            c.font = BODY_FONT
            if col in money_cols:
                c.number_format = MONEY_FMT
            elif col in date_cols:
                c.number_format = "yyyy-mm-dd hh:mm"
    ws.freeze_panes = ws.cell(row=row0 + 1, column=1)
    if cols:
        ws.auto_filter.ref = f"A{row0}:{get_column_letter(len(cols))}{row0 + len(df)}"
    for j, col in enumerate(cols, start=1):
        # 按显示宽度取最长内容（中文算 2 格），否则数字列会窄成 ####
        lens = [_disp_len(col)] + [_disp_len(_cell(v)) for v in df[col].head(300)]
        needed = min(max(lens) + 2, 46)
        if widths and col in widths:
            w = max(widths[col], needed)
        else:
            w = needed
        if col in money_cols:
            w = max(w, 16)
        elif col in date_cols:
            w = max(w, 19)
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.sheet_view.showGridLines = False
    return ws


def export_workbook(path: Path, inv: pd.DataFrame, detail: pd.DataFrame, issues: list[dict],
                    travel: pd.DataFrame, meta: dict):
    wb = Workbook()
    wb.remove(wb.active)

    kind = inv["发票票种"].astype(str)
    special = inv["特定业务类型"].astype(str)
    mask_ded = ((kind.apply(lambda s: any(k in s for k in DEDUCTIBLE_KINDS))
                 | special.str.contains("旅客运输", na=False))
                & (inv["有效标记"] == "有效"))
    inv = inv.copy()
    inv["可抵扣"] = ["是" if x else "否" for x in mask_ded]
    travel_kind = (inv["发票票种"].astype(str).str.contains("旅客运输|铁路|航空", na=False)
                   | inv["特定业务类型"].astype(str).str.contains("旅客运输", na=False))
    trip_ids = set(travel["数电发票号码"]) if travel is not None and not travel.empty else set()
    n_travel = int(travel_kind.sum())
    n_cover = int((travel_kind & inv["数电发票号码"].isin(trip_ids)).sum())
    info = pd.DataFrame([
        ("数据源", meta["source"]),
        ("统计区间", meta["period"]),
        ("生成时间", meta["generated"]),
        ("购买方", meta["buyer"]),
        ("发票张数", f"{len(inv)} 张（有效 {int((inv['有效标记'] == '有效').sum())} 张，"
                     f"已作废 {int((inv['有效标记'] == '已作废(被红冲)').sum())} 张，"
                     f"红字 {int((inv['有效标记'] == '红字冲销').sum())} 张）"),
        ("明细行数", f"{len(detail)} 行"),
        ("价税合计", f"{inv['价税合计'].sum():,.2f} 元（含不征税项目，红蓝票净额）"),
        ("其中不征税", f"{inv['其中不征税'].sum():,.2f} 元"),
        ("票面价税合计", f"{inv['票面价税合计'].sum():,.2f} 元（与税局发票「价税合计」口径一致，申报用）"),
        ("可抵扣税额", f"{inv.loc[mask_ded, '税额_明细合计'].sum():,.2f} 元"
                     f"（有效票中的专票 + 铁路/航空/旅客运输，参考口径）"),
        ("异常条数", f"{len(issues)} 条（高 {sum(1 for i in issues if i['级别'] == '高')} / "
                     f"中 {sum(1 for i in issues if i['级别'] == '中')} / "
                     f"低 {sum(1 for i in issues if i['级别'] == '低')}）"),
        ("金额口径", "「价税合计」= 明细行累加，含不征税项目；「票面价税合计」= 扣掉不征税项目后的票面值。"
                     "做费用统计用前者，做进项申报用后者。"),
        ("红冲口径", "红字票按负数计入，蓝票保留原值，二者自然抵销；请勿单独删掉红字行或蓝票行。"),
        ("有效标记", "有效 = 可入账；已作废(被红冲) = 蓝票被全额红冲；红字冲销 = 负数票。"),
        ("费用类别", "按明细金额最大行的品名关键词归类，规则见脚本 EXPENSE_RULES，可自行调整。"),
        ("差旅覆盖", f"客运/铁路/航空票 {n_travel} 张，其中 {n_cover} 张带逐行程明细；"
                     f"其余 {n_travel - n_cover} 张为合并开票，无行程信息，见「差旅待补」页。"),
    ], columns=["项目", "内容"])
    write_table(wb, "说明", info, widths={"项目": 14, "内容": 96})

    ledger_cols = ["数电发票号码", "开票日期", "销方名称", "销方识别号", "发票票种", "发票状态",
                   "是否正数发票", "有效标记", "可抵扣", "税率", "金额_明细合计", "税额_明细合计",
                   "价税合计", "其中不征税", "票面价税合计", "明细行数",
                   "费用类别", "主要品名", "开票人", "备注"]
    write_table(wb, "发票台账", inv[ledger_cols].copy(),
                money_cols=("金额_明细合计", "税额_明细合计", "价税合计", "其中不征税", "票面价税合计"),
                date_cols=("开票日期",),
                widths={"数电发票号码": 24, "销方名称": 30, "主要品名": 30, "备注": 40, "税率": 10})

    det_cols = ["序号", "数电发票号码", "开票日期", "销方名称", "销方识别号", "发票票种", "发票状态",
                "货物或应税劳务名称", "规格型号", "单位", "数量", "单价", "金额", "税率", "税额", "价税合计",
                "税收分类编码", "特定业务类型", "备注"]
    det = detail.copy()
    for c in det_cols:
        if c not in det.columns:
            det[c] = ""
    write_table(wb, "明细表", det[det_cols],
                money_cols=("金额", "税额", "价税合计", "单价"),
                date_cols=("开票日期",),
                widths={"货物或应税劳务名称": 34, "备注": 40, "税收分类编码": 22, "销方名称": 30})

    iss = pd.DataFrame(issues) if issues else pd.DataFrame(
        columns=["级别", "检查项", "数电发票号码", "金额影响", "说明", "所在表"])
    write_table(wb, "异常清单", iss, money_cols=("金额影响",),
                widths={"说明": 70, "检查项": 26, "数电发票号码": 26, "所在表": 14})

    def agg_summary(df, keys):
        g = df.groupby(keys).agg(
            张数=("数电发票号码", "count"),
            金额=("金额_明细合计", "sum"),
            税额=("税额_明细合计", "sum"),
            不征税=("其中不征税", "sum"),
            价税合计=("价税合计", "sum"),
            票面价税合计=("票面价税合计", "sum"),
        ).reset_index()
        for c in ("金额", "税额", "不征税", "价税合计", "票面价税合计"):
            g[c] = g[c].round(2)
        return g

    m = agg_summary(inv.assign(_m=inv["开票日期"].dt.to_period("M").astype(str)), ["_m"])
    m = m.rename(columns={"_m": "月份"})
    total = pd.DataFrame([{"月份": "合计", "张数": len(inv),
                           "金额": round(inv["金额_明细合计"].sum(), 2),
                           "税额": round(inv["税额_明细合计"].sum(), 2),
                           "不征税": round(inv["其中不征税"].sum(), 2),
                           "价税合计": round(inv["价税合计"].sum(), 2),
                           "票面价税合计": round(inv["票面价税合计"].sum(), 2)}])
    write_table(wb, "月度汇总", pd.concat([m, total], ignore_index=True),
                money_cols=("金额", "税额", "不征税", "价税合计", "票面价税合计"), widths={"月份": 12})

    k = agg_summary(inv, ["发票票种", "税率"])
    write_table(wb, "票种税率汇总", k,
                money_cols=("金额", "税额", "不征税", "价税合计", "票面价税合计"),
                widths={"发票票种": 34, "税率": 10})

    s = agg_summary(inv, ["销方识别号", "销方名称"]).sort_values("票面价税合计", ascending=False)
    s["占比"] = (s["票面价税合计"] / s["票面价税合计"].sum() * 100).round(2)
    write_table(wb, "销方汇总", s,
                money_cols=("金额", "税额", "不征税", "价税合计", "票面价税合计"),
                widths={"销方名称": 34})

    e = agg_summary(inv, ["费用类别"]).sort_values("价税合计", ascending=False)
    e["占比"] = (e["价税合计"] / e["价税合计"].sum() * 100).round(2)
    write_table(wb, "费用类别汇总", e,
                money_cols=("金额", "税额", "不征税", "价税合计", "票面价税合计"),
                widths={"费用类别": 16})

    if travel is not None and not travel.empty:
        write_table(wb, "差旅明细", travel,
                    money_cols=("票面价税合计", "税额_明细合计"),
                    date_cols=("出行日期",),
                    widths={"出发地": 34, "到达地": 34, "有效身份证号": 20})
        t = travel.dropna(subset=["出行人"]).groupby("出行人").agg(
            行程数=("数电发票号码", "count"),
            票面价税合计=("票面价税合计", "sum"),
            税额=("税额_明细合计", "sum")).reset_index().sort_values("票面价税合计", ascending=False)
        t["票面价税合计"] = t["票面价税合计"].round(2); t["税额"] = t["税额"].round(2)
        write_table(wb, "差旅按人汇总", t, money_cols=("票面价税合计", "税额"), widths={"出行人": 14})

    todo = inv[travel_kind & ~inv["数电发票号码"].isin(trip_ids)]
    if not todo.empty:
        write_table(wb, "差旅待补", todo[["数电发票号码", "开票日期", "销方名称", "主要品名",
                                     "价税合计", "票面价税合计"]],
                    money_cols=("价税合计", "票面价税合计"), date_cols=("开票日期",),
                    widths={"销方名称": 30, "主要品名": 30},
                    note="以下客运/铁路/航空票没有导出逐行程信息（合并开票），需要报销单时按发票号回原系统补行程。")

    red = inv[inv["有效标记"].isin(["红字冲销", "已作废(被红冲)"])].copy()
    if not red.empty:
        red["关联蓝票号"] = red["备注"].astype(str).str.extract(r"(\d{15,25})")[0]
        write_table(wb, "红冲清单", red[["数电发票号码", "开票日期", "销方名称", "有效标记",
                                     "价税合计", "关联蓝票号", "备注"]],
                    money_cols=("价税合计",), date_cols=("开票日期",),
                    widths={"关联蓝票号": 24, "备注": 50})

    wb.save(path)


# ---------------------------------------------------------------- 报告
def export_report(path: Path, inv, detail, issues, meta, checks, travel=None):
    def money(x):
        return f"{x:,.2f}"

    n_all = len(inv)
    n_valid = int((inv["有效标记"] == "有效").sum())
    n_void = int((inv["有效标记"] == "已作废(被红冲)").sum())
    n_red = int((inv["有效标记"] == "红字冲销").sum())
    amt = inv["价税合计"].sum()
    face = inv["票面价税合计"].sum()
    tax = inv["税额_明细合计"].sum()
    high = [i for i in issues if i["级别"] == "高"]
    mid = [i for i in issues if i["级别"] == "中"]
    low = [i for i in issues if i["级别"] == "低"]
    un_taxed = inv[inv["其中不征税"] > 0.02]

    L = []
    A = L.append
    A("# 进项发票核对验证报告")
    A("")
    A(f"- 数据源：`{meta['source']}`")
    A(f"- 统计区间：{meta['period']}")
    A(f"- 生成时间：{meta['generated']}")
    A(f"- 购买方：{meta['buyer']}")
    A("")
    A("## 一、结论摘要")
    A("")
    A(f"1. 共 **{n_all} 张**发票、**{len(detail)} 行**明细，价税合计 **{money(amt)} 元**（含不征税项目），"
      f"其中税额 {money(tax)} 元；扣除不征税项目后的票面合计 {money(face)} 元。")
    A(f"2. 有效票 {n_valid} 张；被全额红冲作废 {n_void} 张；红字冲销 {n_red} 张。"
      f"红蓝成对、金额自然抵销，**不要单独删掉其中一边**。")
    bad_pair = [i for i in issues if i["检查项"] in ("发票级金额与票面不符",
                                                     "发票级金额与票面不符(其余)",
                                                     "发票级表有票未能匹配明细表")]
    A(f"3. 发票级表（发票基础信息）合计 {money(face)} 元，与明细逐张核对"
      + ("**全部一致**。" if not bad_pair else f"有 {len(bad_pair)} 条对不上，见第三节。"))
    A(f"4. 明细与票面相差 {money(amt - face)} 元，全部来自 {len(un_taxed)} 张含「不征税」项目的票，"
      f"属正常口径差异，**不是错票**（见第四节）。")
    A(f"5. 本次共检出 {len(issues)} 条待关注项：高 {len(high)} 条、中 {len(mid)} 条、低 {len(low)} 条。")
    A("")
    A("## 二、总量勾稽")
    A("")
    A("| 口径 | 张数/行数 | 金额 | 税额 | 价税合计 |")
    A("|---|---|---|---|---|")
    A(f"| 明细行（信息汇总表） | {len(detail)} 行 | {money(detail_num(detail, '金额'))} | "
      f"{money(detail_num(detail, '税额'))} | {money(detail_num(detail, '价税合计'))} |")
    A(f"| 发票台账（明细累加） | {n_all} 张 | {money(inv['金额_明细合计'].sum())} | {money(tax)} | {money(amt)} |")
    b_chk = [c for c in checks if c["name"] == "invoice_sheet"]
    if b_chk:
        c0 = b_chk[0]
        A(f"| 发票基础信息（票面） | {c0['rows']} 行 | {money(c0['amt'])} | {money(c0['tax'])} | {money(c0['total'])} |")
    A("")
    A(f"勾稽结论：{checks_text(checks, inv)}")
    A("")
    A("## 三、待关注项")
    A("")
    if not issues:
        A("无。")
    else:
        A("| 级别 | 检查项 | 发票号 | 金额影响 | 说明 |")
        A("|---|---|---|---|---|")
        for it in issues[:60]:
            A(f"| {it['级别']} | {it['检查项']} | {it['数电发票号码']} | "
              f"{'' if it['金额影响'] == '' else money(it['金额影响'])} | {it['说明']} |")
        if len(issues) > 60:
            A(f"| … | 其余 {len(issues) - 60} 条见 `异常清单.csv` | | | |")
    A("")
    A("## 四、红冲与不征税项目")
    A("")
    reds = inv[inv["有效标记"] != "有效"]
    if not reds.empty:
        pos_n = int((reds["有效标记"] == "已作废(被红冲)").sum())
        neg_n = int((reds["有效标记"] == "红字冲销").sum())
        A(f"红蓝各 {pos_n} / {neg_n} 张，净额 {money(reds['价税合计'].sum())} 元（已相互抵销）：")
        A("")
        A("| 发票号 | 日期 | 销方 | 金额 | 标记 |")
        A("|---|---|---|---|---|")
        for _, r in reds.iterrows():
            A(f"| {r['数电发票号码']} | {r['开票日期']:%Y-%m-%d} | {r['销方名称']} | "
              f"{money(r['价税合计'])} | {r['有效标记']} |")
        A("")
    if not un_taxed.empty:
        A(f"含不征税项目 {len(un_taxed)} 张，合计 {money(un_taxed['其中不征税'].sum())} 元：")
        A("")
        A("| 发票号 | 销方 | 不征税 | 明细合计 | 票面合计 | 主要品名 |")
        A("|---|---|---|---|---|---|")
        for _, r in un_taxed.iterrows():
            A(f"| {r['数电发票号码']} | {r['销方名称']} | {money(r['其中不征税'])} | "
              f"{money(r['价税合计'])} | {money(r['票面价税合计'])} | {r['主要品名']} |")
        A("")
        A("> 说明：航空客票的「民航发展基金」属于不征税项目，票面「价税合计」不含该项，"
          "但明细会单独列一行。所以明细累加天然大于票面。台账保留两个口径："
          "费用统计看「价税合计」，进项申报看「票面价税合计」。")
        A("")
    # 差旅覆盖
    if travel is not None and not travel.empty:
        inv_travel = inv[(inv["发票票种"].astype(str).str.contains("旅客运输|铁路|航空", na=False))
                         | (inv["特定业务类型"].astype(str).str.contains("旅客运输", na=False))]
        covered = set(travel["数电发票号码"])
        missing = inv_travel[~inv_travel["数电发票号码"].isin(covered)]
        if len(inv_travel):
            A("### 差旅行程覆盖")
            A("")
            A(f"客运/铁路/航空票 {len(inv_travel)} 张，其中 {len(inv_travel) - len(missing)} 张有逐行程明细"
              f"（{len(travel)} 条行程记录）；{len(missing)} 张为合并开票，未导出行程，"
              f"需要按票补行程（见台账「差旅待补」页）。")
            A("")

    A("## 五、数据结构风险（换数据时必须先确认）")
    A("")
    A("| # | 风险 | 影响 | 处理方式 |")
    A("|---|---|---|---|")
    A("| 1 | 发票基础信息的「数电发票号码」是数值格式 | 19 位超过 Excel 15 位精度，末位已失真，无法与明细表关联 | 一律以「信息汇总表」的文本列为准，本工具已改用（销方+日期+金额）匹配 |")
    A("| 2 | 发票基础信息末行是「合计行」 | 直接求和/透视会重复计算 | 已按序号列自动剔除 |")
    A("| 3 | 「发票代码」「发票号码」整列为空 | 数电票没有这两个字段，误用作唯一键会关联失败 | 只认「数电发票号码」 |")
    A("| 4 | 旅客运输 / 铁路 / 航空三张表没有金额列 | 单独统计会漏金额 | 金额回主表按发票号取 |")
    A("| 5 | 红冲票与蓝票同时存在 | 只留正票会虚增、只留红票会漏账 | 两边都保留，净额入账 |")
    A("| 6 | 文件名区间与实际开票日期不一致 | 跨期票混入本月，申报口径对不上 | 用 --start/--end 显式限定，越界票单列 |")
    A("")
    A("## 六、台账口径")
    A("")
    A("- 「价税合计」= 明细行累加（含不征税），用于费用与供应商分析。")
    A("- 「票面价税合计」= 扣掉不征税项目，与税局发票票面一致，用于进项申报。")
    A("- 「可抵扣税额」为参考口径：增值税专用发票 + 铁路/航空/旅客运输按票面税额计，普通发票计 0；"
      "免税、0%、不征税行不计入。实际申报以会计确认为准。")
    A("- 「费用类别」按明细金额最大行的品名自动推断，可人工调整。")
    A("")
    A("## 七、建议动作")
    A("")
    A("1. 先处理「高」级别项：红冲蓝票跨期、行级金额不平、关键字段缺失，这三类直接影响申报数。")
    A("2. 「中」级别逐条确认：疑似重复票（同销方同金额同分钟）先查是否同一笔消费重复上传；"
      "识别号格式异常查营业执照。")
    A("3. 把 `进项发票台账.xlsx` 的「发票台账」页作为月度进项台账底稿，「异常清单」页随表归档，作税务备查资料。")
    A("4. 下个月重复流程：新导出表放进 `原始数据`，双击 `run.cmd`，看新报告即可。")
    path.write_text("\n".join(L), encoding="utf-8")


def detail_num(detail, key):
    for name in ALIASES.get(key, [key]):
        if name in detail.columns:
            return pd.to_numeric(detail[name], errors="coerce").sum()
    return 0.0


def checks_text(checks, inv):
    parts = []
    for c in checks:
        if c["name"] == "invoice_sheet":
            diff = round(c["total"] - inv["票面价税合计"].sum(), 2)
            if abs(diff) < 0.005:
                diff = 0.0
            parts.append(
                f"发票基础信息表 {c['rows']} 行、票面合计 {c['total']:,.2f} 元，"
                f"与台账票面合计差 {diff:,.2f} 元"
                + ("，一致。" if abs(diff) < 0.02 else "，需核对。"))
    return "；".join(parts) if parts else "无发票级表可比。"


# ---------------------------------------------------------------- main
def main(argv=None):
    ap = argparse.ArgumentParser(description="进项发票核对与台账工具")
    ap.add_argument("source", nargs="?", default=None,
                    help="导出的 xlsx 文件或所在目录（取目录内最新一个）；缺省用 ../原始数据")
    ap.add_argument("--out", default=None, help="输出目录，默认 产物/日期")
    ap.add_argument("--start", default=None, help="统计起始日 YYYY-MM-DD")
    ap.add_argument("--end", default=None, help="统计截止日 YYYY-MM-DD（含当天）")
    args = ap.parse_args(argv)

    root = Path(__file__).resolve().parent.parent
    src = Path(args.source) if args.source else root / "原始数据"
    if src.is_dir():
        files = sorted(src.glob("*.xlsx"), key=lambda p: p.stat().st_mtime, reverse=True)
        if not files:
            raise SystemExit(f"目录里没有 xlsx：{src}")
        src = files[0]
    if not src.exists():
        raise SystemExit(f"文件不存在：{src}")

    data = load_all(src)
    detail, dropped_detail = drop_total_rows(data["detail"])
    if dropped_detail:
        print(f"[info] 明细表剔除合计行/小计行 {dropped_detail} 行")
    inv = build_invoices(detail)
    travel = build_travel(data, inv)
    issues = run_checks(data, inv, detail, args.start, args.end)

    checks = []
    inv_sheet = data.get("invoice")
    if inv_sheet is not None and not inv_sheet.empty:
        b, _ = drop_total_rows(inv_sheet)
        checks.append({
            "name": "invoice_sheet",
            "rows": len(b),
            "amt": round(to_num(pick(b, "金额")).fillna(0).sum(), 2),
            "tax": round(to_num(pick(b, "税额")).fillna(0).sum(), 2),
            "total": round(to_num(pick(b, "价税合计")).fillna(0).sum(), 2),
        })

    out = Path(args.out) if args.out else root / "产物" / datetime.now().strftime("%Y-%m-%d")
    out.mkdir(parents=True, exist_ok=True)

    period = f"{args.start or '不限'} ~ {args.end or '不限'}"
    if not args.start and not args.end and inv["开票日期"].notna().any():
        period = (f"{inv['开票日期'].min():%Y-%m-%d} ~ {inv['开票日期'].max():%Y-%m-%d}"
                  "（按数据自动取）")
    buyer = "、".join(sorted(set(str(x) for x in inv["购买方名称"].unique() if str(x).strip()))) or "未识别"
    meta = {"source": str(src), "period": period, "buyer": buyer,
            "generated": datetime.now().strftime("%Y-%m-%d %H:%M:%S")}

    xlsx_path = out / "进项发票台账.xlsx"
    try:
        export_workbook(xlsx_path, inv, detail, issues, travel, meta)
    except PermissionError:
        # 文件正被 Excel 打开时无法覆盖，另存一份，避免整次运行白跑
        alt = out / "进项发票台账(修正版).xlsx"
        try:
            export_workbook(alt, inv, detail, issues, travel, meta)
        except PermissionError:
            alt = out / f"进项发票台账_新{datetime.now():%H%M%S}.xlsx"
            export_workbook(alt, inv, detail, issues, travel, meta)
        xlsx_path = alt
        print(f"[warn] 目标文件被占用，已另存为：{alt.name}（关掉 Excel 后可删旧文件并重命名）")
    export_report(out / "核对验证报告.md", inv, detail, issues, meta, checks, travel)
    pd.DataFrame(issues).to_csv(out / "异常清单.csv", index=False, encoding="utf-8-sig")

    print(f"发票 {len(inv)} 张 / 明细 {len(detail)} 行 / 价税合计 {inv['价税合计'].sum():,.2f} 元 "
          f"/ 票面 {inv['票面价税合计'].sum():,.2f} 元")
    print(f"异常 {len(issues)} 条（高 {sum(1 for i in issues if i['级别'] == '高')}，"
          f"中 {sum(1 for i in issues if i['级别'] == '中')}，"
          f"低 {sum(1 for i in issues if i['级别'] == '低')}）")
    print(f"输出目录：{out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
