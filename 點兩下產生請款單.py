#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
中油客戶請款單產生器

每次執行會先從 GitHub 拉最新版程式（加 --no-pull 可跳過）。

用法：
    python3 點兩下產生請款單.py                 # 處理「上個月、本月、下個月」三個月份資料夾
    python3 點兩下產生請款單.py 2026-01         # 只處理指定月份（可一次給多個）
    python3 點兩下產生請款單.py 2026-01 --no-overwrite   # 已存在時另存為 01月(自動產出).xlsx（預設直接覆蓋）

資料夾結構：
    <本程式所在資料夾>/
        客戶折讓表.xlsx          每位客戶的折讓規則
        2026-01/
            輸入/A002-1.xlsx     中油系統匯出的「簽帳客戶對帳單」
            輸入/A041-1.xlsx ...
            01月.xlsx            <- 產出

產出內容：
    本月總計、簽收單，以及每位客戶的請款對帳單：
        -1 依交易時間排序；-2 依車號分組；-3 依油品分組（車號或油品只有一種時不產出該表）
    輸入資料夾中沒有的客戶不會產出任何表。
"""
import argparse
import datetime as dt
import hashlib
import os
import re
import subprocess
import sys
from collections import defaultdict
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter

BASE_DIR = Path(__file__).resolve().parent
DISCOUNT_FILE = BASE_DIR / "客戶折讓表.xlsx"
INPUT_DIR_NAME = "輸入"

COMPANY_TITLE = "通進加油站股份有限公司－－－主成加油站"
COMPANY_ADDR = "台東縣太麻里鄉美和村美和86之31號　　TEL：089-515915"
FOOTER_LINES = [
    ("請在5天內將現金送達加油站或將金額匯款", False),
    ("", False),
    ("請  滙 入：台灣中小企業銀行台東分行（050）", True),
    ("    帳 號：77012101281", True),
    ("    戶 名：通進加油站股份有限公司", True),
    ("", False),
    ("請  滙 入：華南銀行台東分行（008）", True),
    ("    帳 號：830100168131", True),
    ("    戶 名：通進加油站股份有限公司", True),
]

FONT_NAME = "標楷體"
F10 = Font(name=FONT_NAME, size=10)
F10B = Font(name=FONT_NAME, size=10, bold=True)
F16B = Font(name=FONT_NAME, size=16, bold=True)
CENTER = Alignment(horizontal="center", vertical="center")
LEFT = Alignment(horizontal="left", vertical="center")
RIGHT = Alignment(horizontal="right", vertical="center")
THIN = Side(style="thin")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
BOTTOM_MEDIUM = Border(bottom=Side(style="medium"))

FMT_PRICE = r"0.0_);[Red]\(0.0\)"
FMT_QTY = r"0.00_);[Red]\(0.00\)"
FMT_INT = "#,##0_ "
FMT_DATETIME = r"[$-409]yyyy/m/d\ h:mm\ AM/PM;@"
FMT_SUM2 = r"#,##0.00_);[Red]\(#,##0.00\)"
FMT_SUM0 = r"#,##0_);[Red]\(#,##0\)"

# 本月總計的油品分類（欄位順序）
PRODUCT_CATEGORIES = ["92", "95", "柴油", "尿素", "副產品"]


# ---------------------------------------------------------------- 工具
def roc(d):
    """西元日期 -> 民國 yyymmdd 字串，例如 2026-01-01 -> 1150101"""
    return f"{d.year - 1911}{d.month:02d}{d.day:02d}"


def product_category(product):
    p = product or ""
    if "九二" in p or "92" in p:
        return "92"
    if "九五" in p or "95" in p or "九八" in p or "98" in p:
        return "95"
    if "柴油" in p:
        return "柴油"
    if "尿素" in p:
        return "尿素"
    return "副產品"


def months_around_today(today=None):
    today = today or dt.date.today()
    first = today.replace(day=1)
    prev_last = first - dt.timedelta(days=1)
    next_first = (first + dt.timedelta(days=32)).replace(day=1)
    return [prev_last.strftime("%Y-%m"), first.strftime("%Y-%m"), next_first.strftime("%Y-%m")]


def safe_sheet_name(name):
    name = re.sub(r"[\[\]\*\?/\\:]", "", name)
    return name[:31]


# ---------------------------------------------------------------- 折讓表
def load_discount_table(path):
    """回傳 {客戶編號: dict(name, mode, value, note)}
    mode = 'mult'  : 折讓金額 = 數量 * value
    mode = 'minus' : 折讓後單價 = 單價 - value（老爺模式）
    """
    table = {}
    if not path.exists():
        print(f"[警告] 找不到折讓表 {path}，所有客戶折讓以 0 計算")
        return table
    ws = openpyxl.load_workbook(path, data_only=True).active
    for row in ws.iter_rows(min_row=2, values_only=True):
        code, name, rule, note = (list(row) + [None] * 4)[:4]
        if not code:
            continue
        rule = str(rule or "")
        m = re.search(r"數量\s*\*\s*([\d.]+)", rule)
        if m:
            table[str(code).strip()] = dict(name=str(name).strip(), mode="mult", value=float(m.group(1)), note=note or "")
            continue
        m = re.search(r"單價\s*-\s*([\d.]+)", rule)
        if m:
            table[str(code).strip()] = dict(name=str(name).strip(), mode="minus", value=float(m.group(1)), note=note or "")
            continue
        print(f"[警告] 折讓表 {code} 的規則「{rule}」無法解析，折讓以 0 計算")
        table[str(code).strip()] = dict(name=str(name).strip(), mode="mult", value=0.0, note=note or "")
    return table


# ---------------------------------------------------------------- 讀取輸入
def parse_input_file(path):
    """讀取中油對帳單，回傳 dict(code, start, end, rows)"""
    m = re.match(r"^([A-Za-z]\d{3})", path.stem)
    if not m:
        print(f"[略過] {path.name}：檔名開頭不是客戶編號（如 A002）")
        return None
    code = m.group(1).upper()
    ws = openpyxl.load_workbook(path, data_only=True).active

    # 資料區間
    start = end = None
    for row in ws.iter_rows(min_row=1, max_row=10, values_only=True):
        for v in row:
            if isinstance(v, str) and "資料區間" in v:
                found = re.findall(r"(\d{4})[-/](\d{1,2})[-/](\d{1,2})", v)
                if len(found) >= 2:
                    start = dt.date(*map(int, found[0]))
                    end = dt.date(*map(int, found[1]))
    # 表頭列
    header_row = None
    for r in range(1, min(ws.max_row, 30) + 1):
        if ws.cell(r, 2).value == "島號":
            header_row = r
            break
    if header_row is None:
        print(f"[略過] {path.name}：找不到「島號」表頭")
        return None
    headers = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(header_row, c).value
        if isinstance(v, str) and v.strip():
            headers[v.strip()] = c

    def col(*names):
        for n in names:
            if n in headers:
                return headers[n]
        raise KeyError(names)

    c_name, c_car, c_time = col("客戶名稱"), col("車號"), col("交易時間")
    c_prod, c_price, c_qty, c_amt = col("油品"), col("單價"), col("數量"), col("交易金額")

    rows = []
    for r in range(header_row + 1, ws.max_row + 1):
        if ws.cell(r, 2).value == "合計":
            break
        t = ws.cell(r, c_time).value
        if not isinstance(t, dt.datetime):
            continue
        rows.append(dict(
            name=str(ws.cell(r, c_name).value or "").strip(),
            car=str(ws.cell(r, c_car).value or "").strip(),
            time=t,
            product=str(ws.cell(r, c_prod).value or "").strip(),
            price=float(ws.cell(r, c_price).value or 0),
            qty=float(ws.cell(r, c_qty).value or 0),
            amount=float(ws.cell(r, c_amt).value or 0),
        ))
    rows = remove_reversals(rows, path.name)
    if not rows:
        print(f"[略過] {path.name}：沒有交易資料")
        return None
    if start is None:
        start, end = min(x["time"] for x in rows).date(), max(x["time"] for x in rows).date()
    return dict(code=code, start=start, end=end, rows=rows, file=path.name)


def remove_reversals(rows, fname):
    """負數交易（沖銷）與被沖銷的原交易成對移除：同車號、同油品、同單價、數量與金額剛好相反"""
    keep = list(rows)
    for neg in [x for x in rows if x["qty"] < 0 or x["amount"] < 0]:
        if neg not in keep:
            continue
        for pos in keep:
            if (pos is not neg and pos["car"] == neg["car"] and pos["product"] == neg["product"]
                    and abs(pos["qty"] + neg["qty"]) < 1e-6 and abs(pos["amount"] + neg["amount"]) < 1e-6):
                keep.remove(pos)
                keep.remove(neg)
                print(f"[沖銷] {fname}：{neg['car']} {neg['product']} {pos['qty']:g} 公升 / {pos['amount']:g} 元 的交易已被沖銷，兩筆均不列入")
                break
        else:
            print(f"[注意] {fname}：{neg['car']} {neg['time']} 有負數交易 {neg['qty']:g} 公升但找不到對應原交易，仍照列")
    return keep


def build_billing_units(parsed_files, discount_table):
    """同一客戶、同一資料區間的檔案合併為一個請款單位；不同區間（如 1-15 / 16-31）各自一個單位"""
    groups = defaultdict(list)
    for p in parsed_files:
        groups[(p["code"], p["start"], p["end"])].append(p)
    units = []
    for (code, start, end), items in sorted(groups.items()):
        rows = [r for it in items for r in it["rows"]]
        disc = discount_table.get(code)
        if disc is None:
            print(f"[警告] 客戶 {code} 不在折讓表中，折讓以 0 計算")
            disc = dict(name=rows[0]["name"], mode="mult", value=0.0, note="")
        units.append(dict(code=code, start=start, end=end, rows=rows, disc=disc, label=""))
    # 同一客戶有多個區間時，加上 (1-15) 之類的標籤
    by_code = defaultdict(list)
    for u in units:
        by_code[u["code"]].append(u)
    for code, us in by_code.items():
        if len(us) > 1:
            for u in us:
                u["label"] = f"({u['start'].day}-{u['end'].day})"
    return units


# ---------------------------------------------------------------- 客戶請款單
def write_statement(wb, unit, suffix, groups):
    """groups: list of list(rows)。只有一組 = 依時間排序；多組 = 分組且每組有小計列"""
    disc = unit["disc"]
    minus_mode = disc["mode"] == "minus"
    title = f"{unit['code']}{disc['name']}{unit['label']}{suffix}"
    ws = wb.create_sheet(safe_sheet_name(title))

    widths = {"A": 9.78, "B": 9.78, "C": 20.78, "D": 8.78, "E": 6.78, "F": 9.78, "G": 9.78, "H": 9.78, "I": 9.78}
    for k, v in widths.items():
        ws.column_dimensions[k].width = v
    last_col = 9 if minus_mode else 8
    last_letter = get_column_letter(last_col)

    def put(r, c, v, font=F10, align=None, fmt=None, border=None):
        cell = ws.cell(r, c, v)
        cell.font = font
        if align is not None:
            cell.alignment = align
        if fmt:
            cell.number_format = fmt
        if border:
            cell.border = border
        return cell

    # 表頭
    ws.merge_cells("A1:H1"); put(1, 1, COMPANY_TITLE, Font(name=FONT_NAME, size=18, bold=True), CENTER)
    ws.merge_cells("A2:H2"); put(2, 1, COMPANY_ADDR, Font(name=FONT_NAME, size=12), CENTER)
    ws.merge_cells("A3:H3"); put(3, 1, "請 款 對 帳 單 ", Font(name=FONT_NAME, size=16), CENTER)
    ws.merge_cells("A5:H5"); put(5, 1, f"客戶名稱：{disc['name']}", F10, LEFT)
    ws.merge_cells("A6:H6"); put(6, 1, f"資料區間：{roc(unit['start'])} 至 {roc(unit['end'])}", F10, LEFT)
    ws.row_dimensions[1].height = 25.05
    ws.row_dimensions[2].height = 19.95
    ws.row_dimensions[3].height = 19.95
    for r in range(4, 8):
        ws.row_dimensions[r].height = 16.5

    heads = ["客戶名稱", "車    號", "交   易   時   間", "油    品", "單 價", "數  量", "交易金額"]
    heads += ["折讓後單價", "折讓後金額"] if minus_mode else ["折讓金額"]
    for i, h in enumerate(heads, 1):
        put(7, i, h, F10, CENTER)

    # 明細
    r = 8
    subtotal_rows = []
    for gi, rows in enumerate(groups):
        first = r
        for x in rows:
            put(r, 1, x["name"], F10, CENTER)
            put(r, 2, x["car"], F10, CENTER)
            put(r, 3, x["time"], F10, LEFT, FMT_DATETIME)
            put(r, 4, x["product"], F10, LEFT)
            put(r, 5, x["price"], F10, RIGHT, FMT_PRICE)
            put(r, 6, x["qty"], F10, RIGHT, FMT_QTY)
            put(r, 7, x["amount"], F10, RIGHT, FMT_PRICE)
            if minus_mode:
                put(r, 8, f"=E{r}-{disc['value']:g}", F10, RIGHT, FMT_PRICE)
                put(r, 9, f"=H{r}*F{r}", F10, RIGHT, FMT_PRICE)
            else:
                put(r, 8, f"=F{r}*{disc['value']:g}", F10, RIGHT, FMT_PRICE)
            ws.row_dimensions[r].height = 15
            r += 1
        if len(groups) > 1:  # 分組小計列
            last = r - 1
            rng = f"{first}:{last}" if last > first else f"{first}"
            put(r, 6, f"=SUM(F{rng.replace(':', ':F')})", F10, RIGHT, FMT_QTY)
            put(r, 7, f"=SUM(G{rng.replace(':', ':G')})", F10, RIGHT, FMT_PRICE)
            if minus_mode:
                put(r, 9, f"=SUM(I{rng.replace(':', ':I')})", F10, RIGHT, FMT_PRICE)
            else:
                put(r, 8, f"=SUM(H{rng.replace(':', ':H')})", F10, RIGHT, FMT_PRICE)
            subtotal_rows.append(r)
            r += 1
    data_last = r - 1

    # 小計 / 應收款
    r += 1
    sub_r = r
    amount_col = "I" if minus_mode else "H"
    if len(groups) > 1:
        def chain(col):
            return "=SUM(" + "+".join(f"{col}{sr}" for sr in subtotal_rows) + ")"
        f_qty, f_amt, f_disc = chain("F"), chain("G"), chain(amount_col)
    else:
        f_qty, f_amt, f_disc = (f"=SUM({c}8:{c}{data_last})" for c in ("F", "G", amount_col))
    for c in range(4, last_col + 1):
        ws.cell(sub_r, c).border = BOTTOM_MEDIUM
    put(sub_r, 4, "小計", F10B, LEFT, border=BOTTOM_MEDIUM)
    put(sub_r, 6, f_qty, F10B, RIGHT, FMT_QTY, BOTTOM_MEDIUM)
    put(sub_r, 7, f_amt, F10B, RIGHT, FMT_PRICE, BOTTOM_MEDIUM)
    put(sub_r, 8 if not minus_mode else 9, f_disc, F10B, RIGHT, FMT_PRICE, BOTTOM_MEDIUM)
    ws.row_dimensions[sub_r].height = 16.5

    rec_r = sub_r + 2
    for c in range(4, last_col + 1):
        ws.cell(rec_r, c).border = BOTTOM_MEDIUM
    put(rec_r, 4, "應收款", F10B, LEFT, border=BOTTOM_MEDIUM)
    if minus_mode:
        put(rec_r, 8, f"=I{sub_r}", F10B, RIGHT, FMT_INT, BOTTOM_MEDIUM)
        receivable_ref = f"H{rec_r}"
    else:
        put(rec_r, 7, f"=G{sub_r}-H{sub_r}", F10B, RIGHT, FMT_INT, BOTTOM_MEDIUM)
        receivable_ref = f"G{rec_r}"
    ws.row_dimensions[rec_r].height = 16.5
    r = rec_r + 2

    # 要有結餘款的客戶：結餘款小表（匯入金額由人工填寫）
    if "結餘" in str(disc.get("note", "")):
        prev_end = unit["start"] - dt.timedelta(days=1)
        put(r, 1, f"{prev_end.month:02d}/{prev_end.day:02d}", F10, LEFT); put(r, 2, "尚餘", F10, LEFT); put(r, 3, None, F10, RIGHT, FMT_INT)
        put(r + 1, 1, "", F10, LEFT); put(r + 1, 2, "匯入", F10, LEFT); put(r + 1, 3, None, F10, RIGHT, FMT_INT)
        put(r + 2, 1, f"{unit['start'].month:02d}/{unit['start'].day}~{unit['end'].month:02d}/{unit['end'].day}", F10, LEFT)
        put(r + 2, 2, "油款", F10, LEFT); put(r + 2, 3, f"={receivable_ref}", F10, RIGHT, FMT_INT)
        put(r + 3, 1, f"{unit['end'].month:02d}/{unit['end'].day:02d}", F10, LEFT); put(r + 3, 2, "餘額", F10, LEFT)
        put(r + 3, 3, f"=SUM(C{r}+C{r + 1}-C{r + 2})", F10, RIGHT, FMT_INT)
        r += 5

    # 頁尾
    for text, bold in FOOTER_LINES:
        if text:
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
            put(r, 1, text, F10B if bold else F10, Alignment(horizontal="left"))
        r += 1
    put(r, 6, "謝謝您", F10, Alignment(vertical="center"))

    # 列印設定
    ws.print_area = f"A1:{last_letter}{r}"
    ws.page_setup.orientation = "portrait"
    ws.page_setup.paperSize = 9  # A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins.left, ws.page_margins.right = 0.63, 0.24
    ws.page_margins.top, ws.page_margins.bottom = 0.39, 0.24
    return ws


def write_customer_sheets(wb, unit):
    rows = sorted(unit["rows"], key=lambda x: x["time"])
    cars = sorted({x["car"] for x in rows})
    cat_order = {c: i for i, c in enumerate(PRODUCT_CATEGORIES)}
    products = sorted({x["product"] for x in rows}, key=lambda p: (cat_order[product_category(p)], p))

    variants = [[rows]]  # -1 依時間
    if len(cars) > 1:
        variants.append([[x for x in rows if x["car"] == c] for c in cars])  # 依車號
    if len(products) > 1:
        variants.append([[x for x in rows if x["product"] == p] for p in products])  # 依油品

    if len(variants) == 1:
        write_statement(wb, unit, "", variants[0])
    else:
        for i, groups in enumerate(variants, 1):
            write_statement(wb, unit, f"-{i}", groups)


# ---------------------------------------------------------------- 本月總計 / 簽收單
def unit_totals(unit):
    """回傳 {分類: dict(qty, amount, discount)}"""
    disc = unit["disc"]
    t = defaultdict(lambda: dict(qty=0.0, amount=0.0, discount=0.0))
    for x in unit["rows"]:
        c = product_category(x["product"])
        t[c]["qty"] += x["qty"]
        t[c]["amount"] += x["amount"]
        if disc["mode"] == "minus":
            t[c]["discount"] += x["amount"] - (x["price"] - disc["value"]) * x["qty"]
        else:
            t[c]["discount"] += x["qty"] * disc["value"]
    return t


def write_summary(wb, units, year, month):
    ws = wb.create_sheet("本月總計")
    widths = {"A": 12.1, "B": 20.6, "C": 10.8, "D": 14, "E": 14, "F": 16, "G": 14, "H": 14, "I": 16, "J": 16, "K": 16, "L": 30}
    for k, v in widths.items():
        ws.column_dimensions[k].width = v

    def put(r, c, v, fmt=None, align=None):
        cell = ws.cell(r, c, v)
        cell.font = F16B
        cell.border = BOX
        if fmt:
            cell.number_format = fmt
        if align:
            cell.alignment = align
        return cell

    heads = ["編號", "單    位", "", "92", "95", "柴油", "尿素", "副產品", "合     計", "淨收入", "實際收入", "備      註"]
    for i, h in enumerate(heads, 1):
        put(1, i, h, align=Alignment(horizontal="center"))

    r = 2
    for u in units:
        totals = unit_totals(u)
        name = f"{u['disc']['name']}{u['label']}"
        for k, (label, key, fmt) in enumerate([("公升", "qty", FMT_SUM2), ("金額", "amount", FMT_SUM0), ("折讓", "discount", FMT_SUM2)]):
            rr = r + k
            put(rr, 1, u["code"] if k == 0 else None, align=LEFT)
            put(rr, 2, name if k == 0 else None, align=LEFT)
            put(rr, 3, label, align=LEFT)
            for ci, cat in enumerate(PRODUCT_CATEGORIES):
                v = totals[cat][key] if cat in totals else None
                if v is not None:
                    v = round(v, 2 if key != "amount" else 0)
                put(rr, 4 + ci, v, fmt, RIGHT)
            put(rr, 9, f"=SUM(D{rr}:H{rr})", fmt, RIGHT)
            put(rr, 10, f"=SUM(I{r + 1}-I{r + 2})" if k == 1 else None, FMT_SUM0, RIGHT)
            put(rr, 11, None, FMT_SUM0, RIGHT)
            put(rr, 12, None)
        r += 3
    last_data = r - 1

    r += 1
    for k, label in enumerate(["公升", "金額", "折讓"]):
        rr = r + k
        put(rr, 1, "合計", align=LEFT)
        put(rr, 2, "合  計" if k == 0 else None, align=LEFT)
        put(rr, 3, label, align=LEFT)
        fmt = FMT_SUM0 if label == "金額" else FMT_SUM2
        for c in range(4, 10):
            L = get_column_letter(c)
            put(rr, c, f'=SUMIF($C$2:$C${last_data},"{label}",{L}2:{L}{last_data})', fmt, RIGHT)
        put(rr, 10, f"=SUM(I{r + 1}-I{r + 2})" if k == 1 else None, FMT_SUM0, RIGHT)
        put(rr, 11, f"=SUM(K2:K{last_data})" if k == 1 else None, FMT_SUM0, RIGHT)
        put(rr, 12, None)
    r += 3
    ws.cell(r + 1, 1, f"{year - 1911}年{month:02d}月").font = F16B

    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = 9
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    return ws


def write_sign_sheet(wb, units, year, month):
    ws = wb.create_sheet("簽收單")
    for k, v in {"A": 10.7, "B": 9.2, "C": 24, "D": 17.8, "E": 16.7, "F": 16.9}.items():
        ws.column_dimensions[k].width = v
    ws.merge_cells("A2:F2")
    c = ws.cell(2, 1, f"{year - 1911}年{month:02d}月客戶帳單簽收")
    c.font = F16B
    c.alignment = LEFT
    heads = ["編號", "公司", "單位", "客戶簽名", "簽收日期", "加油員簽名"]
    for i, h in enumerate(heads, 1):
        cell = ws.cell(5, i, h)
        cell.font = F16B
        cell.border = BOX
        cell.alignment = Alignment(horizontal="left" if i <= 3 else "center")
    r = 6
    for u in units:
        vals = [u["code"], "", f"{u['disc']['name']}{u['label']}", None, None, None]
        for i, v in enumerate(vals, 1):
            cell = ws.cell(r, i, v)
            cell.font = F16B if i <= 3 else Font(name=FONT_NAME, size=16)
            cell.border = BOX
            cell.alignment = LEFT if i <= 3 else CENTER
        ws.row_dimensions[r].height = 30
        r += 1
    ws.page_setup.orientation = "portrait"
    ws.page_setup.paperSize = 9
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    return ws


# ---------------------------------------------------------------- 自動更新
def _file_hash(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def pull_latest_and_reexec():
    """從 GitHub 拉最新版；若本程式被更新，就用新版重新執行自己"""
    if not (BASE_DIR / ".git").exists():
        return
    before = _file_hash(__file__)
    try:
        res = subprocess.run(
            ["git", "-C", str(BASE_DIR), "pull", "--ff-only", "--quiet"],
            capture_output=True, text=True, timeout=60,
        )
        if res.returncode != 0:
            print(f"[更新] git pull 失敗，使用目前版本繼續執行：{res.stderr.strip()}")
            return
    except FileNotFoundError:
        print("[更新] 找不到 git，使用目前版本繼續執行")
        return
    except subprocess.TimeoutExpired:
        print("[更新] git pull 逾時（可能沒有網路），使用目前版本繼續執行")
        return
    if _file_hash(__file__) != before:
        print("[更新] 程式已更新為最新版，重新執行…")
        os.execv(sys.executable, [sys.executable, __file__, *sys.argv[1:], "--no-pull"])
    else:
        print("[更新] 已是最新版")


# ---------------------------------------------------------------- 主流程
def process_month(month_str, discount_table, overwrite=True):
    year, month = map(int, month_str.split("-"))
    month_dir = BASE_DIR / month_str
    input_dir = month_dir / INPUT_DIR_NAME
    if not input_dir.is_dir():
        print(f"[{month_str}] 找不到 {input_dir}，略過")
        return None
    files = sorted(p for p in input_dir.iterdir() if p.suffix.lower() in (".xlsx", ".xlsm") and not p.name.startswith("~$"))
    parsed = [p for p in (parse_input_file(f) for f in files) if p]
    if not parsed:
        print(f"[{month_str}] 輸入資料夾沒有可用的檔案，略過")
        return None
    units = build_billing_units(parsed, discount_table)

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    write_summary(wb, units, year, month)
    write_sign_sheet(wb, units, year, month)
    for u in units:
        write_customer_sheets(wb, u)

    out = month_dir / f"{month:02d}月.xlsx"
    if out.exists() and not overwrite:
        out = month_dir / f"{month:02d}月(自動產出).xlsx"
        print(f"[{month_str}] {month:02d}月.xlsx 已存在，改存為 {out.name}")
    wb.save(out)
    print(f"[{month_str}] 完成：{out}  客戶/請款單位 {len(units)} 個，工作表 {len(wb.sheetnames)} 張")
    for u in units:
        print(f"    {u['code']} {u['disc']['name']}{u['label']}  {roc(u['start'])}~{roc(u['end'])}  {len(u['rows'])} 筆")
    return out


def main(argv=None):
    ap = argparse.ArgumentParser(description="產生中油客戶請款單")
    ap.add_argument("months", nargs="*", help="要處理的月份 YYYY-MM；不給則處理上月/本月/下月")
    ap.add_argument("--overwrite", dest="overwrite", action="store_true", default=True,
                    help="輸出檔已存在時直接覆蓋（預設）")
    ap.add_argument("--no-overwrite", dest="overwrite", action="store_false",
                    help="輸出檔已存在時改存為「MM月(自動產出).xlsx」")
    ap.add_argument("--no-pull", action="store_true", help="執行前不從 GitHub 拉最新版")
    args = ap.parse_args(argv)

    if not args.no_pull:
        pull_latest_and_reexec()

    months = args.months or months_around_today()
    for m in months:
        if not re.fullmatch(r"\d{4}-\d{2}", m):
            sys.exit(f"月份格式錯誤：{m}（應為 YYYY-MM）")
    discount_table = load_discount_table(DISCOUNT_FILE)
    for m in months:
        process_month(m, discount_table, overwrite=args.overwrite)


if __name__ == "__main__":
    main()
