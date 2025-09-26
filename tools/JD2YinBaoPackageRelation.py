#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
JD2YinBaoPackageRelation.py

功能：
  - 从原始 Excel（默认 Sheet1，标题行在第 2 行）读取箱单数据
  - 生成标准化四列：大件商品条码 / 小件商品条码 / 换算关系 / 示例
  - 两级关系：有中品时自动生成 “箱→中(用箱品内装数)” 和 “中→单(用中品内装数)”
  - 多条码仅取第一个（支持 , ， ; ； 空格 换行 / ／ | 丨 分隔）
  - 条码始终按字符串处理，防止 .0 / 科学计数法
  - 可选：接收一个“商品资料”Excel，用来校验大/小件条码是否已存在于商品库，并输出缺失清单

依赖：
  pip install pandas openpyxl
"""

import argparse
import os
import re
from typing import List, Set, Tuple
import pandas as pd

# 允许的分隔符（多条码情况下取第一个）
SEPS = [",", "，", ";", "；", " ", "\n", "\t", "/", "／", "|", "丨"]


def first_token(s: str) -> str:
    """按常见分隔符切分，取第一个非空 token"""
    if s is None:
        return ""
    s = str(s).strip()
    if not s:
        return ""
    for sep in SEPS:
        if sep in s:
            parts = [p.strip() for p in s.split(sep) if p.strip()]
            return parts[0] if parts else ""
    return s


def sanitize_barcode(raw) -> str:
    """
    清洗为条码字符串：
      - 先取第一个 token
      - 去尾部 .0
      - 若含非数字，尝试提取第一个 6 位以上数字串；失败则仅保留所有数字
    """
    if pd.isna(raw):
        return ""
    s = first_token(str(raw))

    if s.endswith(".0"):
        s = s[:-2]

    if not s.isdigit():
        m = re.search(r"\d{6,}", s)
        if m:
            s = m.group(0)
        else:
            s = "".join(re.findall(r"\d", s))

    s = re.sub(r"\D", "", s)
    return s


def to_int(n):
    """数量转 int（允许 12.0），失败返回 None"""
    try:
        if pd.isna(n):
            return None
        return int(round(float(n)))
    except Exception:
        return None


def load_relations(
    src_file: str,
    sheet_name: str = "Sheet1",
    header_row: int = 1,
) -> pd.DataFrame:
    """
    读取箱单原始表并生成关系记录 DataFrame
    列要求：箱品条形码*、单品条形码*、箱品内装数*，可选：中品条形码、中品内装数
    """
    dtype_map = {
        "箱品条形码*": "string",
        "中品条形码": "string",
        "单品条形码*": "string",
    }

    df = pd.read_excel(
        src_file,
        sheet_name=sheet_name,
        header=header_row,
        dtype=dtype_map,
        engine="openpyxl",
    )

    for col in ["箱品条形码*", "单品条形码*", "箱品内装数*"]:
        if col not in df.columns:
            raise ValueError(f"缺少必需列：{col}（请确认表头行 header={header_row+1} 与列名一致）")

    records = []

    for _, row in df.iterrows():
        box_code = sanitize_barcode(row.get("箱品条形码*"))
        single_code = sanitize_barcode(row.get("单品条形码*"))
        box_qty = to_int(row.get("箱品内装数*"))

        mid_code = sanitize_barcode(row.get("中品条形码")) if "中品条形码" in df.columns else ""
        mid_qty = to_int(row.get("中品内装数")) if "中品内装数" in df.columns else None

        # 有中品：先 箱→中 用“箱品内装数*”；再 中→单 用“中品内装数”
        if mid_code and (box_qty is not None and box_qty > 0):
            records.append({
                "大件商品条码": box_code,
                "小件商品条码": mid_code,
                "换算关系": box_qty,
                "示例": f"{box_code} = {mid_code} * {box_qty}",
            })
            if single_code and (mid_qty is not None and mid_qty > 0):
                records.append({
                    "大件商品条码": mid_code,
                    "小件商品条码": single_code,
                    "换算关系": mid_qty,
                    "示例": f"{mid_code} = {single_code} * {mid_qty}",
                })

        # 无中品：直接 箱→单 用“箱品内装数*”
        elif box_code and single_code and (box_qty is not None and box_qty > 0):
            records.append({
                "大件商品条码": box_code,
                "小件商品条码": single_code,
                "换算关系": box_qty,
                "示例": f"{box_code} = {single_code} * {box_qty}",
            })

    return pd.DataFrame(records, columns=["大件商品条码", "小件商品条码", "换算关系", "示例"])


# -------------------- 校验（可选） --------------------

PRODUCT_BARCODE_CANDIDATES = [
    "条形码", "商品条码", "条码", "条码/简码",
    "箱品条形码*", "中品条形码", "单品条形码*",
]

def collect_product_barcodes(product_file: str) -> Set[str]:
    """从商品资料文件中收集所有条码（多列&多条码取第一个），返回集合"""
    xl = pd.ExcelFile(product_file, engine="openpyxl")
    frames: List[pd.DataFrame] = []
    for sheet in xl.sheet_names:
        try:
            df = pd.read_excel(product_file, sheet_name=sheet, dtype="string", engine="openpyxl")
            frames.append(df)
        except Exception:
            pass
    if not frames:
        return set()

    merged = pd.concat(frames, ignore_index=True)
    codes = set()
    for col in PRODUCT_BARCODE_CANDIDATES:
        if col in merged.columns:
            for v in merged[col].astype("string").fillna(""):
                code = sanitize_barcode(v)
                if code:
                    codes.add(code)
    return codes


def validate_relations(df_rel: pd.DataFrame, product_codes: Set[str]) -> Tuple[Set[str], Set[str]]:
    """返回（缺失的大件条码集合，缺失的小件条码集合）——指不在商品库中的条码"""
    missing_large = set()
    missing_small = set()
    for _, r in df_rel.iterrows():
        big = str(r["大件商品条码"]).strip()
        sml = str(r["小件商品条码"]).strip()
        if big and big not in product_codes:
            missing_large.add(big)
        if sml and sml not in product_codes:
            missing_small.add(sml)
    return missing_large, missing_small


def save_missing_report(missing_large: Set[str], missing_small: Set[str], path: str):
    """将缺失条码清单存为 Excel（两个sheet）"""
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        pd.DataFrame(sorted(missing_large), columns=["缺失大件商品条码"]).to_excel(writer, index=False, sheet_name="缺失大件条码")
        pd.DataFrame(sorted(missing_small), columns=["缺失小件商品条码"]).to_excel(writer, index=False, sheet_name="缺失小件条码")


# -------------------- CLI --------------------

def main():
    parser = argparse.ArgumentParser(description="银豹箱单关系转换（支持校验产品库）")
    parser.add_argument("--src", required=True, help="输入箱单 Excel 文件路径")
    parser.add_argument("--out", help="输出 Excel（默认同名 _converted.xlsx）")
    parser.add_argument("--sheet", default="Sheet1", help="箱单工作表名，默认 Sheet1")
    parser.add_argument("--header-row", type=int, default=1, help="箱单表头所在 0-based 行号，默认 1（即第2行为表头）")

    parser.add_argument("--products", help="商品资料 Excel（用于校验条码是否存在）")
    parser.add_argument("--missing-report", help="缺失条码清单输出路径（提供了 --products 才生效）")

    args = parser.parse_args()

    # 生成关系表
    df_rel = load_relations(args.src, sheet_name=args.sheet, header_row=args.header_row)

    # 输出关系表
    out_file = args.out or f"{os.path.splitext(args.src)[0]}_converted.xlsx"
    df_rel.to_excel(out_file, index=False)
    print(f"✅ 已生成关系表：{out_file}（{len(df_rel)} 条）")

    # 可选：校验商品库
    if args.products:
        product_codes = collect_product_barcodes(args.products)
        print(f"📦 商品库条码数量：{len(product_codes)}")

        missing_large, missing_small = validate_relations(df_rel, product_codes)

        print(f"🔎 缺失大件条码：{len(missing_large)}")
        print(f"🔎 缺失小件条码：{len(missing_small)}")

        if args.missing_report:
            save_missing_report(missing_large, missing_small, args.missing_report)
            print(f"📝 已输出缺失条码清单：{args.missing_report}")

        if missing_large or missing_small:
            print("\n⚠️ 提示：导入箱单关系前，需先在商品资料中建立所有相关条码（大件与小件都要有）。")

if __name__ == "__main__":
    main()
