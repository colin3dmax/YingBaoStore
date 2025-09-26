#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
JDFixPackageRelationByName.py  (keep header layout)

写出“修复后.xlsx”时，保留与原始关系表相同的表头行位置：
- 若原始表的表头在第 2 行（header_row_rel=1），修复后文件也在第 2 行写表头；
- 表头之前的提示/说明行原样写回顶部。
"""

import argparse
import os
import re
from typing import Dict, List, Tuple, Set
import pandas as pd

# ==== 可调参数 ====
NAME_CANDIDATES = [
    "商品名称", "品名", "名称", "商品全名", "商品名",
    "单品名称", "单品名称（可不填）",
    "箱品名称", "箱品名称（可不填）",
    "中品名称", "中品名称（可不填）",
]
BARCODE_CANDIDATES = [
    "条形码", "商品条码", "条码", "条码/简码", "主条码",
    "单品条形码*", "中品条形码", "箱品条形码*",
]
RELATION_FIELDS = [
    ("箱品名称（可不填）", "箱品条形码*"),
    ("中品名称（可不填）", "中品条形码"),
    ("单品名称（可不填）", "单品条形码*"),
]
SEPS = [",", "，", ";", "；", " ", "\n", "\t", "/", "／", "|", "丨"]

def first_token(s: str) -> str:
    if s is None: return ""
    s = str(s).strip()
    if not s: return ""
    for sep in SEPS:
        if sep in s:
            parts = [p.strip() for p in s.split(sep) if p.strip()]
            return parts[0] if parts else ""
    return s

def sanitize_barcode(raw) -> str:
    if pd.isna(raw): return ""
    s = first_token(str(raw))
    if s.endswith(".0"): s = s[:-2]
    if not s.isdigit():
        m = re.search(r"\d{6,}", s)
        s = m.group(0) if m else "".join(re.findall(r"\d", s))
    return re.sub(r"\D", "", s)

def normalize_name(x) -> str:
    if pd.isna(x): return ""
    return re.sub(r"\s+", " ", str(x).strip())

def pick_preferred_barcode(codes: List[str]) -> str:
    cleaned = [sanitize_barcode(c) for c in codes if sanitize_barcode(c)]
    if not cleaned: return ""
    for c in cleaned:
        if len(c) == 13 and not c.startswith("205"):
            return c
    return cleaned[0]

def load_product_name_to_barcode(products_path: str) -> Tuple[Dict[str, str], Set[str]]:
    xl = pd.ExcelFile(products_path, engine="openpyxl")
    name_to_codes: Dict[str, List[str]] = {}
    all_codes: Set[str] = set()
    for sheet in xl.sheet_names:
        try:
            df = pd.read_excel(products_path, sheet_name=sheet, dtype="string", engine="openpyxl")
        except Exception:
            df = pd.read_excel(products_path, sheet_name=sheet, engine="openpyxl")
        name_cols = [c for c in NAME_CANDIDATES if c in df.columns]
        code_cols = [c for c in BARCODE_CANDIDATES if c in df.columns]
        if not name_cols or not code_cols: continue
        name_col = name_cols[0]
        for _, row in df.iterrows():
            name = normalize_name(row.get(name_col))
            if not name: continue
            for ccol in code_cols:
                bc = sanitize_barcode(row.get(ccol))
                if bc:
                    all_codes.add(bc)
                    name_to_codes.setdefault(name, []).append(bc)
    name_to_barcode = {n: pick_preferred_barcode(cs) for n, cs in name_to_codes.items()}
    return name_to_barcode, all_codes

def read_relation(path: str, sheet: str = "Sheet1", header_row: int = 1) -> Tuple[pd.DataFrame, pd.DataFrame]:
    # 读标准数据（带表头）
    df_rel = pd.read_excel(path, sheet_name=sheet, header=header_row, dtype="string", engine="openpyxl")
    # 读原始行，保留表头之前内容
    raw = pd.read_excel(path, sheet_name=sheet, header=None, dtype="string", engine="openpyxl")
    if header_row > 0 and not raw.empty:
        ncols = len(df_rel.columns)
        df_prefix = raw.iloc[:header_row, :ncols].copy()
        df_prefix.columns = list(df_rel.columns)[:df_prefix.shape[1]]
    else:
        df_prefix = pd.DataFrame(columns=df_rel.columns)
    return df_rel, df_prefix

def fix_relation_barcodes(df_rel: pd.DataFrame, name_to_barcode: Dict[str, str], all_product_codes: Set[str]):
    df_new = df_rel.copy()
    logs: List[dict] = []
    for idx, row in df_rel.iterrows():
        for name_col, code_col in RELATION_FIELDS:
            if name_col not in df_rel.columns or code_col not in df_rel.columns: continue
            name = normalize_name(row.get(name_col))
            code = sanitize_barcode(row.get(code_col))
            needs_fix = (not code) or (code not in all_product_codes) or code.startswith("205")
            if needs_fix and name:
                new_code = name_to_barcode.get(name, "")
                if new_code and new_code != code:
                    df_new.at[idx, code_col] = new_code
                    logs.append({
                        "row_index": int(idx),
                        "字段": f"{name_col} -> {code_col}",
                        "商品名称": name,
                        "原条码": code,
                        "新条码": new_code,
                        "原因": "按名称在商品库匹配后回填"
                    })
    return df_new, logs

def write_fixed_with_preserved_header(out_path: str, df_fixed: pd.DataFrame, df_prefix: pd.DataFrame, sheet_name: str):
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        start_row = 0
        if not df_prefix.empty:
            df_prefix.to_excel(writer, index=False, header=False, sheet_name=sheet_name, startrow=start_row)
            start_row += len(df_prefix)
        df_fixed.to_excel(writer, index=False, header=True, sheet_name=sheet_name, startrow=start_row)

def main():
    parser = argparse.ArgumentParser(description="按商品名称回填箱单关系表中的缺失/异常条码（保持表头行位置一致）")
    parser.add_argument("--relation", required=True, help="箱单关系表 Excel 路径（含箱/中/单名称与条码列）")
    parser.add_argument("--products", required=True, help="商品资料 Excel 路径（名称->条码来源）")
    parser.add_argument("--sheet-rel", default="Sheet1", help="箱单关系表的工作表名（默认 Sheet1）")
    parser.add_argument("--header-row-rel", type=int, default=1, help="箱单关系表表头所在 0-based 行号（默认 1，即第2行）")
    parser.add_argument("--out", required=False, help="修复后关系表输出路径（默认同名 *_fixed.xlsx）")
    parser.add_argument("--log", required=False, help="修复日志输出路径（默认同名 *_fixlog.xlsx）")
    args = parser.parse_args()

    name_to_barcode, all_codes = load_product_name_to_barcode(args.products)
    df_rel, df_prefix = read_relation(args.relation, sheet=args.sheet_rel, header_row=args.header_row_rel)
    df_fixed, logs = fix_relation_barcodes(df_rel, name_to_barcode, all_codes)

    base, _ = os.path.splitext(args.relation)
    out_path = args.out or f"{base}_fixed.xlsx"
    log_path = args.log or f"{base}_fixlog.xlsx"

    write_fixed_with_preserved_header(out_path, df_fixed, df_prefix, args.sheet_rel)

    # 日志
    if logs:
        df_log = pd.DataFrame(logs, columns=["row_index", "字段", "商品名称", "原条码", "新条码", "原因"])
    else:
        df_log = pd.DataFrame(columns=["row_index", "字段", "商品名称", "原条码", "新条码", "原因"])
    with pd.ExcelWriter(log_path, engine="openpyxl") as writer:
        df_log.to_excel(writer, index=False, sheet_name="修复日志")
        snap_map = pd.DataFrame(sorted([(k, v) for k, v in load_product_name_to_barcode(args.products)[0].items()]),
                                columns=["商品名称", "条码"])
        snap_map.to_excel(writer, index=False, sheet_name="名称→条码映射快照")

    print(f"✅ 修复完成（保持表头行位置）：{out_path}")
    print(f"📝 修复日志：{log_path}")

if __name__ == "__main__":
    main()