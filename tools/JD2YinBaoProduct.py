# Create a reusable Python script that reproduces the full cleaning & export pipeline
# script_path = "data/商品导入清洗脚本.py"
# script_code = r'''#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
商品导入清洗脚本（JD 万商导出 -> 导入模板）
-------------------------------------------------
功能点：
1) 条码列只保留第一个；扩展条码保存剩余条码（逗号分隔），自动去空/去重。
2) 规格补全：当“主编码”有值且规格为空时，从“名称”提取(500ml/12*500ml/800g等)；提取不到再用“商品类型”；两者都有时组合为“类型 规格”。
3) 品牌长度限制：<=30字符，超长截断；空/None/'nan' 清空。
4) 重量必须为正数：若毛重<=0或缺失，则尝试从规格解析 g/kg 并换算为克(g)；规格是 ml/L 或无法解析则留空。
5) 商品状态：映射为“启用/禁用”。
6) 每个导出文件最多 max_rows 行（默认 1500）自动分包。

使用方法：
    python 商品导入清洗脚本.py \
        --src "导出saas商品详情5549981758728787637.xlsx" \
        --out_prefix "商品导入模版_清洗输出" \
        --max_rows 1500

依赖：pandas, openpyxl
"""

import argparse
import math
import re
from typing import List, Optional

import numpy as np
import pandas as pd


def pick(frame: pd.DataFrame, *names: List[str]) -> pd.Series:
    for n in names:
        if n in frame.columns:
            return frame[n]
    return pd.Series([np.nan] * len(frame))


def to_num(s: pd.Series) -> pd.Series:
    return pd.to_numeric(s, errors="coerce")


def extract_spec_from_name(name: str) -> str:
    """从名称中提取规格信息（优先多包装，其次单规格，再数量单位）"""
    if not isinstance(name, str):
        return ""
    # 12*500ml / 500ml*12
    m = re.findall(r"(\d+\s*[x×*]\s*\d+\s*(?:ml|mL|ML|l|L|g|G|kg|KG|片|粒|瓶|听|罐|袋|包|盒))", name, flags=re.IGNORECASE)
    if m:
        return m[0].replace(" ", "")
    # 500ml / 2L / 30g
    m2 = re.findall(r"(\d+(?:\.\d+)?\s*(?:ml|mL|ML|l|L|g|G|kg|KG|片|粒|瓶|听|罐|袋|包|盒))", name, flags=re.IGNORECASE)
    if m2:
        return m2[0].replace(" ", "")
    # 12瓶 / 24听
    m3 = re.findall(r"(\d+\s*(?:瓶|听|罐|袋|包|盒|片|粒))", name, flags=re.IGNORECASE)
    if m3:
        return m3[0].replace(" ", "")
    return ""


def parse_weight_from_spec(spec_text: str) -> Optional[float]:
    """从规格中解析 g/kg 重量，返回克(g)。解析不到返回 None。"""
    if not isinstance(spec_text, str):
        return None
    m = re.search(r"(\d+(?:\.\d+)?)\s*(kg|KG|千克|公斤|g|G|克)", spec_text)
    if not m:
        return None
    val = float(m.group(1))
    unit = m.group(2).lower()
    grams = None
    if unit in ["kg", "千克", "公斤"]:
        grams = val * 1000
    elif unit in ["g", "克"]:
        grams = val
    if grams is not None and grams > 0:
        return round(grams, 2)
    return None


def main(src: str, out_prefix: str, max_rows: int = 1500):
    xls_src = pd.ExcelFile(src)
    df_src = pd.read_excel(src, sheet_name=xls_src.sheet_names[0])

    # -------- 条码 & 扩展条码 --------
    barcode_raw = pick(df_src, "条码/简码", "条码", "商品条码").astype(str).str.replace(r"\.0$", "", regex=True)
    splits = barcode_raw.str.split(r"[ ,/|；;]")
    barcode_main = splits.str[0].fillna("")
    # 剩余条码去空、去重、去与主条码相同
    def rest_barcodes(lst, main):
        if not isinstance(lst, list) or len(lst) <= 1:
            return np.nan
        cleaned = []
        seen = set()
        for x in lst[1:]:
            v = str(x).strip()
            if not v or v == main or v in seen:
                continue
            seen.add(v)
            cleaned.append(v)
        return ",".join(cleaned) if cleaned else np.nan
    barcode_ext = [rest_barcodes(lst, main) for lst, main in zip(splits, barcode_main)]

    # -------- 基础映射 --------
    sell_mode = pick(df_src, "售卖方式").astype(str)
    is_weight = sell_mode.str.contains("称重", na=False) | sell_mode.str.contains("散称", na=False)
    is_count = ~is_weight

    unit = pick(df_src, "销售单位", "主单位", "单位")
    spec = pick(df_src, "规格").fillna(pick(df_src, "规格.1"))
    cost = to_num(pick(df_src, "平均进货价", "进货价"))
    price = to_num(pick(df_src, "门店零售价(元)", "销售价", "零售价"))
    member_price = to_num(pick(df_src, "门店会员价(元)", "会员价"))
    margin_rate = (price - cost) / price * 100
    margin_rate = margin_rate.replace([np.inf, -np.inf], np.nan).round(2)
    brand = pick(df_src, "商品品牌", "品牌").astype(str)
    vendor = pick(df_src, "供应商", "供货商")
    weight_col = to_num(pick(df_src, "毛重")).fillna(to_num(pick(df_src, "毛重.1")))
    stock_qty = to_num(pick(df_src, "库存", "库存量"))
    category = pick(df_src, "店内末级品类", "系统末级品类", "分类")
    goods_type = pick(df_src, "商品类型", "类型").fillna("")

    status_raw = pick(df_src, "上架状态", "商品状态").astype(str)
    enable_keys = ["上架", "在架", "销售中", "启用", "可售", "有效"]
    disable_keys = ["下架", "停售", "禁用", "停用", "不可售", "无效"]
    status_std = []
    for v in status_raw.fillna("").tolist():
        s = "启用"
        for k in disable_keys:
            if k in v:
                s = "禁用"
                break
        else:
            for k in enable_keys:
                if k in v:
                    s = "启用"
                    break
        status_std.append(s)
    status = pd.Series(status_std)

    points = pick(df_src, "是否参与积分").map(lambda x: "是" if str(x) in ["1", "是", "true", "True", "Y", "参加", "参与"] else ("否" if pd.notna(x) else np.nan))
    sku_no = pick(df_src, "货号")

    out = pd.DataFrame({
        "名称（必填）": pick(df_src, "商品名称", "名称"),
        "分类（必填）": category,
        "条码": barcode_main,
        "扩展条码": pd.Series(barcode_ext),
        "主编码": barcode_main,
        "规格": spec.astype(str),
        "主单位": unit,
        "库存量": stock_qty,
        "进货价（必填）": cost.round(4),
        "销售价（必填）": price.round(4),
        "毛利率": margin_rate,
        "批发价": pd.Series([np.nan]*len(df_src)),
        "会员价": member_price.round(4),
        "会员折扣": pd.Series([np.nan]*len(df_src)),
        "积分商品": points,
        "库存上限": pd.Series([np.nan]*len(df_src)),
        "库存下限": pd.Series([np.nan]*len(df_src)),
        "库位": pd.Series([np.nan]*len(df_src)),
        "品牌": brand,
        "供货商": vendor,
        "生产日期": pd.Series([np.nan]*len(df_src)),
        "保质期": pd.Series([np.nan]*len(df_src)),
        "拼音码": pd.Series([np.nan]*len(df_src)),
        "货号": sku_no,
        "自定义1": pd.Series([np.nan]*len(df_src)),
        "自定义2": pd.Series([np.nan]*len(df_src)),
        "自定义3": pd.Series([np.nan]*len(df_src)),
        "重量": weight_col,
        "是否称重": np.where(is_weight, "是", "否"),
        "是否传秤": pd.Series([np.nan]*len(df_src)),
        "是否计数商品": np.where(is_count, "是", "否"),
        "称编码": pd.Series([np.nan]*len(df_src)),
        "商品状态": status,
        "商品描述": pd.Series([np.nan]*len(df_src)),
        "标签": pd.Series([np.nan]*len(df_src)),
        "创建日期": pd.Series([np.nan]*len(df_src)),
    })

    # -------- 规格补全（主编码有值 & 规格为空） --------
    name_series = out["名称（必填）"].fillna("")
    type_series = goods_type.fillna("").astype(str)
    need_fill = (out["主编码"].notna() & (out["主编码"].astype(str).str.strip()!="")) & (
        out["规格"].isna() | (out["规格"].astype(str).str.lower().isin(["nan",""])))
    name_spec = name_series.astype(str).map(extract_spec_from_name)

    filled_spec = []
    for i in range(len(out)):
        if need_fill.iloc[i]:
            cand1 = name_spec.iloc[i]
            cand2 = type_series.iloc[i]
            if cand1 and cand2:
                cand = f"{cand2} {cand1}"
            elif cand1:
                cand = cand1
            elif cand2:
                cand = cand2
            else:
                nm = str(name_series.iloc[i])
                cand = nm[:12] if nm else "通用"
            filled_spec.append(cand)
        else:
            filled_spec.append(out["规格"].iloc[i])
    out.loc[need_fill, "规格"] = pd.Series(filled_spec)[need_fill.index[need_fill]]

    # -------- 品牌 限制30字符 & 清理 --------
    out["品牌"] = out["品牌"].astype(str).apply(lambda s: s[:30] if isinstance(s, str) else s)
    out["品牌"] = out["品牌"].replace({"nan": np.nan, "NaN": np.nan, "None": np.nan, "": np.nan})

    # -------- 重量：正数约束 & 规格解析 --------
    fixed_weight = []
    for i, w in enumerate(out["重量"]):
        if pd.isna(w) or (isinstance(w, (int, float)) and w <= 0):
            grams = parse_weight_from_spec(out["规格"].iloc[i])
            fixed_weight.append(grams if grams is not None and grams > 0 else np.nan)
        else:
            fixed_weight.append(w)
    out["重量"] = pd.to_numeric(pd.Series(fixed_weight), errors="coerce")

    # -------- 基础清洗 --------
    out = out[~out["名称（必填）"].isna() & ~out["分类（必填）"].isna()].copy()
    for c in ["进货价（必填）","销售价（必填）","会员价","批发价","库存量","重量","毛利率"]:
        out[c] = pd.to_numeric(out[c], errors="coerce")

    # -------- 分包导出 --------
    n = len(out)
    num_parts = (n + max_rows - 1) // max_rows if n else 1
    paths = []
    for i in range(num_parts):
        start = i * max_rows
        end = min((i + 1) * max_rows, n)
        part = out.iloc[start:end].copy()
        path = f"{out_prefix}_part{i+1}.xlsx"
        with pd.ExcelWriter(path, engine="openpyxl") as writer:
            part.to_excel(writer, sheet_name="导入模版", index=False)
        paths.append(path)

    print(f"共导出 {n} 条，分成 {len(paths)} 个文件：")
    for p in paths:
        print(" -", p)


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--src", required=True, help="JD 万商导出 Excel 路径（.xlsx）")
    parser.add_argument("--out_prefix", default="商品导入模版_清洗输出", help="导出文件前缀")
    parser.add_argument("--max_rows", type=int, default=1500, help="每个文件的最大行数")
    args = parser.parse_args()
    main(args.src, args.out_prefix, args.max_rows)

with open(script_path, "w", encoding="utf-8") as f:
    f.write(script_code)
