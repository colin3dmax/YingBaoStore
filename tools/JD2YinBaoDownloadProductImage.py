#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
批量下载商品图片脚本 (JD2YinBaoDownloadProductImage.py)

目录结构：
  项目根目录/
    ├─ data/   ← 输入/输出数据 (含下载的图片)
    │    ├─ 导出saas商品详情.xlsx
    │    └─ images/   ← 下载的图片和报表
    └─ tools/  ← 脚本

使用方法：
  cd tools
  python JD2YinBaoDownloadProductImage.py --src "../data/导出saas商品详情5549981758728787637.xlsx"

依赖：
  pip install pandas pillow requests tqdm
"""

import argparse
import io
import os
import re
import sys
from urllib.parse import urlparse

import pandas as pd
import requests
from PIL import Image, ImageOps
from tqdm import tqdm

ALLOWED_EXTS = {".jpg", ".jpeg", ".png"}
ALLOWED_MIMES = {"image/jpeg", "image/png"}

BARCODE_CANDIDATES = ["条码/简码", "条码", "商品条码", "主编码", "EAN", "UPC", "Barcode"]
URL_CANDIDATES = ["图片", "商品图片", "商品主图", "图片URL", "图片url", "主图", "图片链接", "图片地址"]

def guess_col(df: pd.DataFrame, candidates):
    for name in candidates:
        if name in df.columns:
            return name
    return None

def split_urls(s: str):
    if not isinstance(s, str):
        return []
    return [p for p in re.split(r"[,\s;；\n\r]+", s.strip()) if p]

def ensure_dir(d):
    os.makedirs(d, exist_ok=True)

def sanitize_barcode(bc: str) -> str:
    """从条码单元格中取第一个条码作为文件名，保留前导零，去掉 .0 尾巴和非法字符。"""
    if bc is None:
        return ""
    s = str(bc).strip()
    # Excel 数值型导出常见的 1234567890.0 情况
    s = re.sub(r"\.0$", "", s)

    # 先统一分隔符（中文逗号/分号等），再按常见分隔符切分，取第一个
    s_norm = re.sub(r"[，；、]", ",", s)
    first = re.split(r"[ ,/|;]+", s_norm)[0].strip() if s_norm else ""

    # 去除文件名非法字符
    first = re.sub(r"[\\/:*?\"<>|]", "", first)
    return first

def fetch_image(url: str, timeout=20):
    try:
        resp = requests.get(url, timeout=timeout, stream=True)
        resp.raise_for_status()
        content_type = resp.headers.get("Content-Type", "").split(";")[0].strip().lower()
        if content_type == "image/png":
            ext = ".png"
        else:
            ext = ".jpg"
        return resp.content, ext, None
    except Exception as e:
        return None, None, f"下载失败: {e}"

def process_image(content: bytes, size=750, max_bytes=3*1024*1024, ext=".jpg"):
    try:
        im = Image.open(io.BytesIO(content))
        if im.mode not in ("RGB", "RGBA"):
            im = im.convert("RGB")
        im = ImageOps.contain(im, (size, size))
        canvas = Image.new("RGB", (size, size), (255, 255, 255))
        x = (size - im.width) // 2
        y = (size - im.height) // 2
        canvas.paste(im, (x, y))

        out_ext = ext.lower() if ext.lower() in ALLOWED_EXTS else ".jpg"
        quality = 95
        while quality >= 50:
            buf = io.BytesIO()
            if out_ext in [".jpg", ".jpeg"]:
                canvas.save(buf, format="JPEG", quality=quality, optimize=True)
            else:
                canvas.save(buf, format="PNG", optimize=True)
                if buf.tell() > max_bytes:
                    out_ext = ".jpg"
                    continue
            if buf.tell() <= max_bytes:
                return buf.getvalue(), out_ext, None
            quality -= 5
        return None, None, "压缩仍超过限制"
    except Exception as e:
        return None, None, f"图片处理失败: {e}"

def save_bytes(data: bytes, path: str):
    with open(path, "wb") as f:
        f.write(data)

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--src", required=True, help="输入文件（Excel/CSV），默认放在 ../data 下")
    parser.add_argument("--out_dir", default="../data/images", help="输出图片目录（默认 ../data/images）")
    parser.add_argument("--barcode_col", default=None, help="条码列名")
    parser.add_argument("--url_col", default=None, help="图片URL列名")
    args = parser.parse_args()

    if args.src.lower().endswith(".csv"):
        df = pd.read_csv(args.src)
    else:
        xls = pd.ExcelFile(args.src)
        df = pd.read_excel(xls, sheet_name=xls.sheet_names[0])

    bc_col = args.barcode_col or guess_col(df, BARCODE_CANDIDATES)
    url_col = args.url_col or guess_col(df, URL_CANDIDATES)
    if not bc_col or not url_col:
        print("无法识别条码或图片列，请用 --barcode_col --url_col 指定")
        sys.exit(1)

    ensure_dir(args.out_dir)

    report = []
    for idx in tqdm(range(len(df)), desc="下载中"):
        bc = sanitize_barcode(df.iloc[idx][bc_col])
        urls = split_urls(str(df.iloc[idx][url_col]))
        if not bc or not urls:
            continue
        for j, u in enumerate(urls, start=1):
            content, ext, err = fetch_image(u)
            if err:
                report.append([bc, u, "fail", err])
                continue
            processed, out_ext, perr = process_image(content, ext=ext)
            if perr:
                report.append([bc, u, "fail", perr])
                continue
            fn = f"{bc}{out_ext}" if j == 1 else f"{bc}_{j}{out_ext}"
            path = os.path.join(args.out_dir, fn)
            save_bytes(processed, path)
            report.append([bc, u, "ok", path])

    rep_path = os.path.join(args.out_dir, "download_report.csv")
    pd.DataFrame(report, columns=["barcode", "url", "status", "info"]).to_csv(rep_path, index=False, encoding="utf-8-sig")
    print(f"\n完成 ✅ 图片目录: {args.out_dir}\n报表: {rep_path}")

if __name__ == "__main__":
    main()
