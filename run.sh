#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
run.py — 一键运行：
1) 数据清洗导出（银豹导入模版）
2) 商品图片下载
3) 箱单关系转换 + 缺失条码校验（默认优先使用“修复后的关系表”）
4) 全部执行（已调整为先修复再转换）
5) 按名称修复箱单关系表里的条码

依赖：
  - Python 3.8+
  - pip install pandas openpyxl
"""

import argparse
import subprocess
import sys
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
TOOLS_DIR = BASE_DIR / "tools"

# 默认文件
SRC_FILE = DATA_DIR / "导出saas商品详情5549981758891505190.xlsx"
OUT_PREFIX = DATA_DIR / "商品导入模版_清洗输出"

PACKAGE_FILE = DATA_DIR / "laiyijia001箱单关系_1758891484458.xlsx"      # 原始关系表
PACKAGE_OUT = DATA_DIR / "laiyijia001箱单关系_converted.xlsx"
MISSING_REPORT = DATA_DIR / "缺失条码清单.xlsx"

FIXED_REL_OUT = DATA_DIR / "laiyijia001箱单关系_修复后.xlsx"            # 修复后的关系表（优先使用）
FIX_LOG_OUT   = DATA_DIR / "laiyijia001箱单关系_修复日志.xlsx"

PYTHON = sys.executable  # 使用当前解释器

def run_cmd(args):
    print(">>>", " ".join(str(a) for a in args))
    res = subprocess.run(args, check=False)
    if res.returncode != 0:
        print(f"⚠️ 命令返回非零状态：{res.returncode}")
    return res.returncode

def task_clean_export(src=SRC_FILE, out_prefix=OUT_PREFIX, max_rows=1500):
    return run_cmd([PYTHON, str(TOOLS_DIR / "JD2YinBaoProduct.py"),
                    "--src", str(src),
                    "--out_prefix", str(out_prefix),
                    "--max_rows", str(max_rows)])

def task_download_images(src=SRC_FILE, out_dir=DATA_DIR / "images"):
    return run_cmd([PYTHON, str(TOOLS_DIR / "JD2YinBaoDownloadProductImage.py"),
                    "--src", str(src),
                    "--out_dir", str(out_dir)])

def resolve_relation_for_package(user_specified: Path|None, fixed_path: Path = FIXED_REL_OUT, fallback: Path = PACKAGE_FILE) -> Path:
    """
    优先使用修复后的关系表；如果用户通过 --package 指定了路径，则强制使用用户指定；
    若修复后的文件不存在，则回退到原始关系表。
    """
    if user_specified is not None:
        print(f"ℹ️ 使用用户指定的关系表：{user_specified}")
        return user_specified
    if fixed_path.exists():
        print(f"ℹ️ 发现修复后的关系表，优先使用：{fixed_path}")
        return fixed_path
    print(f"ℹ️ 未找到修复后的关系表，回退使用原始：{fallback}")
    return fallback

def task_package_relation(package_src: Path, package_out: Path = PACKAGE_OUT, products: Path = SRC_FILE, missing_report: Path = MISSING_REPORT):
    return run_cmd([PYTHON, str(TOOLS_DIR / "JD2YinBaoPackageRelation.py"),
                    "--src", str(package_src),
                    "--out", str(package_out),
                    "--products", str(products),
                    "--missing-report", str(missing_report)])

def task_fix_by_name(relation: Path, products: Path = SRC_FILE, out_fixed: Path = FIXED_REL_OUT, log_out: Path = FIX_LOG_OUT):
    # 脚本可放 tools/ 下；若你放在根目录，也可兼容
    fix_script = TOOLS_DIR / "JDFixPackageRelationByName.py"
    if not fix_script.exists():
        fix_script = BASE_DIR / "JDFixPackageRelationByName.py"
    return run_cmd([PYTHON, str(fix_script),
                    "--relation", str(relation),
                    "--products", str(products),
                    "--out", str(out_fixed),
                    "--log", str(log_out)])

def interactive_menu():
    print("===================================")
    print(" 银豹导入辅助工具 (run.py)")
    print("===================================")
    print("1) 数据清洗导出 (生成银豹导入模版)")
    print("2) 商品图片下载 (按条码命名到 data/images/)")
    print("3) 箱单关系转换 + 校验缺失条码（默认优先使用修复后的关系表）")
    print("4) 全部执行（先修复再转换）")
    print("5) 按名称修复箱单关系表条码")
    print("0) 退出")
    choice = input("请输入选项 (0-5): ").strip()
    return choice

def main():
    parser = argparse.ArgumentParser(description="银豹工具一键运行器")
    parser.add_argument("--task", choices=["clean", "images", "package", "all", "fixnames"],
                        help="不启用交互菜单，直接执行指定任务")
    parser.add_argument("--src", type=str, help="商品资料 Excel 路径（默认 data/导出saas商品详情*.xlsx）")
    parser.add_argument("--package", type=str, help="箱单关系表 Excel 路径（默认优先 data/…_修复后.xlsx，其次 data/…箱单关系.xlsx）")
    parser.add_argument("--out-prefix", type=str, help="清洗导出前缀（默认 data/商品导入模版_清洗输出）")
    parser.add_argument("--package-out", type=str, help="箱单关系转换输出（默认 data/laiyijia001箱单关系_converted.xlsx）")
    parser.add_argument("--missing-report", type=str, help="缺失条码清单输出（默认 data/缺失条码清单.xlsx）")
    parser.add_argument("--fixed-out", type=str, help="按名称修复后的关系表输出（默认 data/laiyijia001箱单关系_修复后.xlsx）")
    parser.add_argument("--fix-log", type=str, help="修复日志输出（默认 data/laiyijia001箱单关系_修复日志.xlsx）")

    args = parser.parse_args()

    src = Path(args.src) if args.src else SRC_FILE
    # 这里不直接用 args.package；交给 resolve_relation_for_package 决定最终来源
    user_package = Path(args.package) if args.package else None

    out_prefix = Path(args.out_prefix) if args.out_prefix else OUT_PREFIX
    package_out = Path(args.package_out) if args.package_out else PACKAGE_OUT
    missing_report = Path(args.missing_report) if args.missing_report else MISSING_REPORT
    fixed_out = Path(args.fixed_out) if args.fixed_out else FIXED_REL_OUT
    fix_log = Path(args.fix_log) if args.fix_log else FIX_LOG_OUT

    if args.task:
        if args.task == "clean":
            task_clean_export(src, out_prefix)
        elif args.task == "images":
            task_download_images(src)
        elif args.task == "package":
            rel_for_pkg = resolve_relation_for_package(user_package, fixed_out, PACKAGE_FILE)
            task_package_relation(rel_for_pkg, package_out, src, missing_report)
        elif args.task == "fixnames":
            # 修复输入：优先用（用户指定 / 原始），输出写到 fixed_out
            rel_input = user_package if user_package is not None else PACKAGE_FILE
            task_fix_by_name(rel_input, src, fixed_out, fix_log)
        elif args.task == "all":
            # 顺序：清洗 -> 图片 -> 按名称修复 -> 关系转换（默认已能使用修复后的表）
            task_clean_export(src, out_prefix)
            task_download_images(src)
            rel_input = user_package if user_package is not None else PACKAGE_FILE
            task_fix_by_name(rel_input, src, fixed_out, fix_log)
            rel_for_pkg = resolve_relation_for_package(user_package, fixed_out, PACKAGE_FILE)
            task_package_relation(rel_for_pkg, package_out, src, missing_report)
        return

    # 交互菜单
    choice = interactive_menu()
    if choice == "1":
        task_clean_export(src, out_prefix)
    elif choice == "2":
        task_download_images(src)
    elif choice == "3":
        rel_for_pkg = resolve_relation_for_package(user_package, fixed_out, PACKAGE_FILE)
        task_package_relation(rel_for_pkg, package_out, src, missing_report)
    elif choice == "4":
        task_clean_export(src, out_prefix)
        task_download_images(src)
        rel_input = user_package if user_package is not None else PACKAGE_FILE
        task_fix_by_name(rel_input, src, fixed_out, fix_log)
        rel_for_pkg = resolve_relation_for_package(user_package, fixed_out, PACKAGE_FILE)
        task_package_relation(rel_for_pkg, package_out, src, missing_report)
    elif choice == "5":
        rel_input = user_package if user_package is not None else PACKAGE_FILE
        task_fix_by_name(rel_input, src, fixed_out, fix_log)
    elif choice == "0":
        print("退出")
    else:
        print("无效的选项")

if __name__ == "__main__":
    main()
