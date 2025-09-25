#!/bin/bash

# 1) 安装依赖（如本地未安装）
pip3 install pandas openpyxl

# 2) 运行脚本
python3 tools/JD2YinBaoProduct.py \
  --src "data/导出saas商品详情5549981758728787637.xlsx" \
  --out_prefix "data/商品导入模版_清洗输出" \
  --max_rows 1500