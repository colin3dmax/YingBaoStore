# 1) 安装依赖（如本地未安装）
pip install pandas openpyxl

# 2) 运行脚本
python 商品导入清洗脚本.py \
  --src "导出saas商品详情5549981758728787637.xlsx" \
  --out_prefix "商品导入模版_清洗输出" \
  --max_rows 1500