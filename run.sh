#!/bin/bash
# 一键运行数据清洗和图片下载
# 提供菜单让用户选择执行哪个功能

BASE_DIR=$(dirname "$0")
DATA_DIR="$BASE_DIR/data"
TOOLS_DIR="$BASE_DIR/tools"

SRC_FILE="$DATA_DIR/导出saas商品详情5549981758728787637.xlsx"
OUT_PREFIX="$DATA_DIR/商品导入模版_清洗输出"

echo "==================================="
echo " 银豹导入辅助工具"
echo "==================================="
echo "请选择要执行的功能："
echo "1) 数据清洗导出 (生成银豹导入模版)"
echo "2) 商品图片下载 (按条码命名到 data/images/)"
echo "3) 两个都执行"
echo "0) 退出"
echo "-----------------------------------"
read -p "请输入选项 (0-3): " choice

case $choice in
  1)
    echo ">>> 执行数据清洗导出..."
    python3 "$TOOLS_DIR/JD2YinBaoProduct.py" \
      --src "$SRC_FILE" \
      --out_prefix "$OUT_PREFIX" \
      --max_rows 1500
    ;;
  2)
    echo ">>> 执行商品图片下载..."
    python3 "$TOOLS_DIR/JD2YinBaoDownloadProductImage.py" \
      --src "$SRC_FILE" \
      --out_dir "$DATA_DIR/images"
    ;;
  3)
    echo ">>> 执行数据清洗导出..."
    python3 "$TOOLS_DIR/JD2YinBaoProduct.py" \
      --src "$SRC_FILE" \
      --out_prefix "$OUT_PREFIX" \
      --max_rows 1500

    echo ""
    echo ">>> 执行商品图片下载..."
    python3 "$TOOLS_DIR/JD2YinBaoDownloadProductImage.py" \
      --src "$SRC_FILE" \
      --out_dir "$DATA_DIR/images"
    ;;
  0)
    echo "退出"
    exit 0
    ;;
  *)
    echo "无效的选项"
    exit 1
    ;;
esac

echo ""
echo "=== 完成 ✅ ==="
