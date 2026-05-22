#!/bin/bash
set -e

# CI 构建脚本 (Gitee Go)
echo "安装依赖..."
pip install -r requirements.txt

echo "验证导入..."
python -c "from converter import convert_pdf_to_editable_docx, convert_pdf_to_office2007_docx; print('Import OK')"

echo "验证 Flask 应用..."
python -c "from app import app; print('Flask app OK')"

echo "构建完成"