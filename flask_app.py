"""
Excel 数据提取 API (Flask版本 - 适用于PythonAnywhere)
用于从Excel文件中提取指定列的数据
"""

from flask import Flask, request, jsonify
from flask_cors import CORS
import pandas as pd
from io import BytesIO
import base64
import requests
from datetime import datetime

app = Flask(__name__)
CORS(app)  # 允许跨域请求

# 定义两种Excel格式的列映射
EXCEL_FORMATS = {
    "边坡检查表": {
        "identifier_columns": ["序号", "桩号"],
        "extract_columns": ["序号", "桩号", "检查时间", "责任人", "问题描述", "整改措施"]
    },
    "测量成果表": {
        "identifier_columns": ["点号", "X坐标"],
        "extract_columns": ["点号", "X坐标", "Y坐标", "高程", "备注"]
    }
}

def detect_excel_format(df):
    """检测Excel文件格式"""
    columns = set(df.columns.tolist())
    
    for format_name, format_info in EXCEL_FORMATS.items():
        id_cols = set(format_info["identifier_columns"])
        if id_cols.issubset(columns):
            return format_name, format_info["extract_columns"]
    
    return None, list(df.columns)

def read_excel_from_base64(base64_string):
    """从Base64字符串读取Excel"""
    excel_bytes = base64.b64decode(base64_string)
    return pd.read_excel(BytesIO(excel_bytes))

def read_excel_from_url(url):
    """从URL读取Excel"""
    response = requests.get(url, timeout=30)
    response.raise_for_status()
    return pd.read_excel(BytesIO(response.content))

@app.route('/', methods=['GET'])
def health_check():
    """健康检查端点"""
    return jsonify({
        "status": "healthy",
        "service": "Excel Extract API",
        "version": "1.0.0",
        "timestamp": datetime.now().isoformat()
    })

@app.route('/extract', methods=['POST'])
def extract_excel():
    """从Excel文件提取数据"""
    try:
        data = request.get_json()
        
        if not data:
            return jsonify({
                "success": False,
                "message": "请求体为空"
            }), 400
        
        file_url = data.get('file_url')
        file_base64 = data.get('file_base64')
        custom_columns = data.get('columns')
        
        # 读取Excel文件
        if file_base64:
            df = read_excel_from_base64(file_base64)
        elif file_url:
            df = read_excel_from_url(file_url)
        else:
            return jsonify({
                "success": False,
                "message": "请提供 file_url 或 file_base64"
            }), 400
        
        # 检测格式并确定要提取的列
        detected_format, default_columns = detect_excel_format(df)
        
        # 使用自定义列或默认列
        columns_to_extract = custom_columns if custom_columns else default_columns
        
        # 过滤存在的列
        available_columns = [col for col in columns_to_extract if col in df.columns]
        
        if not available_columns:
            return jsonify({
                "success": False,
                "message": f"未找到指定的列。可用列: {df.columns.tolist()}"
            }), 400
        
        # 提取数据
        extracted_df = df[available_columns]
        
        # 生成新的Excel文件
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            extracted_df.to_excel(writer, index=False, sheet_name='提取数据')
        
        excel_bytes = output.getvalue()
        result_base64 = base64.b64encode(excel_bytes).decode('utf-8')
        
        return jsonify({
            "success": True,
            "message": f"成功提取 {len(extracted_df)} 行数据",
            "detected_format": detected_format or "未知格式",
            "extracted_columns": available_columns,
            "row_count": len(extracted_df),
            "file_base64": result_base64
        })
        
    except Exception as e:
        return jsonify({
            "success": False,
            "message": f"处理失败: {str(e)}"
        }), 500

# PythonAnywhere 需要这个
if __name__ == '__main__':
    app.run(debug=True)
