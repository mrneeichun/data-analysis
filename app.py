# ==================== Flask API 后端 ====================
# 部署到腾讯云函数计算（免费）
from flask import Flask, request, jsonify
import pandas as pd
import io
import sys
import os

# 添加当前目录到路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# 导入原有的数据处理模块
from i3000 import clean_i3000
from i6000 import clean_i6000
from 术前 import analyze_术前
from 肿瘤 import analyze_肿瘤
from 甲功 import analyze_甲功

app = Flask(__name__)

# 跨域支持
@app.after_request
def after_request(response):
    response.headers['Access-Control-Allow-Origin'] = '*'
    response.headers['Access-Control-Allow-Methods'] = 'GET, POST, OPTIONS'
    response.headers['Access-Control-Allow-Headers'] = 'Content-Type'
    return response

@app.route('/api/analyze', methods=['POST', 'OPTIONS'])
def analyze():
    """分析接口"""
    if request.method == 'OPTIONS':
        return jsonify({'status': 'ok'})
    
    try:
        # 获取上传的文件
        file = request.files.get('file')
        if not file:
            return jsonify({'error': '请上传Excel文件'}), 400
        
        # 获取参数
        machine_type = request.form.get('machine_type', 'i3000')
        project_type = request.form.get('project_type', '术前八项')
        
        # 读取Excel文件
        file_bytes = file.read()
        df = pd.read_excel(io.BytesIO(file_bytes), engine='openpyxl')
        
        # 根据仪器类型清洗数据
        if machine_type == "i3000":
            df = clean_i3000(df)
        elif machine_type == "i6000":
            df = clean_i6000(df)
        
        # 根据项目类型分析数据
        if project_type == "术前八项":
            summary_df, mode_df, sample_map, index_col = analyze_术前(df)
            result = {
                'raw_data': df.to_dict(orient='records'),
                'summary': summary_df.to_dict(orient='records'),
                'mode_distribution': mode_df.to_dict(orient='records'),
                'type': '术前八项'
            }
            
        elif project_type == "甲功":
            summary_df = analyze_甲功(df)
            result = {
                'raw_data': df.to_dict(orient='records'),
                'summary': summary_df.to_dict(orient='records'),
                'type': '甲功'
            }
            
        elif project_type == "肿瘤":
            summary_df = analyze_肿瘤(df)
            result = {
                'raw_data': df.to_dict(orient='records'),
                'summary': summary_df.to_dict(orient='records'),
                'type': '肿瘤'
            }
        else:
            return jsonify({'error': '不支持的项目类型'}), 400
        
        return jsonify({
            'status': 'success',
            'data': result
        })
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/health', methods=['GET'])
def health():
    """健康检查"""
    return jsonify({'status': 'ok'})

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
