import pandas as pd
from flask import Flask, request, render_template, send_from_directory, jsonify,send_file
import os

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

@app.route('/')
def index():
    return send_from_directory('.', 'index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({"error": "文件未上传"}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({"error": "未选择文件"}), 400

    filepath = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
    file.save(filepath)

    return jsonify({"message": "文件上传成功", "filename": file.filename}), 200

@app.route('/process', methods=['POST'])
def process_data():
    data = request.json
    filename = data['filename']
    choice = data['choice']
    custom_name = data['custom_name']

    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)

    # 读取 Excel 文件
    if filename.endswith('.xlsx'):
        df = pd.read_excel(filepath, engine='openpyxl')
    else:
        df = pd.read_excel(filepath)

    # 根据选择筛选数据
    if choice == "option1":
        filtered_data = df[(df['总工资'] > 100000) & (df['总工资'] < 200000) & (df['级别'] < 5)]
    elif choice == "option2":
        filtered_data = df[(df['总工资'] > 200000) & (df['总工资'] < 250000)]
    elif choice == "option3":
        filtered_data = df[(df['总工资'] < 100000)]

    # 计算工资的平均值
    avg_salary = filtered_data['总工资'].mean() if not filtered_data.empty else 0
    avg_row = pd.DataFrame({'薪资平均值': [avg_salary]})
    filtered_data = pd.concat([filtered_data, avg_row], ignore_index=True)

    # 导出处理后的数据
    output_file = f"{custom_name}.xlsx"
    filtered_data.to_excel(output_file, index=False)

    return jsonify({"download_link": output_file}), 200

@app.route('/download/<filename>', methods=['GET'])
def download_file(filename):
    return send_file(filename, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
