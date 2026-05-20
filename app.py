"""
箱唛生成系统 - 后端服务
功能：上传Excel和Word模板，自动生成包含多页箱唛的Word文档
核心：在ZIP层面操作docx，完整保留原始字体、图片、样式
"""

from flask import Flask, render_template, request, jsonify, send_file
import os
import re
import pandas as pd
import io
import zipfile

import tempfile

UPLOAD_DIR = os.path.join(tempfile.gettempdir(), 'carton_label_uploads')
GENERATED_DIR = os.path.join(tempfile.gettempdir(), 'carton_label_generated')

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_DIR
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024
app.config['GENERATED_FOLDER'] = GENERATED_DIR

os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['GENERATED_FOLDER'], exist_ok=True)


def find_variables_in_template(doc_path):
    """扫描Word模板，找出所有【X】格式的变量"""
    with zipfile.ZipFile(doc_path, 'r') as z:
        doc_xml = z.read('word/document.xml').decode('utf-8')
    variables = set(re.findall(r'【([^】]+)】', doc_xml))
    return sorted(list(variables))


def read_excel_data(excel_path):
    """读取Excel，返回DataFrame"""
    df = pd.read_excel(excel_path, dtype=str)  # 全部读为字符串，保留原始显示
    df.columns = df.columns.str.strip().str.replace('\n', '')
    df = df.dropna(how='all')
    df = df.reset_index(drop=True)
    return df


def format_excel_value(raw_value):
    """
    格式化Excel值：整数去掉小数点，保留原始显示
    raw_value 是字符串（因为dtype=str）
    """
    if pd.isna(raw_value) or str(raw_value).strip() == '':
        return ''
    val = str(raw_value).strip()
    # 尝试判断是否为整数（如 "8.0" → "8"）
    try:
        num = float(val)
        if num == int(num):
            return str(int(num))
    except (ValueError, OverflowError):
        pass
    return val


def replace_text_in_xml(xml_str, row_data, variables_mapping):
    """在XML字符串中替换变量占位符"""
    for var_name, excel_col in variables_mapping.items():
        placeholder = f'【{var_name}】'
        if placeholder in xml_str:
            if excel_col in row_data.index:
                value = format_excel_value(row_data[excel_col])
                xml_str = xml_str.replace(placeholder, value)
    return xml_str


def generate_single_docx(template_path, row_data, variables_mapping):
    """
    基于模板生成单个docx的bytes
    完整复制模板ZIP中的所有文件，仅修改document.xml中的变量
    """
    with zipfile.ZipFile(template_path, 'r') as template_zip:
        # 读取并替换 document.xml
        doc_xml = template_zip.read('word/document.xml').decode('utf-8')
        doc_xml = replace_text_in_xml(doc_xml, row_data, variables_mapping)

        # 构建 output bytes
        output_buf = io.BytesIO()
        with zipfile.ZipFile(output_buf, 'w', zipfile.ZIP_DEFLATED) as out_zip:
            for item in template_zip.infolist():
                if item.filename == 'word/document.xml':
                    out_zip.writestr(item, doc_xml.encode('utf-8'))
                else:
                    out_zip.writestr(item, template_zip.read(item.filename))
        output_buf.seek(0)
        return output_buf.read()


def generate_merged_docx(template_path, df, variables_mapping):
    """
    生成合并的多页docx
    在document.xml字符串层面复制body内容，添加分页符，替换变量
    完整保留原始字体、图片、样式、关系引用
    """
    with zipfile.ZipFile(template_path, 'r') as template_zip:
        doc_xml = template_zip.read('word/document.xml').decode('utf-8')

        # 提取 <w:body>...</w:body> 内部内容
        body_match = re.search(r'(<w:body>)(.*?)(</w:body>)', doc_xml, re.DOTALL)
        if not body_match:
            body_match = re.search(r'(<[^>]*body>)(.*?)(</[^>]*body>)', doc_xml, re.DOTALL)
        if not body_match:
            raise ValueError("无法解析文档body")

        body_open = body_match.group(1)
        body_content = body_match.group(2)
        body_close = body_match.group(3)

        # 分页符XML
        page_break_xml = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'

        # 拼接所有页面
        all_pages = []
        for idx, row in df.iterrows():
            page_xml = replace_text_in_xml(body_content, row, variables_mapping)
            all_pages.append(page_xml)
            if idx < len(df) - 1:
                all_pages.append(page_break_xml)

        new_body_content = ''.join(all_pages)
        new_doc_xml = doc_xml.replace(
            body_open + body_content + body_close,
            body_open + new_body_content + body_close
        )

        # 写入输出（完整复制模板所有文件，仅替换document.xml）
        output_buf = io.BytesIO()
        with zipfile.ZipFile(output_buf, 'w', zipfile.ZIP_DEFLATED) as out_zip:
            for item in template_zip.infolist():
                if item.filename == 'word/document.xml':
                    out_zip.writestr(item, new_doc_xml.encode('utf-8'))
                else:
                    out_zip.writestr(item, template_zip.read(item.filename))
        output_buf.seek(0)
        return output_buf.read()


def generate_carton_labels(excel_path, template_path, variables_mapping):
    """生成多页合并Word文档"""
    df = read_excel_data(excel_path)
    doc_bytes = generate_merged_docx(template_path, df, variables_mapping)

    timestamp = pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')
    output_filename = f'箱唛_{timestamp}.docx'
    output_path = os.path.join(app.config['GENERATED_FOLDER'], output_filename)

    with open(output_path, 'wb') as f:
        f.write(doc_bytes)

    return output_path, len(df)


def generate_zip_download(excel_path, template_path, variables_mapping):
    """生成ZIP压缩包（每个箱唛单独一个文件）"""
    df = read_excel_data(excel_path)

    timestamp = pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')
    zip_filename = f'箱唛_全部_{timestamp}.zip'
    zip_path = os.path.join(app.config['GENERATED_FOLDER'], zip_filename)

    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for idx, row in df.iterrows():
            doc_bytes = generate_single_docx(template_path, row, variables_mapping)
            file_name = f'箱唛_{idx + 1}.docx'
            zf.writestr(file_name, doc_bytes)

    return zip_path, len(df)


# ==================== Flask 路由 ====================

@app.route('/')
def index():
    return render_template('index.html')


@app.route('/api/analyze', methods=['POST'])
def analyze():
    """分析上传的文件"""
    try:
        excel_file = request.files.get('excel')
        template_file = request.files.get('template')

        if not excel_file or not template_file:
            return jsonify({'error': '请上传Excel和Word模板文件'}), 400

        excel_path = os.path.join(app.config['UPLOAD_FOLDER'], 'data.xlsx')
        template_path = os.path.join(app.config['UPLOAD_FOLDER'], 'template.docx')

        excel_file.save(excel_path)
        template_file.save(template_path)

        variables = find_variables_in_template(template_path)

        df = pd.read_excel(excel_path, dtype=str)
        df.columns = df.columns.str.strip().str.replace('\n', '')
        excel_columns = df.columns.tolist()
        row_count = len(df.dropna(how='all'))

        return jsonify({
            'variables': variables,
            'excel_columns': excel_columns,
            'row_count': row_count
        })

    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/api/generate', methods=['POST'])
def generate():
    """生成箱唛文档（多页合并）"""
    try:
        data = request.get_json()
        variables_mapping = data.get('variables_mapping', {})

        excel_path = os.path.join(app.config['UPLOAD_FOLDER'], 'data.xlsx')
        template_path = os.path.join(app.config['UPLOAD_FOLDER'], 'template.docx')

        if not os.path.exists(excel_path) or not os.path.exists(template_path):
            return jsonify({'error': '文件未找到，请重新上传'}), 400

        output_path, doc_count = generate_carton_labels(
            excel_path, template_path, variables_mapping
        )

        return jsonify({
            'success': True,
            'message': f'成功生成 {doc_count} 页箱唛',
            'download_url': f'/api/download/{os.path.basename(output_path)}'
        })

    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/api/download/<filename>')
def download(filename):
    """下载生成的文件"""
    file_path = os.path.join(app.config['GENERATED_FOLDER'], filename)
    if os.path.exists(file_path):
        return send_file(file_path, as_attachment=True)
    return jsonify({'error': '文件不存在'}), 404


@app.route('/api/generate-zip', methods=['POST'])
def generate_zip():
    """生成ZIP文件（逐个下载模式）"""
    try:
        data = request.get_json()
        variables_mapping = data.get('variables_mapping', {})

        excel_path = os.path.join(app.config['UPLOAD_FOLDER'], 'data.xlsx')
        template_path = os.path.join(app.config['UPLOAD_FOLDER'], 'template.docx')

        if not os.path.exists(excel_path) or not os.path.exists(template_path):
            return jsonify({'error': '文件未找到，请重新上传'}), 400

        zip_path, doc_count = generate_zip_download(
            excel_path, template_path, variables_mapping
        )

        return jsonify({
            'success': True,
            'message': f'成功生成 {doc_count} 个箱唛文档',
            'download_url': f'/api/download/{os.path.basename(zip_path)}'
        })

    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


if __name__ == '__main__':
    print('=' * 50)
    print('📦 箱唛生成系统')
    print('=' * 50)
    print('请在浏览器中打开: http://localhost:5000')
    print('按 Ctrl+C 停止服务')
    print('=' * 50)
    app.run(debug=True, host='0.0.0.0', port=5000)
