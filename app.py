"""
箱唛生成系统 - Gradio 版本（适配 Hugging Face Spaces）
功能：上传Excel和Word模板，自动生成包含多页箱唛的Word文档
"""

import os
import re
import tempfile
import pandas as pd
import io
import zipfile
import gradio as gr

# ==================== 核心逻辑 ====================

def find_variables_in_template(doc_path):
    """扫描Word模板，找出所有【X】格式的变量"""
    with zipfile.ZipFile(doc_path, 'r') as z:
        doc_xml = z.read('word/document.xml').decode('utf-8')
    variables = set(re.findall(r'【([^】]+)】', doc_xml))
    return sorted(list(variables))


def read_excel_data(excel_path):
    """读取Excel文件"""
    df = pd.read_excel(excel_path, dtype=str)
    df.columns = df.columns.str.strip().str.replace('\n', '')
    df = df.dropna(how='all')
    df = df.reset_index(drop=True)
    return df


def format_excel_value(raw_value):
    """格式化Excel值：整数去掉小数点"""
    if pd.isna(raw_value) or str(raw_value).strip() == '':
        return ''
    val = str(raw_value).strip()
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


def generate_merged_docx(template_path, df, variables_mapping):
    """生成多页合并Word文档"""
    with zipfile.ZipFile(template_path, 'r') as template_zip:
        doc_xml = template_zip.read('word/document.xml').decode('utf-8')

        body_match = re.search(r'(<w:body>)(.*?)(</w:body>)', doc_xml, re.DOTALL)
        if not body_match:
            body_match = re.search(r'(<[^>]*body>)(.*?)(</[^>]*body>)', doc_xml, re.DOTALL)
        if not body_match:
            raise ValueError("无法解析文档body")

        body_open = body_match.group(1)
        body_content = body_match.group(2)
        body_close = body_match.group(3)

        page_break_xml = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'

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

        output_buf = io.BytesIO()
        with zipfile.ZipFile(output_buf, 'w', zipfile.ZIP_DEFLATED) as out_zip:
            for item in template_zip.infolist():
                if item.filename == 'word/document.xml':
                    out_zip.writestr(item, new_doc_xml.encode('utf-8'))
                else:
                    out_zip.writestr(item, template_zip.read(item.filename))
        output_buf.seek(0)
        return output_buf.read()


def generate_zip_download(template_path, df, variables_mapping):
    """生成ZIP压缩包"""
    output_buf = io.BytesIO()
    with zipfile.ZipFile(output_buf, 'w', zipfile.ZIP_DEFLATED) as zf:
        for idx, row in df.iterrows():
            with zipfile.ZipFile(template_path, 'r') as template_zip:
                doc_xml = template_zip.read('word/document.xml').decode('utf-8')
                doc_xml = replace_text_in_xml(doc_xml, row, variables_mapping)

                doc_buf = io.BytesIO()
                with zipfile.ZipFile(doc_buf, 'w', zipfile.ZIP_DEFLATED) as out_zip:
                    for item in template_zip.infolist():
                        if item.filename == 'word/document.xml':
                            out_zip.writestr(item, doc_xml.encode('utf-8'))
                        else:
                            out_zip.writestr(item, template_zip.read(item.filename))
                doc_buf.seek(0)
                zf.writestr(f'箱唛_{idx + 1}.docx', doc_buf.read())
    output_buf.seek(0)
    return output_buf.read()


# ==================== Gradio 界面 ====================

def get_file_content(file_obj):
    """从Gradio文件对象获取内容"""
    if file_obj is None:
        return None
    # Gradio 返回的是文件路径
    if isinstance(file_obj, str):
        with open(file_obj, 'rb') as f:
            return f.read()
    # 或者是 namedtuple
    if hasattr(file_obj, 'name'):
        with open(file_obj.name, 'rb') as f:
            return f.read()
    # 或者直接是 bytes
    if isinstance(file_obj, bytes):
        return file_obj
    return None


def analyze_files(excel_file, template_file):
    """分析上传的文件"""
    if excel_file is None or template_file is None:
        return "❌ 请同时上传Excel和Word模板文件", gr.update(visible=False), gr.update(visible=False)

    try:
        # 获取文件内容
        excel_content = get_file_content(excel_file)
        template_content = get_file_content(template_file)

        if excel_content is None or template_content is None:
            return "❌ 无法读取文件内容", gr.update(visible=False), gr.update(visible=False)

        # 保存到临时文件
        excel_path = os.path.join(tempfile.gettempdir(), 'carton_excel.xlsx')
        template_path = os.path.join(tempfile.gettempdir(), 'carton_template.docx')

        with open(excel_path, 'wb') as f:
            f.write(excel_content)
        with open(template_path, 'wb') as f:
            f.write(template_content)

        # 分析
        variables = find_variables_in_template(template_path)
        df = pd.read_excel(excel_path, dtype=str)
        df.columns = df.columns.str.strip().str.replace('\n', '')
        excel_columns = df.columns.tolist()
        row_count = len(df.dropna(how='all'))

        info = f"""✅ 分析完成！

📋 模板变量: {variables}
📊 Excel列名: {excel_columns}
📄 数据行数: {row_count} 行

请在下方为每个变量选择对应的Excel列，然后点击生成。"""

        # 构建下拉框选项
        choices = ["-- 请选择 --"] + excel_columns

        # 尝试自动匹配
        default_vals = []
        for var in variables:
            if var in excel_columns:
                default_vals.append(var)
            else:
                default_vals.append("-- 请选择 --")

        return info, gr.update(choices=choices, value=default_vals[0] if len(default_vals) > 0 else "-- 请选择 --", visible=True), gr.update(visible=True)

    except Exception as e:
        import traceback
        traceback.print_exc()
        return f"❌ 分析失败: {str(e)}", gr.update(visible=False), gr.update(visible=False)


def generate_merged(excel_file, template_file, mapping_text):
    """生成多页合并Word文档"""
    if excel_file is None or template_file is None:
        return None, "❌ 请先上传文件"

    try:
        excel_content = get_file_content(excel_file)
        template_content = get_file_content(template_file)

        excel_path = os.path.join(tempfile.gettempdir(), 'carton_excel.xlsx')
        template_path = os.path.join(tempfile.gettempdir(), 'carton_template.docx')

        with open(excel_path, 'wb') as f:
            f.write(excel_content)
        with open(template_path, 'wb') as f:
            f.write(template_content)

        # 解析映射关系
        variables = find_variables_in_template(template_path)
        df = read_excel_data(excel_path)
        excel_columns = df.columns.tolist()

        # 自动映射：变量名与列名相同时自动匹配
        variables_mapping = {}
        for var in variables:
            if var in excel_columns:
                variables_mapping[var] = var

        if len(variables_mapping) < len(variables):
            unmapped = [v for v in variables if v not in variables_mapping]
            return None, f"❌ 以下变量在Excel中找不到对应列: {unmapped}"

        doc_bytes = generate_merged_docx(template_path, df, variables_mapping)

        # 保存到临时文件供下载
        output_path = os.path.join(tempfile.gettempdir(), '箱唛_合并.docx')
        with open(output_path, 'wb') as f:
            f.write(doc_bytes)

        return output_path, f"✅ 成功生成 {len(df)} 页箱唛！点击下方文件下载。"

    except Exception as e:
        import traceback
        traceback.print_exc()
        return None, f"❌ 生成失败: {str(e)}"


def generate_zip(excel_file, template_file, mapping_text):
    """生成ZIP压缩包"""
    if excel_file is None or template_file is None:
        return None, "❌ 请先上传文件"

    try:
        excel_content = get_file_content(excel_file)
        template_content = get_file_content(template_file)

        excel_path = os.path.join(tempfile.gettempdir(), 'carton_excel.xlsx')
        template_path = os.path.join(tempfile.gettempdir(), 'carton_template.docx')

        with open(excel_path, 'wb') as f:
            f.write(excel_content)
        with open(template_path, 'wb') as f:
            f.write(template_content)

        variables = find_variables_in_template(template_path)
        df = read_excel_data(excel_path)
        excel_columns = df.columns.tolist()

        variables_mapping = {}
        for var in variables:
            if var in excel_columns:
                variables_mapping[var] = var

        if len(variables_mapping) < len(variables):
            unmapped = [v for v in variables if v not in variables_mapping]
            return None, f"❌ 以下变量在Excel中找不到对应列: {unmapped}"

        zip_bytes = generate_zip_download(template_path, df, variables_mapping)

        output_path = os.path.join(tempfile.gettempdir(), '箱唛_全部.zip')
        with open(output_path, 'wb') as f:
            f.write(zip_bytes)

        return output_path, f"✅ 成功生成 {len(df)} 个箱唛文档！点击下方文件下载。"

    except Exception as e:
        import traceback
        traceback.print_exc()
        return None, f"❌ 生成失败: {str(e)}"


# ==================== 构建界面 ====================

css = """
.gradio-container {
    max-width: 800px !important;
    margin: auto !important;
}
"""

with gr.Blocks(title="📦 箱唛生成系统", css=css) as demo:
    gr.Markdown("""
    # 📦 箱唛生成系统
    
    **使用说明：**
    1. 上传 Excel 数据文件和 Word 模板文件
    2. 点击「分析文件」查看变量信息
    3. 点击「生成」下载箱唛文档
    
    模板变量格式：**【变量名】**，例如 【A】【B】【C】
    """)

    with gr.Row():
        excel_input = gr.File(label="📊 Excel 数据文件", file_types=[".xlsx", ".xls"])
        template_input = gr.File(label="📄 Word 模板文件", file_types=[".docx"])

    analyze_btn = gr.Button("🔍 分析文件", variant="secondary", size="lg")
    info_text = gr.Textbox(label="📋 分析结果", interactive=False, lines=6, show_copy_button=True)

    with gr.Row(visible=True) as generate_row:
        merged_btn = gr.Button("📄 生成单个Word文档（多页合并）", variant="primary", size="lg")
        zip_btn = gr.Button("📦 生成ZIP压缩包（逐个文件）", variant="secondary", size="lg")

    status_text = gr.Textbox(label="状态", interactive=False, lines=2)
    output_file = gr.File(label="⬇️ 下载生成的文件")

    # 事件绑定
    analyze_btn.click(
        fn=analyze_files,
        inputs=[excel_input, template_input],
        outputs=[info_text, generate_row, output_file]
    )

    merged_btn.click(
        fn=generate_merged,
        inputs=[excel_input, template_input, info_text],
        outputs=[output_file, status_text]
    )

    zip_btn.click(
        fn=generate_zip,
        inputs=[excel_input, template_input, info_text],
        outputs=[output_file, status_text]
    )


if __name__ == "__main__":
    demo.launch()
