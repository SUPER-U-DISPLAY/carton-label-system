"""
箱唛生成系统 - Gradio 版本（适配 Hugging Face Spaces）
功能：上传Excel和Word模板，自动生成包含多页箱唛的Word文档
核心：在ZIP层面操作docx，完整保留原始字体、图片、样式
"""

import os
import re
import tempfile
import pandas as pd
import io
import zipfile
import gradio as gr

# ==================== 核心逻辑（与Flask版完全一致） ====================

def find_variables_in_template(doc_path):
    with zipfile.ZipFile(doc_path, 'r') as z:
        doc_xml = z.read('word/document.xml').decode('utf-8')
    variables = set(re.findall(r'【([^】]+)】', doc_xml))
    return sorted(list(variables))


def read_excel_data(excel_path):
    df = pd.read_excel(excel_path, dtype=str)
    df.columns = df.columns.str.strip().str.replace('\n', '')
    df = df.dropna(how='all')
    df = df.reset_index(drop=True)
    return df


def format_excel_value(raw_value):
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
    for var_name, excel_col in variables_mapping.items():
        placeholder = f'【{var_name}】'
        if placeholder in xml_str:
            if excel_col in row_data.index:
                value = format_excel_value(row_data[excel_col])
                xml_str = xml_str.replace(placeholder, value)
    return xml_str


def generate_merged_docx(template_path, df, variables_mapping):
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

def analyze_files(excel_file, template_file):
    """分析上传的文件，返回变量信息和列名"""
    if excel_file is None or template_file is None:
        return "❌ 请同时上传Excel和Word模板文件", None, None, gr.update(visible=False), gr.update(visible=False)

    try:
        # 保存到临时文件
        excel_path = os.path.join(tempfile.gettempdir(), 'carton_excel.xlsx')
        template_path = os.path.join(tempfile.gettempdir(), 'carton_template.docx')

        with open(excel_path, 'wb') as f:
            f.write(excel_file)
        with open(template_path, 'wb') as f:
            f.write(template_file)

        variables = find_variables_in_template(template_path)
        df = pd.read_excel(excel_path, dtype=str)
        df.columns = df.columns.str.strip().str.replace('\n', '')
        excel_columns = df.columns.tolist()
        row_count = len(df.dropna(how='all'))

        info = f"✅ 分析完成！\n\n📋 模板变量: {variables}\n📊 Excel列名: {excel_columns}\n📄 数据行数: {row_count} 行\n\n请在下方为每个变量选择对应的Excel列，然后点击生成。"

        # 构建变量映射的默认选择
        default_choices = []
        for var in variables:
            if var in excel_columns:
                default_choices.append(var)
            else:
                default_choices.append(None)

        # 动态创建下拉框选项
        dropdown_choices = gr.update(choices=["-- 请选择 --"] + excel_columns, value=default_choices[0] if default_choices else None)
        dropdown2_choices = gr.update(choices=["-- 请选择 --"] + excel_columns, value=default_choices[1] if len(default_choices) > 1 else None)
        dropdown3_choices = gr.update(choices=["-- 请选择 --"] + excel_columns, value=default_choices[2] if len(default_choices) > 2 else None)

        return info, dropdown_choices, dropdown2_choices, dropdown3_choices, gr.update(visible=True), gr.update(visible=True)

    except Exception as e:
        return f"❌ 分析失败: {str(e)}", None, None, gr.update(visible=False), gr.update(visible=False)


def generate_merged(excel_file, template_file, col_a, col_b, col_c):
    """生成多页合并Word文档"""
    if excel_file is None or template_file is None:
        return None, "❌ 请先上传文件"

    try:
        excel_path = os.path.join(tempfile.gettempdir(), 'carton_excel.xlsx')
        template_path = os.path.join(tempfile.gettempdir(), 'carton_template.docx')

        with open(excel_path, 'wb') as f:
            f.write(excel_file)
        with open(template_path, 'wb') as f:
            f.write(template_file)

        variables = find_variables_in_template(template_path)
        variables_mapping = {}

        # 根据变量数量构建映射
        var_cols = [col_a, col_b, col_c]
        for i, var in enumerate(variables):
            if i < len(var_cols) and var_cols[i] and var_cols[i] != "-- 请选择 --":
                variables_mapping[var] = var_cols[i]

        if len(variables_mapping) < len(variables):
            unmapped = [v for v in variables if v not in variables_mapping]
            return None, f"❌ 请为以下变量选择Excel列: {unmapped}"

        df = read_excel_data(excel_path)
        doc_bytes = generate_merged_docx(template_path, df, variables_mapping)

        return doc_bytes, f"✅ 成功生成 {len(df)} 页箱唛！点击下方文件下载。"

    except Exception as e:
        import traceback
        traceback.print_exc()
        return None, f"❌ 生成失败: {str(e)}"


def generate_zip(excel_file, template_file, col_a, col_b, col_c):
    """生成ZIP压缩包"""
    if excel_file is None or template_file is None:
        return None, "❌ 请先上传文件"

    try:
        excel_path = os.path.join(tempfile.gettempdir(), 'carton_excel.xlsx')
        template_path = os.path.join(tempfile.gettempdir(), 'carton_template.docx')

        with open(excel_path, 'wb') as f:
            f.write(excel_file)
        with open(template_path, 'wb') as f:
            f.write(template_file)

        variables = find_variables_in_template(template_path)
        variables_mapping = {}

        var_cols = [col_a, col_b, col_c]
        for i, var in enumerate(variables):
            if i < len(var_cols) and var_cols[i] and var_cols[i] != "-- 请选择 --":
                variables_mapping[var] = var_cols[i]

        if len(variables_mapping) < len(variables):
            unmapped = [v for v in variables if v not in variables_mapping]
            return None, f"❌ 请为以下变量选择Excel列: {unmapped}"

        df = read_excel_data(excel_path)
        zip_bytes = generate_zip_download(template_path, df, variables_mapping)

        return zip_bytes, f"✅ 成功生成 {len(df)} 个箱唛文档！点击下方文件下载。"

    except Exception as e:
        import traceback
        traceback.print_exc()
        return None, f"❌ 生成失败: {str(e)}"


# ==================== 构建界面 ====================

with gr.Blocks(title="📦 箱唛生成系统") as demo:
    gr.Markdown("# 📦 箱唛生成系统")
    gr.Markdown("上传Excel数据和Word模板，自动生成批量箱唛文档。模板中的变量格式为【变量名】。")

    with gr.Row():
        with gr.Column():
            excel_input = gr.File(label="📊 Excel 数据文件", file_types=[".xlsx", ".xls"])
            template_input = gr.File(label="📄 Word 模板文件", file_types=[".docx"])
            analyze_btn = gr.Button("🔍 分析文件", variant="secondary")

    info_text = gr.Textbox(label="📋 分析结果", interactive=False, lines=6)

    with gr.Row(visible=False) as mapping_row:
        with gr.Column():
            gr.Markdown("### 🔗 变量映射")
            col_a = gr.Dropdown(label="变量 A → Excel列", choices=[], visible=False)
            col_b = gr.Dropdown(label="变量 B → Excel列", choices=[], visible=False)
            col_c = gr.Dropdown(label="变量 C → Excel列", choices=[], visible=False)

    with gr.Row(visible=False) as generate_row:
        with gr.Column():
            merged_btn = gr.Button("📄 生成单个Word文档（多页合并）", variant="primary")
            zip_btn = gr.Button("📦 生成ZIP压缩包（逐个文件）", variant="secondary")

    status_text = gr.Textbox(label="状态", interactive=False, lines=2)
    output_file = gr.File(label="⬇️ 下载生成的文件")

    # 事件绑定
    analyze_btn.click(
        fn=analyze_files,
        inputs=[excel_input, template_input],
        outputs=[info_text, col_a, col_b, col_c, mapping_row, generate_row]
    )

    merged_btn.click(
        fn=generate_merged,
        inputs=[excel_input, template_input, col_a, col_b, col_c],
        outputs=[output_file, status_text]
    )

    zip_btn.click(
        fn=generate_zip,
        inputs=[excel_input, template_input, col_a, col_b, col_c],
        outputs=[output_file, status_text]
    )


if __name__ == "__main__":
    demo.launch()
