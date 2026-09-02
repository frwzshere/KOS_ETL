import streamlit as st
import io
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import re
from datetime import datetime

def process_excel(file_bytes):
    wb = load_workbook(filename=io.BytesIO(file_bytes), data_only=False)

    # ----- 1. 构建映射字典（从“名单-昵称”表） -----
    if "名单-昵称" not in wb.sheetnames:
        raise ValueError("工作簿中缺少“名单-昵称”工作表")
    name_sheet = wb["名单-昵称"]
    headers = [cell.value for cell in name_sheet[1]]
    try:
        col_nick = headers.index("抖音昵称") + 1
        col_emp = headers.index("员工姓名") + 1
    except ValueError:
        raise ValueError("“名单-昵称”表中缺少“抖音昵称”或“员工姓名”列")
    
    name_map = {}
    for row in range(2, name_sheet.max_row + 1):
        nick = name_sheet.cell(row, col_nick).value
        emp = name_sheet.cell(row, col_emp).value
        if nick and emp:
            name_map[str(nick)] = str(emp)

    red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

    # ===== 辅助函数：清理客户来源（去掉末尾括号） =====
    def clean_source(val):
        if val is None:
            return ""
        s = str(val)
        return re.sub(r'[（(][^）)]*[）)]$', '', s).strip()

    # ===== 辅助函数：格式化日期 =====
    def format_date_cell(cell):
        val = cell.value
        if isinstance(val, datetime):
            cell.value = val.strftime("%Y/%m/%d")
        elif isinstance(val, str):
            # 尝试解析常见格式（可扩展）
            for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%d-%m-%Y", "%d/%m/%Y"):
                try:
                    dt = datetime.strptime(val, fmt)
                    cell.value = dt.strftime("%Y/%m/%d")
                    break
                except ValueError:
                    continue

    # ===== 2. 处理“DOU+”表 =====
    if "DOU+" in wb.sheetnames:
        dou_sheet = wb["DOU+"]
        headers_dou = [cell.value for cell in dou_sheet[1]]
        # 找到各列索引
        try:
            col_date = headers_dou.index("下单时间") + 1
        except ValueError:
            raise ValueError("“DOU+”表中缺少“下单时间”列")
        # 假设用于映射的列是“客户来源”（请根据实际列名修改）
        try:
            col_source_dou = headers_dou.index("客户来源") + 1
        except ValueError:
            raise ValueError("“DOU+”表中缺少用于映射的“客户来源”列")
        try:
            col_person_dou = headers_dou.index("人员") + 1
        except ValueError:
            raise ValueError("“DOU+”表中缺少“人员”列")
        
        for row in range(2, dou_sheet.max_row + 1):
            # 日期格式化
            format_date_cell(dou_sheet.cell(row, col_date))
            
            # 人员映射
            source_cell = dou_sheet.cell(row, col_source_dou)
            person_cell = dou_sheet.cell(row, col_person_dou)
            cleaned = clean_source(source_cell.value)
            matched = name_map.get(cleaned)
            if matched:
                person_cell.value = matched
            else:
                person_cell.fill = red_fill  # 无法匹配标红

    # ===== 3. 处理“留资”表 =====
    if "留资" not in wb.sheetnames:
        raise ValueError("工作簿中缺少“留资”工作表")
    liu_sheet = wb["留资"]
    headers_liu = [cell.value for cell in liu_sheet[1]]
    try:
        col_date_liu = headers_liu.index("日期") + 1
    except ValueError:
        raise ValueError("“留资”表中缺少“日期”列")
    try:
        col_source_liu = headers_liu.index("客户来源") + 1
    except ValueError:
        raise ValueError("“留资”表中缺少“客户来源”列")
    try:
        col_person_liu = headers_liu.index("人员") + 1
    except ValueError:
        raise ValueError("“留资”表中缺少“人员”列")
    
    for row in range(2, liu_sheet.max_row + 1):
        # 日期格式化
        format_date_cell(liu_sheet.cell(row, col_date_liu))
        # 人员映射
        source_cell = liu_sheet.cell(row, col_source_liu)
        person_cell = liu_sheet.cell(row, col_person_liu)
        cleaned = clean_source(source_cell.value)
        matched = name_map.get(cleaned)
        if matched:
            person_cell.value = matched
        else:
            person_cell.fill = red_fill

    # ===== 4. 保存并返回 =====
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()

# ----- Streamlit 界面（保持不变） -----
st.set_page_config(page_title="Excel 数据转换工具（保留格式）", layout="centered")
st.title("📁 Excel 数据转换工具")
st.markdown("上传你的 Excel 文件，自动处理 DOU+、留资、名单-昵称 三个工作表，并保留所有格式。")

uploaded_file = st.file_uploader("选择 Excel 文件", type=["xlsx", "xlsm"])

if uploaded_file is not None:
    st.info(f"已上传：{uploaded_file.name} (大小：{uploaded_file.size} 字节)")
    if st.button("🚀 开始处理", type="primary"):
        with st.spinner("正在处理，请稍候..."):
            try:
                file_bytes = uploaded_file.read()
                processed_data = process_excel(file_bytes)
                st.success("✅ 处理完成！")
                st.download_button(
                    label="📥 下载处理后的 Excel 文件",
                    data=processed_data,
                    file_name=f"processed_{uploaded_file.name}",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"❌ 处理出错：{e}")
