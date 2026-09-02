import streamlit as st
import io
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
import re
from datetime import datetime

# 定义处理函数（接收字节流，返回字节流）
def process_excel(file_bytes):
    # 加载工作簿（保留所有格式）
    wb = load_workbook(filename=io.BytesIO(file_bytes), data_only=False)

    # ----- 1. 构建映射字典（从“名单-昵称”表） -----
    if "名单-昵称" not in wb.sheetnames:
        raise ValueError("工作簿中缺少“名单-昵称”工作表")
    name_sheet = wb["名单-昵称"]
    # 找到“抖音昵称”和“员工姓名”的列索引（假设第一行为标题）
    headers = [cell.value for cell in name_sheet[1]]  # 第1行是标题
    try:
        col_nick = headers.index("抖音昵称") + 1
        col_emp = headers.index("员工姓名") + 1
    except ValueError:
        raise ValueError("“名单-昵称”表中缺少“抖音昵称”或“员工姓名”列")
    
    name_map = {}
    for row in range(2, name_sheet.max_row + 1):
        nick = name_sheet.cell(row, col_nick).value
        emp = name_sheet.cell(row, col_emp).value
        if nick and emp:  # 非空
            name_map[str(nick)] = str(emp)

    # ----- 2. 处理“DOU+”表：格式化日期 -----
    if "DOU+" in wb.sheetnames:
        dou_sheet = wb["DOU+"]
        # 找到“下单时间”列
        headers_dou = [cell.value for cell in dou_sheet[1]]
        try:
            col_date = headers_dou.index("下单时间") + 1
        except ValueError:
            raise ValueError("“DOU+”表中缺少“下单时间”列")
        for row in range(2, dou_sheet.max_row + 1):
            cell = dou_sheet.cell(row, col_date)
            val = cell.value
            if isinstance(val, datetime):
                cell.value = val.strftime("%Y/%m/%d")
            elif isinstance(val, str):
                # 尝试解析常见日期格式，若成功则格式化
                try:
                    dt = datetime.strptime(val, "%Y-%m-%d")  # 可扩展其他格式
                    cell.value = dt.strftime("%Y/%m/%d")
                except ValueError:
                    pass  # 保持原样

    # ----- 3. 处理“留资”表：清理客户来源 + 映射人员 + 标红 -----
    if "留资" not in wb.sheetnames:
        raise ValueError("工作簿中缺少“留资”工作表")
    liu_sheet = wb["留资"]
    headers_liu = [cell.value for cell in liu_sheet[1]]
    try:
        col_source = headers_liu.index("客户来源") + 1
        col_person = headers_liu.index("人员") + 1
    except ValueError:
        raise ValueError("“留资”表中缺少“客户来源”或“人员”列")

    # 定义红色填充
    red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

    for row in range(2, liu_sheet.max_row + 1):
        source_cell = liu_sheet.cell(row, col_source)
        person_cell = liu_sheet.cell(row, col_person)

        # 清理客户来源：去掉末尾括号及内容
        raw = source_cell.value
        if raw is not None:
            raw_str = str(raw)
            cleaned = re.sub(r'[（(][^）)]*[）)]$', '', raw_str).strip()
        else:
            cleaned = ""

        # 在映射中查找
        matched_name = name_map.get(cleaned)
        if matched_name is not None:
            # 更新人员列
            person_cell.value = matched_name
        else:
            # 无法匹配 → 标红背景（但不修改内容）
            person_cell.fill = red_fill

    # ----- 4. 将修改后的工作簿写入内存 -----
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()

# ----- Streamlit 界面 -----
st.set_page_config(page_title="Excel 数据转换工具（保留格式）", layout="centered")
st.title("📁 Excel 数据转换工具")
st.markdown("上传你的 Excel 文件，自动处理 DOU+、留资、名单-昵称 三个工作表，并保留所有格式。")

uploaded_file = st.file_uploader("选择 Excel 文件", type=["xlsx", "xlsm"])

if uploaded_file is not None:
    # 显示文件信息
    st.info(f"已上传：{uploaded_file.name} (大小：{uploaded_file.size} 字节)")

    # 处理按钮（或自动处理）
    if st.button("🚀 开始处理", type="primary"):
        with st.spinner("正在处理，请稍候..."):
            try:
                # 读取上传文件的字节
                file_bytes = uploaded_file.read()
                processed_data = process_excel(file_bytes)
                # 提供下载按钮
                st.success("✅ 处理完成！")
                st.download_button(
                    label="📥 下载处理后的 Excel 文件",
                    data=processed_data,
                    file_name=f"processed_{uploaded_file.name}",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"❌ 处理出错：{e}")