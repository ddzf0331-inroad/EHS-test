import streamlit as st
import pandas as pd
import io
import csv
import zipfile
import os

# ================= 配置区域 =================
TARGET_MONITOR_POINT = "ABS装置焚烧炉废气排放口"

# ================= 辅助函数：保留3位小数 =================
def format_decimal(val):
    if val is None: return ''
    val_str = str(val).strip()
    if not val_str: return ''
    try:
        f_val = float(val_str)
        return "{:.3f}".format(f_val)
    except ValueError:
        return val_str

# ================= 核心逻辑：读取上传的文件流 =================
def load_file(uploaded_file):
    """
    读取 Streamlit 上传的文件对象，返回字典 {sheet_name: data_list}
    """
    filename = uploaded_file.name
    ext = os.path.splitext(filename)[1].lower()
    result = {}

    # 1. Excel 文件处理
    if ext in ['.xlsx', '.xls']:
        try:
            # Streamlit 的 uploaded_file 直接就是二进制流，可以直接喂给 pandas
            # 必须指定 engine，且 openpyxl/xlrd 需要已安装
            engine = 'xlrd' if ext == '.xls' else 'openpyxl'
            xls = pd.ExcelFile(uploaded_file, engine=engine)
            
            for sheet_name in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
                result[sheet_name] = df.fillna('').values.tolist()
            return result
        except Exception as e:
            st.error(f"Excel 读取失败: {e}")
            return None

    # 2. CSV 文件处理 (需要处理编码)
    else:
        # 读取二进制内容
        bytes_data = uploaded_file.getvalue()
        encodings = ['utf-8', 'gbk', 'gb18030', 'utf-8-sig']
        
        for enc in encodings:
            try:
                # 解码为字符串
                string_data = bytes_data.decode(enc)
                # 使用 csv 模块读取字符串流
                f_io = io.StringIO(string_data)
                reader = csv.reader(f_io)
                data = list(reader)
                return {'CSV_Content': data}
            except:
                continue
        
        st.error("无法识别该文件的编码 (CSV)，请确保文件未损坏。")
        return None

# ================= 数据处理主逻辑 =================
def process_data(source_file, template_file):
    # 1. 读取文件
    source_dict = load_file(source_file)
    template_dict = load_file(template_file)

    if not source_dict or not template_dict:
        return None

    # 2. 解析数据源 (找到包含数据的Sheet)
    source_rows = []
    for s_name, rows in source_dict.items():
        if len(rows) > 5:
            source_rows = rows
            break
    
    # 寻找日期行
    data_start_idx = -1
    for i, row in enumerate(source_rows):
        if len(row) > 0:
            s = str(row[0]).strip()
            if s.startswith('20') and ('-' in s or '/' in s):
                data_start_idx = i
                break
    
    if data_start_idx == -1:
        st.error("数据源中未找到日期行 (例如 2025-08-01...)")
        return None

    # 3. 提取数据
    source_map = {}
    valid_count = 0
    for row in source_rows[data_start_idx:]:
        if len(row) < 10: continue
        try:
            time_str = str(row[0]).strip()
            date_part = time_str[:10]
            hour = 0
            if ' ' in time_str:
                try: hour = int(time_str.split(' ')[1].split(':')[0])
                except: pass
            
            if date_part not in source_map: source_map[date_part] = {}
            
            def g(idx): 
                val = row[idx] if idx < len(row) else ''
                return format_decimal(val)

            # A=0, B=1(流量), E=4(NOx), J=9(非甲烷), O=14(O2), R=17(流速), U=20(温度), X=23(湿度)
            # 修改需求：NOx取排放量(索引6)，NMHC取排放量(索引11)
            source_map[date_part][hour] = {
                'flow': g(1), 
                'nox':  g(6),   # 排放量
                'nmhc': g(11),  # 排放量
                'o2':   g(14), 
                'velo': g(17), 
                'temp': g(20), 
                'humi': g(23)
            }
            valid_count += 1
        except: continue
    
    st.info(f"成功解析数据源：共 {valid_count} 条有效数据，涵盖 {len(source_map)} 天。")

    # 4. 定位模板
    target_sheet_name = None
    target_template_rows = []
    target_row_idx = -1
    clean_target = TARGET_MONITOR_POINT.replace(" ", "").strip()

    for sheet_name, rows in template_dict.items():
        for i, row in enumerate(rows):
            row_str = "".join([str(x) for x in row]).replace(" ", "").replace("　", "").replace("\t", "")
            if clean_target in row_str:
                target_row_idx = i
                target_template_rows = rows
                target_sheet_name = sheet_name
                break
        if target_row_idx != -1: break

    if target_row_idx == -1:
        st.error(f"在模板的所有 Sheet 中都找不到关键词：'{clean_target}'")
        return None
    
    st.success(f"模板匹配成功！使用 Sheet: '{target_sheet_name}'")

    # 5. 生成结果 (写入内存中的 ZIP)
    zip_buffer = io.BytesIO()
    fill_start = target_row_idx + 3

    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
        for date_key, day_data in source_map.items():
            # 复制模板数据
            out_rows = [r[:] for r in target_template_rows]
            
            # 修改日期
            if len(out_rows[target_row_idx]) > 0:
                out_rows[target_row_idx][0] = date_key
            
            # 填充数据
            for h in range(24):
                r = fill_start + h
                if r >= len(out_rows): break
                while len(out_rows[r]) < 10: out_rows[r].append('')
                
                d = day_data.get(h, {})
                if d:
                    out_rows[r][1] = d['flow']
                    out_rows[r][2] = d['temp']
                    out_rows[r][3] = d['humi']
                    out_rows[r][4] = d['o2']
                    out_rows[r][5] = d['velo']
                    out_rows[r][7] = d['nmhc']
                    out_rows[r][8] = d['nox']
            
            # 将生成的 CSV 转换为字符串流，再写入 ZIP
            # 使用 utf-8-sig 以便 Excel 打开不乱码
            csv_buffer = io.StringIO()
            writer = csv.writer(csv_buffer)
            writer.writerows(out_rows)
            
            # 写入 zip (文件名, 文件内容)
            zf.writestr(f"{date_key}_日报表.csv", csv_buffer.getvalue().encode('utf-8-sig'))
    
    # 指针归位
    zip_buffer.seek(0)
    return zip_buffer

# ================= 网页界面布局 =================
st.set_page_config(page_title="EHS 日报表生成工具", layout="centered")

st.title("🏭 EHS 环保日报表自动生成工具")
st.markdown("---")

st.markdown("### 1. 上传文件")
col1, col2 = st.columns(2)

with col1:
    source_file = st.file_uploader("上传 [数据源] 文件 (.xlsx)", type=['xlsx', 'xls', 'csv'])

with col2:
    template_file = st.file_uploader("上传 [模板] 文件 (.xls)", type=['xls', 'xlsx', 'csv'])

st.markdown("---")

# 按钮状态逻辑
if source_file and template_file:
    if st.button("🚀 开始处理数据", type="primary"):
        with st.spinner("正在疯狂计算中，请稍候..."):
            # 调用处理函数
            zip_result = process_data(source_file, template_file)
            
            if zip_result:
                st.balloons() # 撒花特效
                st.success("处理完成！请点击下方按钮下载结果。")
                
                # 下载按钮
                st.download_button(
                    label="📥 下载生成的结果 (ZIP压缩包)",
                    data=zip_result,
                    file_name="生成的日报表.zip",
                    mime="application/zip"
                )
else:
    st.info("请先上传两个文件，然后点击运行按钮。")