import streamlit as st
import pdfplumber
import re
import pandas as pd
from tabulate import tabulate
import io

# --------------------------
# 定义订单解析相关函数
# --------------------------

def extract_field(pattern, text):
    match = re.search(pattern, text)
    return match.group(1) if match else None

def process_tables(tables):
    # 展开表格生成字符串列表
    flattened_list = [item.replace('\n', ' ') for sublist1 in tables for sublist2 in sublist1 for item in sublist2 if item]
    result = ", ".join(flattened_list)
    return result

def text_extraction(page):
    text = page.extract_text()
    if text:
        customer_order = extract_field(r'Customer Order #:\s*(\S+)', text)
        purchase_order = extract_field(r'Purchase Order #:\s*(\S+)', text)
        date = extract_field(r'Date:\s*(\S+)', text)
        address_type = extract_field(r'Address Type:\s*(\w+)', text)

        fields = {
            "Customer_Order": customer_order,
            "PO#": purchase_order,
            "PO_Date": date,
            "Address_Type": address_type
        }
        return fields
    return {}

def split_text(text):
    # 根据 'Model Number' 分割文本为两个部分，前半部分作为订单信息，后半部分作为shipping信息
    parts = text.split('Model Number')
    order = parts[0]
    ship = 'Model_Number' + parts[1] if len(parts) > 1 else ''
    return order, ship

def extract_order_details(text):
    pattern = re.compile(
        r'Ordered By:,\s*(?P<Customer_Name>[^,]+(?:,\s*[^,]+)?)\s*,\s*' 
        r'Ship To:,\s*(?P<Street_Address>.+?)\s+'   
        r'(?P<State>\w{2})\s+'
        r'(?P<Zipcode>\d{5}(?:-\d{4})?)\s+'
        r'(?P<Tel>\d{3}-?\d{3}-?\d{4})'
    )
    match = pattern.search(text)
    if match:
        customer_name = match.group('Customer_Name').strip()
        street_address = match.group('Street_Address').strip()
        # 如果 street_address 中包含 Customer_Name 则去除
        if street_address.startswith(customer_name):
            street_address = street_address[len(customer_name):].strip()
        results = {
            'Customer_Name': customer_name,
            'Street_Address': street_address,
            'State': match.group('State'),
            'Zipcode': match.group('Zipcode'),
            'Tel#': match.group('Tel')
        }
    else:
        results = {
            'Customer_Name': None,
            'Street_Address': None,
            'State': None,
            'Zipcode': None,
            'Tel#': None
        }
    return results

def extract_ship_details(text):
    # 读取并重组shipping表格内容
    data = text.split(',')
    # 去掉空字符串或仅包含空格的元素
    data = [element for element in data if element.strip()]
    columns = ['Model_Number', 'Internet_Number', 'Item_Description', 'Qty_Shipped']
    records = []
    i = 4  # 假设前4个字段为表头，数据从第5个开始
    while i < len(data):
        if data[i].strip().startswith('Message:'):
            break  # 遇到 'Message:' 时停止解析
        record = {}
        record['Model_Number'] = data[i].strip()
        record['Internet_Number'] = data[i + 1].strip() if i + 1 < len(data) else ''
        # Item_Description 可能跨越多个片段
        description = []
        j = i + 2
        while j < len(data) and not data[j].strip().isdigit():
            description.append(data[j].strip())
            j += 1
        record['Item_Description'] = ' '.join(description)
        # Qty_Shipped 为数字字段
        if j < len(data):
            try:
                record['Qty_Shipped'] = int(data[j].strip())
            except:
                record['Qty_Shipped'] = None
        records.append(record)
        # 跳到下一个记录的起点
        i = j + 1
    df = pd.DataFrame(records, columns=columns)
    return df

def table_extraction(page):
    tables = page.extract_tables()
    text = process_tables(tables)
    order_text, ship_text = split_text(text)
    order_results = extract_order_details(order_text)
    ship_results = extract_ship_details(ship_text)
    # 如果 shipping 表格中有多条记录，这里只取第一条（也可以进行合并或其他处理）
    if not ship_results.empty:
        ship_row = ship_results.iloc[0].to_dict()
    else:
        ship_row = {}
    combined_results = {**order_results, **ship_row}
    return combined_results

def to_df(pdf_file):
    df = pd.DataFrame()
    with pdfplumber.open(pdf_file) as pdf:
        for i, page in enumerate(pdf.pages):
            try:
                text_results = text_extraction(page)
                table_results = table_extraction(page)
                merged_results = {**text_results, **table_results}
                df1 = pd.DataFrame([merged_results])
                df = pd.concat([df, df1], axis=0).reset_index(drop=True)
            except Exception as e:
                st.error(f"在第 {i+1} 页发生错误：{e}")
                continue
    return df

# --------------------------
# Streamlit 应用主程序
# --------------------------

def main():
    st.title("异常订单检测系统")
    st.write("上传订单 PDF 文件，系统将解析并检测异常订单。")
    
    uploaded_file = st.file_uploader("请选择 PDF 文件", type="pdf")
    if uploaded_file is not None:
        # 将上传的文件转为 BytesIO 对象供 pdfplumber 读取
        pdf_bytes = uploaded_file.read()
        pdf_file = io.BytesIO(pdf_bytes)
        
        with st.spinner("正在解析 PDF 文件，请稍候..."):
            df = to_df(pdf_file)
        
        st.success("解析完成！")
        st.subheader("全部订单数据")
        st.dataframe(df)
        
        # 异常订单筛选条件：Street_Address 不包含 "C/O" 且 Qty_Shipped >= 10
        try:
            # 将 Qty_Shipped 转换为数值型
            df['Qty_Shipped'] = pd.to_numeric(df['Qty_Shipped'], errors='coerce')
            abnormal_df = df[(~df['Street_Address'].fillna('').str.contains('C/O')) & (df['Qty_Shipped'] >= 10)]
            
            st.subheader("异常订单数据")
            if abnormal_df.empty:
                st.info("未能找到任何符合条件的异常订单")
            else:
                st.dataframe(abnormal_df)
                txt_table = tabulate(abnormal_df, headers='keys', tablefmt='grid', showindex=False)
                st.download_button(
                    label="下载异常订单数据",
                    data=txt_table,
                    file_name="abnormal_orders.txt",
                    mime="text/plain"
                )
        except Exception as e:
            st.error(f"筛选异常订单时发生错误：{e}")

if __name__ == '__main__':
    main()
