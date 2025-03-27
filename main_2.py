from pdfminer.high_level import extract_pages, extract_text
import pdfplumber
from pdf2image import convert_from_path
import re
import pandas as pd
from tabulate import tabulate

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

def split_text(text):
    # 根据Model Number所在的位置，拆分Order表格和Shipping表格
    parts = text.split('Model Number')
    order, ship = parts[0], parts[1]
    ship = 'Model_Number' + ship
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

        # 去除掉Street_Address中的Customer_Name的部分
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
    data = [element for element in data if element != ' ']
    columns = ['Model_Number', 'Internet_Number', 'Item_Description', 'Qty_Shipped']
    records = []
    i = 4
    while i < len(data):

        if data[i].strip().startswith('Message:'):
            break  # 遇到 'Message:' 时停止解析

        record = {}
        record['Model_Number'] = data[i].strip()
        record['Internet_Number'] = data[i + 1].strip()
        
        # Item_Description 可能跨越多个片段
        description = []
        j = i + 2
        while j < len(data) and not data[j].strip().isdigit():
            description.append(data[j].strip())
            j += 1
        
        record['Item_Description'] = ' '.join(description)
        
        # Qty_Shipped 是第一个数字字段
        if j < len(data):
            record['Qty_Shipped'] = int(data[j].strip())
        
        records.append(record)
        
        # 跳转到下一个记录的起点
        i = j + 1
        
        df = pd.DataFrame(records, columns=columns)

    return df

def table_extraction(page):
    tables = page.extract_tables()
    text = process_tables(tables)
    order, ship = split_text(text)
    order_results = extract_order_details(order)
    ship_results = extract_ship_details(ship)
    # ship_results 是一个 DataFrame，可能有多行
    for _, ship_row in ship_results.iterrows():
        # 将 order_results 和 ship_row 合并成一个完整的字典 
        combined_results = {**order_results, **ship_row.to_dict()}
    return combined_results

def to_df(pdf_path):
    df = pd.DataFrame()
    pdf = pdfplumber.open(pdf_path)
    for i, page in enumerate(pdf.pages):
        try:
            text_results = text_extraction(page)
            table_results = table_extraction(page)
            merged_results = {**text_results, **table_results}
            df1 = pd.DataFrame([merged_results])
            df = pd.concat([df, df1], axis=0).reset_index(drop=True)
        except Exception as e:
            print("Encounter error {e} at page {i}")
            continue
    return df

def main():
    
    df = to_df(pdf_path)
    print(df)
    try:
        df1 = df[(~df['Street_Address'].str.contains('C/O')) & (df['Qty_Shipped']>=10)]
        columns = ['Qty_Shipped', 'Model_Number', 'Customer_Order', 'PO#', 'PO_Date', 'Customer_Name', 'Tel#', 'Street_Address', 'State', 'Zipcode', 'Address_Type']
        df1 = df1[columns]
        if df1.empty:
            print("未能找到任何符合条件的订单")
        else:
            txt_file_path = "large_purchase.txt"
            with open(txt_file_path, 'a') as f:
                f.write(tabulate(df1, headers='keys', tablefmt='grid', showindex=False))
                f.write('\n')  # Ensure there's a newline at the end

            print(f"数据已成功写入文本文件 {txt_file_path}")
            print(df1)
    except:
        none_rows = df[df['Customer_Name'].isna()]
        print("请检查该订单的格式：", none_rows)

if __name__ == '__main__':
    pdf_path = "2025.03.24 Homedepot-Carro Orders-美西-slips.pdf"
    main()