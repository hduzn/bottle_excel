#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   app.py
@Time    :   2023/07/04
@Author  :   HDUZN
@Version :   1.0
@Contact :   hduzn@vip.qq.com
@License :   (C)Copyright 2023-2024
@Desc    :   1.把Excel文件按Sheet工作表保存成n个Excel文件
             2.表格按A列标题“type”拆分成n个表格,保留标题行
             3.表格按A列标题“type”拆分成n个sheet工作表,保留标题行
             4.把多个Excel文件中的数据合并到一个Excel文件的不同sheet中 merged.xlsx
             5.把多个Excel文件中的数据合并到一个Excel文件的一个sheet中 merged_one_sheet.xlsx
             pip install bottle, pandas, openpyxl, xlsxwriter
'''

# here put the import lib
from bottle import Bottle, request, static_file, template
import os, zipfile, datetime, random, string
import pandas as pd

app = Bottle()

@app.route('/')
def index():
    return template('templates/index.html')

@app.route('/css/<filename:path>')
def send_css(filename):
    return static_file(filename, root='css')

@app.route('/ex_templates/<filename:path>')
def send_xlsx(filename):
    return static_file(filename, root='ex_templates')

@app.route('/favicon.ico')
def favicon():
    return static_file('favicon.ico', root='.')

@app.route('/func')
def func():
    action = request.query.get('action', '')
    # 创建一个字典，将功能的名称和描述映射起来
    action_titles = {
        'fun1': '按Sheet拆成n个表格',
        'fun2': '按A列标题拆成n个表格',
        'fun3': '按A列标题拆分成n个sheet'
    }
    # 根据请求的参数获取对应的功能描述
    action_title = action_titles.get(action, '')
    return template('templates/func.html', action=action, action_title=action_title)

@app.route('/func2')
def func2():
    action = request.query.get('action', '')
    # 创建一个字典，将功能的名称和描述映射起来
    action_titles = {
        'fun4': '多个表格合并到不同sheet',
        'fun5': '多个表格合并到1个sheet'
    }
    # 根据请求的参数获取对应的功能描述
    action_title = action_titles.get(action, '')
    return template('templates/func2.html', action=action, action_title=action_title)

@app.route('/upload', method='POST')
def upload():
    if not os.path.exists('uploads'):
        os.makedirs('uploads')
    
    upload_file = request.files.get('file')
    if upload_file:
        # file_path = os.path.join('uploads', upload_file.filename)
        filename = generate_unique_filename()
        file_path = os.path.join('uploads', filename)
        upload_file.save(file_path)
        # print(file_path)
        
        action = request.forms.get('action')
        if action == 'fun1':
            excel_to_files_by_sheet(file_path)
        elif action == 'fun2':
            excel_split_by_type(file_path)
        elif action == 'fun3':
            excel_split_by_type_to_one(file_path)
        else:
            print('No action!')
        
        # 删除保存在uploads目录下的excel文件
        os.remove(file_path)

        return static_file('output.zip', root='output', download=True)

@app.route('/upload2', method='POST')
def upload2():
    if not os.path.exists('uploads'):
        os.makedirs('uploads')
    
    upload_files = request.files.getall('file[]')
    filenames = []
    sheet_names = []
    if upload_files:
        for upload_file in upload_files:
            temp_upload_file = upload_file.raw_filename
            # print(temp_upload_file)
            # 判断文件是否为Excel文件
            if temp_upload_file.endswith('.xlsx') or temp_upload_file.endswith('.xls'):
                # 提取文件名（不带后缀）作为Sheet名称
                ex_name = os.path.basename(temp_upload_file)
                sheet_name = ex_name.split('.')[0]
                sheet_names.append(sheet_name)

                filename = generate_unique_filename() # 生成唯一的文件名
                file_path = os.path.join('uploads', filename)
                filenames.append(file_path)
                upload_file.save(file_path)
                # print(file_path)
        
        action = request.forms.get('action')
        if action == 'fun4':
            merge_excels_into_sheets(filenames, sheet_names)
        elif action == 'fun5':
            merge_excels_into_one_sheet(filenames)
        else:
            print('No action!')
        
        # 删除保存在uploads目录下的所有excel文件
        for filename in filenames:
            os.remove(filename)
        
        return static_file('output.zip', root='output', download=True)

# 生成唯一的文件名
def generate_unique_filename():
    now = datetime.datetime.now()
    timestamp = now.strftime('%Y%m%d%H%M%S')

    letters = string.ascii_letters + string.digits
    random_digits = ''.join(random.choice(letters) for _ in range(4))
    return f'ex_{timestamp}{random_digits}.xlsx'

# 1.把Excel文件按Sheet工作表保存成n个Excel文件
def excel_to_files_by_sheet(file_path):
    excel_data = pd.read_excel(file_path, sheet_name=None)
    zip_path = os.path.join('output', 'output.zip')
    with zipfile.ZipFile(zip_path, 'w') as zip_file:
        for sheet_name, sheet_data in excel_data.items():
            sheet_file_path = os.path.join('output', f'{sheet_name}.xlsx')
            sheet_data.to_excel(sheet_file_path, index=False)
            zip_file.write(sheet_file_path, arcname=f'{sheet_name}.xlsx')
            os.remove(sheet_file_path)

# 2.表格按A列标题“type”拆分成n个表格,保留标题行
def excel_split_by_type(file_path):
    excel_data = pd.read_excel(file_path)
    column_name = 'type'
    types = excel_data[column_name].unique()
    zip_path = os.path.join('output', 'output.zip')
    with zipfile.ZipFile(zip_path, 'w') as zip_file:
        for t in types:
            filtered_data = excel_data[excel_data[column_name] == t]
            sheet_file_path = os.path.join('output', f'{t}.xlsx')
            filtered_data.to_excel(sheet_file_path, index=False)
            zip_file.write(sheet_file_path, arcname=f'{t}.xlsx')
            os.remove(sheet_file_path)

# 3.表格按A列标题“type”拆分成n个sheet工作表,保留标题行
def excel_split_by_type_to_one(file_path):
    excel_data = pd.read_excel(file_path)
    column_name = 'type'
    types = excel_data[column_name].unique()
    output_file_path = os.path.join('output', 'output.xlsx')
    with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
        for t in types:
            filtered_data = excel_data[excel_data[column_name] == t]
            filtered_data.to_excel(writer, sheet_name=str(t), index=False)  # 将t转换为字符串
    zip_output(output_file_path, 'output.xlsx')

# 4.把多个Excel文件中的数据合并到一个Excel文件的不同sheet中 merged.xlsx
def merge_excels_into_sheets(file_paths, sheet_names):
    output_file_path = os.path.join('output', 'merged.xlsx')
    with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as writer:
        for file_path in file_paths:
                sheet_name = sheet_names[file_paths.index(file_path)]
                df = pd.read_excel(file_path)
                df.to_excel(writer, sheet_name=sheet_name, index=False)
    zip_output(output_file_path, 'merged.xlsx')

# 5.把多个Excel文件中的数据合并到一个Excel文件的一个sheet中
def merge_excels_into_one_sheet(file_paths):
    output_file_path = os.path.join('output', 'merged_one_sheet.xlsx')
    merged_data = pd.DataFrame() # 创建一个空的DataFrame用于存储合并后的数据
    for file_path in file_paths:
        df = pd.read_excel(file_path)
        merged_data = pd.concat([merged_data, df], ignore_index=True, sort=False)
    merged_data.to_excel(output_file_path, index=False)
    zip_output(output_file_path, 'merged_one_sheet.xlsx')

# 把单个文件打包成 output.zip
def zip_output(output_file_path, arcname):
    if os.path.exists(output_file_path):
        zip_path = os.path.join('output', 'output.zip')
        with zipfile.ZipFile(zip_path, 'w') as zip_file:
            zip_file.write(output_file_path, arcname=arcname)
        os.remove(output_file_path)
    else:
        print(f"File {output_file_path} does not exist.")

if __name__ == '__main__':
    app.run(host='0.0.0.0', port='8080')