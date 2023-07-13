#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   app.py
@Time    :   2023/07/04
@Author  :   HDUZN
@Version :   1.1
@Contact :   hduzn@vip.qq.com
@License :   (C)Copyright 2023-2024
@Desc    :   1.把Excel文件按Sheet工作表保存成n个Excel文件
             2.表格按A列标题“type”拆分成n个表格,保留标题行
             3.表格按A列标题“type”拆分成n个sheet工作表,保留标题行
             4.把多个Excel文件中的数据合并到一个Excel文件的不同sheet中
             5.把多个Excel文件中的数据合并到一个Excel文件的一个sheet中
             pip install bottle, pandas, openpyxl, xlsxwriter
'''

# here put the import lib
from bottle import Bottle, request, static_file, template
import os, zipfile, datetime, random, string, shutil
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
        delete_files_in_dir(f'output') # 清空output目录
        user_id = generate_unique_name() # 创建唯一的用户id
        user_dir = f'uploads/{user_id}'
        os.makedirs(user_dir) # 创建用户目录

        file_path = f'uploads/{user_id}/{upload_file.raw_filename}'
        upload_file.save(file_path)
        # print(file_path)
        
        action = request.forms.get('action')
        if action == 'fun1':
            excel_to_files_by_sheet(file_path, user_id)
        elif action == 'fun2':
            excel_split_by_type(file_path, user_id)
        elif action == 'fun3':
            excel_split_by_type_to_one(file_path, user_id)
        else:
            print('No action!')
        
        shutil.copy(f'{user_dir}/output_{user_id}.zip', f'output/output_{user_id}.zip') # 复制压缩包到output目录
        shutil.rmtree(user_dir) # 删除用户目录

        return static_file(f'output_{user_id}.zip', root='output', download=True)

@app.route('/upload2', method='POST')
def upload2():
    if not os.path.exists('uploads'):
        os.makedirs('uploads')
    
    upload_files = request.files.getall('file[]')
    filenames = []
    sheet_names = []
    if upload_files:
        delete_files_in_dir(f'output') # 清空output目录
        user_id = generate_unique_name() # 创建唯一的用户id
        user_dir = f'uploads/{user_id}'
        os.makedirs(user_dir) # 创建用户目录

        for upload_file in upload_files:
            temp_upload_file = upload_file.raw_filename
            # print(temp_upload_file)
            # 判断文件是否为Excel文件
            if temp_upload_file.endswith('.xlsx') or temp_upload_file.endswith('.xls'):
                # 提取文件名（不带后缀）作为Sheet名称
                ex_name = os.path.basename(temp_upload_file)
                sheet_name = ex_name.split('.')[0]
                sheet_names.append(sheet_name)

                file_path = f'uploads/{user_id}/{upload_file.raw_filename}'
                filenames.append(file_path)
                upload_file.save(file_path)
                # print(file_path)
        
        action = request.forms.get('action')
        if action == 'fun4':
            merge_excels_into_sheets(filenames, sheet_names, user_id)
        elif action == 'fun5':
            merge_excels_into_one_sheet(filenames, user_id)
        else:
            print('No action!')
        
        shutil.copy(f'{user_dir}/output_{user_id}.zip', f'output/output_{user_id}.zip') # 复制压缩包到output目录
        shutil.rmtree(user_dir) # 删除用户目录

        return static_file(f'output_{user_id}.zip', root='output', download=True)

# 按时间生成唯一的名字
def generate_unique_name():
    now = datetime.datetime.now()
    timestamp = now.strftime('%Y%m%d%H%M%S')

    letters = string.ascii_letters + string.digits
    random_digits = ''.join(random.choice(letters) for _ in range(4))
    return timestamp + random_digits

# 删除目录下的所有文件
def delete_files_in_dir(directory):
    for filename in os.listdir(directory):
        file_path = os.path.join(directory, filename)
        if os.path.isfile(file_path):
            os.remove(file_path)

# 1.把Excel文件按Sheet工作表保存成n个Excel文件
def excel_to_files_by_sheet(file_path, user_id):
    excel_data = pd.read_excel(file_path, sheet_name=None)
    user_dir = f'uploads/{user_id}'
    zip_path = f'{user_dir}/output_{user_id}.zip'
    with zipfile.ZipFile(zip_path, 'w') as zip_file:
        for sheet_name, sheet_data in excel_data.items():
            sheet_file_path = f'{user_dir}/{sheet_name}.xlsx'
            sheet_data.to_excel(sheet_file_path, index=False)
            zip_file.write(sheet_file_path, arcname=f'{sheet_name}.xlsx')
            os.remove(sheet_file_path)

# 2.表格按A列标题“type”拆分成n个表格,保留标题行
def excel_split_by_type(file_path, user_id):
    excel_data = pd.read_excel(file_path)
    column_name = 'type'
    types = excel_data[column_name].unique()
    user_dir = f'uploads/{user_id}'
    zip_path = f'{user_dir}/output_{user_id}.zip'
    with zipfile.ZipFile(zip_path, 'w') as zip_file:
        for t in types:
            filtered_data = excel_data[excel_data[column_name] == t]
            sheet_file_path = f'{user_dir}/{t}.xlsx'
            filtered_data.to_excel(sheet_file_path, index=False)
            zip_file.write(sheet_file_path, arcname=f'{t}.xlsx')
            os.remove(sheet_file_path)

# 3.表格按A列标题“type”拆分成n个sheet工作表,保留标题行
def excel_split_by_type_to_one(file_path, user_id):
    excel_data = pd.read_excel(file_path)
    column_name = 'type'
    types = excel_data[column_name].unique()
    user_dir = f'uploads/{user_id}'
    xlsx_path = f'{user_dir}/output_{user_id}.xlsx'
    with pd.ExcelWriter(xlsx_path, engine='openpyxl') as writer:
        for t in types:
            filtered_data = excel_data[excel_data[column_name] == t]
            filtered_data.to_excel(writer, sheet_name=str(t), index=False)  # 将t转换为字符串
        
    zip_output(xlsx_path, user_id)

# 4.把多个Excel文件中的数据合并到一个Excel文件的不同sheet中 merged.xlsx
def merge_excels_into_sheets(file_paths, sheet_names, user_id):
    user_dir = f'uploads/{user_id}'
    xlsx_path = f'{user_dir}/output_{user_id}.xlsx'
    with pd.ExcelWriter(xlsx_path, engine='xlsxwriter') as writer:
        for file_path in file_paths:
                sheet_name = sheet_names[file_paths.index(file_path)]
                df = pd.read_excel(file_path)
                df.to_excel(writer, sheet_name=sheet_name, index=False)
    zip_output(xlsx_path, user_id)

# 5.把多个Excel文件中的数据合并到一个Excel文件的一个sheet中
def merge_excels_into_one_sheet(file_paths, user_id):
    user_dir = f'uploads/{user_id}'
    xlsx_path = f'{user_dir}/output_{user_id}.xlsx'
    merged_data = pd.DataFrame() # 创建一个空的DataFrame用于存储合并后的数据
    for file_path in file_paths:
        df = pd.read_excel(file_path)
        merged_data = pd.concat([merged_data, df], ignore_index=True, sort=False)
    merged_data.to_excel(xlsx_path, index=False)
    zip_output(xlsx_path, user_id)

# 把单个文件打包成 output.zip
def zip_output(xlsx_path, user_id):
    if os.path.exists(xlsx_path):
        zip_path = f'uploads/{user_id}/output_{user_id}.zip'
        with zipfile.ZipFile(zip_path, 'w') as zip_file:
            zip_file.write(xlsx_path, arcname=f'output_{user_id}.xlsx')
        # os.remove(xlsx_path)
    else:
        print(f"File {xlsx_path} does not exist.")

if __name__ == '__main__':
    app.run(host='0.0.0.0', port='9881')