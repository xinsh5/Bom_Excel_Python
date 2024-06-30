#!/usr/bin/env python
# !encoding:utf-8
##################
#这个程序是将bom清单里位号分割成25所标准化模版里元件表的格式，每2个位号一行，并计算每行的数量；


#  注意：pandas使用的是Dataframe结构，行、列索引从0开始，和Excel索引值不同！！！！

import time
import tkinter as tk
from tkinter import filedialog
import pandas as pd
import os
import re
# from pandas import options
# import xlsxwriter 
pd.set_option('display.max_rows', None)   #pandas数据显示所有行，否则只显示前5行和后5行
pd.set_option('display.max_columns', None) #pandas数据显示所有列
pd.set_option('display.width', 500) #设置显示宽度，宽度要够大，否则一行显示不全

def is_number(str):
  try:
    # 因为使用float有一个例外是'NaN'
    if str=='NaN':
        return False
    float(str)
    return True
  except ValueError:
    return False

def compare_func(item):
    # 提取每个元素后面的数值部分并转换成int类型
    num = int(''.join([char for char in item if char.isdigit()]))
    return num

def extract_letters_from_list(data_list):
    # 使用正则表达式来提取前面的字母
    letters = [re.match(r'[A-Za-z]+', item).group() for item in data_list]
    return letters

def extract_numbers_from_list(data_list):
    # 使用正则表达式来提取数字
    numbers = [int(re.search(r'\d+', item).group()) for item in data_list]
    return numbers

def convert_to_ranges(nums):
    nums = sorted(set(nums))  # 去重并排序
    ranges = []
    start = nums[0]
    end = nums[0]

    for num in nums[1:]:
        if num == end + 1:
            end = num
        else:
            ranges.append((start, end))
            start = num
            end = num

    ranges.append((start, end))  # 别忘了最后一个范围
    return ranges

def convert_to_range_list(nums):
   
    result = []
    i = 0
    n = len(nums)
    
    while i < n:
        start = nums[i]
        j = i
        
        # 找到连续段落的结尾
        while j < n - 1 and nums[j] + 1 == nums[j + 1]:
            j += 1
        
        if j - i + 1 >= 3:
            # 如果段落长度为3或以上，转换为起止格式
            result.append((start, nums[j]))
        else:
            # 否则逐个保留原样
            for k in range(i, j + 1):
                result.append((nums[k], nums[k]))
        
        i = j + 1
    
    return result


def format_ranges(ranges,prefix):
    return ', '.join([f"{prefix}{start}~{prefix}{end}" if start != end else f"{prefix}{start}" for start, end in ranges])

# def format_ranges(ranges, prefix):
#     result = []
#     for item in ranges:
#         if isinstance(item, tuple):
#             start, end = item
#             result.append(f"{prefix}{start}~{prefix}{end}")
#         else:
#             result.append(f"{prefix}{item}")
#     return result

 

# def add_prefix_to_ranges(range_list, prefix):
#     # 使用列表解析在每个范围字符串前面添加特定字母
#     prefixed_list = [f"{prefix}{range_str}" for range_str in range_list]
#     return prefixed_list
def add_prefix_to_ranges_and_single_numbers(range_list, prefix):
    prefixed_list = []
    for item in range_list:
        if '-' in item:  # 处理范围字符串
            start, end = item.split('-')
            prefixed_range = f"{prefix}{start}-{prefix}{end}"
            prefixed_list.append(prefixed_range)
        else:  # 处理单个数字
            prefixed_number = f"{prefix}{item}"
            prefixed_list.append(prefixed_number)
    return prefixed_list

def parse_and_split(data):
    """解析并拆分连续和不连续的数据"""
    items = data.split(',')
    result = []
    temp = [items[0]]

    for i in range(1, len(items)):
        prev_num = int(re.findall(r'\d+', items[i-1])[0])
        curr_num = int(re.findall(r'\d+', items[i])[0])

        if curr_num == prev_num + 1:
            temp.append(items[i])
        else:
            result.append(temp)
            temp = [items[i]]
    
    result.append(temp)
    return result

def extract_number(element):
    """
    提取元素中的数字部分
    """
    match = re.search(r'\d+', element)
    return int(match.group()) if match else None

def process_cell_content(cell_content):
    # 提取单元格内容并将其解析为列表
    elements = [elem.strip() for elem in cell_content.split(',')]

    # 初始化结果列表
    results = []

    # 处理元素并分组
    i = 0
    while i < len(elements):
        if '~' in elements[i]:
            # 如果元素包含~，则为一组
            start, end = elements[i].split('~')
            start_num = extract_number(start)
            end_num = extract_number(end)
            quantity = end_num - start_num + 1
            results.append((elements[i], quantity))
            i += 1
        else:
            # 检查下一个元素
            if i + 1 < len(elements) and '~' not in elements[i + 1]:
                group = f"{elements[i]},{elements[i + 1]}"
                quantity = 2
                i += 2
            else:
                group = elements[i]
                quantity = 1
                i += 1
            results.append((group, quantity))

    return results


def import_excel(file1):

    current_time_struct = time.localtime()  #获取当前时间
    # 分别获取当前年、月、日、时、分、秒
    current_year = current_time_struct.tm_year
    current_month = current_time_struct.tm_mon
    current_day = current_time_struct.tm_mday
    current_hour = current_time_struct.tm_hour
    current_minute = current_time_struct.tm_min
    current_second = current_time_struct.tm_sec

    INITIAL_LINE_NUM = 0  #初始行号

    ITEM_COLUMN =  0    #清单Ietm所在的列号
    QUANTITY_COLUMN = 1   #清单元器件数量所在的列号
    REFERENCE_COLUMN = 2   #清单元器件位号所在的列号
    MODEL_NUM_COLUMN = 7   #清单元器件最终型号所在的列号
    FOOTPRINT_COLUMN = 4 #清单元器件封装所在的列号
    REVISED_VALUE_COLUMN =5 #清单新增修正后的元器件型号所在的列号
    PART_NAME = 6 #清单元器名称所在列号
    MANUFACTORY_PART_NUM_COLUMN =7 #清单制造商型号所在列号
    MANUFACTORY_COLUMN =8 #清单“厂家”所在列号 
    COMMENT_COLUMN = 9 #清单“备注”注释所在列号
    VENDOR_COLUMN = 10  #清单元器件供货方所在的列号，如：'客供','一博供货'等
    NEW_QUANTITY_COLUMN = 11   #合并相同型号后的元器件数量所在的列号
    NEW_REFERENCE_COLUMN = 12   #合并相同型号后的元器件位号所在的列号
    

    REF_REFERENCE_COLUMN = 2   #参考清单元器件位号所在的列号
    REF_PART_NAME = 6 #清单元器名称所在列号
    REF_VALUE_COLUMN = 3   #参考清单元器件值所在的列号
    REF_FOOTPRINT_COLUMN = 4 #参考清单元器件封装所在的列号
    REF_REVISED_VALUE_COLUMN =5 #参考清单新增修正后的元器件型号所在的列号
    REF_STANDARD_COLUMN = 15 #标准号所在的列号
    REF_MODEL_COLUMN = 13   #清单元器件最终型号所在的列号
    REF_MANUFACTORY_PART_NUM_COLUMN =30 #参考清单制造商型号所在列号
    REF_CUSTOMER_MANUFACTORY_PART_NUM_COLUMN =7 #参考清单制造商型号所在列号
    REF_MANUFACTORY_COLUMN =29 #参考清单“厂家”所在列号 
    REF_CUSTOMER_MANUFACTORY_COLUMN =18 #参考清单“厂家”所在列号 
    REF_VENDOR_COLUMN = 34 #参考清单“物料提供方式”所在列号

    search_result = False

    part_count=0  #器件数量
    first_loop_end_flag = False

    File_Name=file1
    New_File_Name =os.path.splitext(file1)[0] + '_SplitRef' + os.path.splitext(file1)[1]
    File_Log_Name = os.path.splitext(file1)[0] + '_SplitRef' + '.log'  #记录的日志文件
    file_log = open(File_Log_Name,'w')  #打开记录文件
    file_log.write("时间："+f"{current_year}"+"年"+ f"{current_month}"+"月"+f"{current_day}"+"日"
                    +f"{current_hour}"+"时"+f"{current_minute}"+"分"+f"{current_second}"+"秒"+"\n")
    file_log.write("读出的文件："+File_Name+"\n")
    file_log.write("写入的文件："+New_File_Name+"\n")
   
    df = pd.read_excel(File_Name)  

    # 指定列的数据类型
    # dtype_dict = {
    #     '序号': int,
    #     '项目代号':str,
    #     '代号': str,
    #     '名称和型号':str,
    #     '数量':int,
    #     '备注':str
    # }  
    
    processed_df = pd.DataFrame({}) #columns=['序号','项目代号','代号','名称和型号','数量','备注'],dtype=dtype_dict)
    # processed_df = pd.DataFrame(columns=['序号','项目代号','代号','名称和型号','数量','备注'])
    processed_df['序号']=''
    processed_df['项目代号']=''
    processed_df['代号']=''
    processed_df['名称和型号']=''
    processed_df['数量']=''
    processed_df['备注']=''

    new_row={'序号':'','项目代号':'','代号':'','名称和型号':'','数量':'','备注':''}

    max_rows = df.shape[0]  #获取最大行数
    max_columes = df.shape[1]  #获取最大列数
     #插入新列，用于记录原始行索引号
   
    print(f"原始最大列数：{max_columes}")
    file_log.write(f"原始最大列数：{max_columes}"+"\n") 
      
    
    i=INITIAL_LINE_NUM 
    current_row = 0
    while i<max_rows :
        if pd.isnull(df.iloc[i,QUANTITY_COLUMN]):
            print(f"第{df.iloc[i,max_columes-1]}行没有数量")                   
            file_log.write(f"单元格：{df.iloc[i,max_columes-1]} 没有数量"+"\n")
        part_count = df.iloc[i,QUANTITY_COLUMN]
        if pd.isnull(df.iloc[i,REFERENCE_COLUMN]):
            print(f"第{df.iloc[i,max_columes-1]}行没有位号")                   
            file_log.write(f"单元格：{df.iloc[i,max_columes-1]} 没有位号"+"\n")
        reference_num_list = df.iloc[i,REFERENCE_COLUMN]
        print(f"原始位号的内容:{reference_num_list}")
        file_log.write(f"原始位号的内容:\n{reference_num_list}"+"\n")
       
        reference_num_list_num=process_cell_content(reference_num_list)
        print(f"转换完的内容:{reference_num_list_num}")
        file_log.write(f"转换完的内容:\n{reference_num_list_num}"+"\n")
        j=0                
        for group, quantity in reference_num_list_num: 
                                   
            processed_df.loc[len(processed_df)]=new_row
            # processed_df = pd.concat([processed_df, new_row], ignore_index=True)
                                   
            # processed_df = pd.concat([processed_df, pd.DataFrame({'序号':{i},'项目代号':group,'代号':df.iloc[i,REF_STANDARD_COLUMN],\
            #                      '名称和型号':df.iloc[i,REF_MODEL_COLUMN],'数量':{quantity},'备注':df.iloc[i,REF_CUSTOMER_MANUFACTORY_COLUMN]})], ignore_index=True)
            processed_df.iloc[current_row,0]=i
            processed_df.iloc[current_row,1]=group
            if pd.isnull(df.iloc[i,REF_MODEL_COLUMN]):
                processed_df.iloc[current_row,2]=0
                processed_df.iloc[current_row,3]=0
                processed_df.iloc[current_row,5]=0
            else:
                processed_df.iloc[current_row,2]=df.iloc[i,REF_STANDARD_COLUMN]
                processed_df.iloc[current_row,3]=df.iloc[i,REF_MODEL_COLUMN]
                processed_df.iloc[current_row,5]=df.iloc[i,REF_CUSTOMER_MANUFACTORY_COLUMN]
            processed_df.iloc[current_row,4]=quantity
            
            current_row = current_row + 1
            j += 1 
        # next_row_step=j    
        i+=1

    max_rows = processed_df.shape[0]  #获取最大行数
    max_columes = processed_df.shape[1]  #获取最大列数
    writer = pd.ExcelWriter(New_File_Name,engine='xlsxwriter') #使用ExcelWriter需要安装xlsxwriter模块：pip install xlsxwriter 
    processed_df.to_excel(writer, sheet_name='Sheet1', index=False)

    workbook = writer.book
    worksheet = writer.sheets['Sheet1']

    header_format = workbook.add_format({
        'valign': 'vcenter',  # 垂直对齐方式
        'align': 'center', # 水平对齐方式
        # 'bg_color':'#C0C0C0', #设置背景颜色，也可以用'green'
    })
    header_format1 = workbook.add_format({
        'valign': 'vcenter',  # 垂直对齐方式
        'align': 'center', # 水平对齐方式
        # 'bg_color':'#C0C0C0', #设置背景颜色，也可以用'green'
    })
    header_format2 = workbook.add_format({
        'valign': 'vcenter', # 垂直对齐方式
        'align': 'left', # 水平对齐方式
        'text_wrap': True,  #自动换行
        # 'bg_color':'#C0C0C0', #设置背景颜色为灰色
        # 'font_color':'red'  #字体颜色：红色
        # 'italic':True       #字体为斜体
    }) 
    header_format3 = workbook.add_format({
        'valign': 'vcenter',  # 垂直对齐方式
        'align': 'left', # 水平对齐方式
        # 'bg_color':'#C0C0C0', #设置背景颜色为灰色
    })
    header_format4 = workbook.add_format({
        'valign': 'vcenter',  # 垂直对齐方式
        'align': 'left', # 水平对齐方式
        # 'bg_color':'#C0C0C0', #设置背景颜色为灰色
    })
    header_format5 = workbook.add_format({
        'valign': 'vcenter',  # 垂直对齐方式
        'align': 'left', # 水平对齐方式
        # 'bg_color':'#C0C0C0', #设置背景颜色为灰色
    })

    gray_format=workbook.add_format({
        'bg_color':'#C0C0C0', #设置背景颜色为灰色
        'text_wrap': True,  #自动换行
        'valign': 'vcenter',  # 垂直对齐方式
        'align': 'center', # 水平对齐方式
        'border':1
    })
    while_format=workbook.add_format({
        'bg_color':'#FFFFFF', #设置背景颜色为灰色
        'text_wrap': True,  #自动换行
        'valign': 'vcenter',  # 垂直对齐方式
        'align': 'center', # 水平对齐方式
        'border':1
    })

    #   以下循序将Excel表格背景颜色隔行设置为灰色
    for row_num in range(0,max_rows+1,1):
        if row_num % 2 == 0:
            # worksheet.set_row(i,None,row_even_format)
            worksheet.conditional_format(row_num,0,row_num,(max_columes-1), {'type':'no_errors','format': gray_format})
        else:
            worksheet.conditional_format(row_num,0,row_num,(max_columes-1), {'type':'no_errors','format': while_format})

    # worksheet.conditional_format(0,0,0,0, {'type':'no_errors','format': gray_format})
    worksheet.set_column("A:A", 8, header_format) #设置A列宽度为10，格式为:垂直中信对齐；水平中心对齐
    worksheet.set_column("B:B", 20, header_format1)
    worksheet.set_column("C:C", 35,header_format2)
    worksheet.set_column("D:D", 30,header_format2)
    worksheet.set_column("E:E", 8,header_format4)
    worksheet.set_column("F:F", 20,header_format5)
    worksheet.set_column("G:G", 20,header_format5)
    worksheet.set_column("H:H", 25,header_format5)
    worksheet.set_column("I:I", 25,header_format5)
    worksheet.set_column("J:J", 20,header_format5)
    worksheet.set_column("K:K", 20,header_format5)
    # worksheet.set_default_row(30)# 设置所有行高
    # worksheet.set_row(0,15,header_format)#设置指定行

    format_border = workbook.add_format({'border':5})   # 设置边框格式
    worksheet.conditional_format('A1:XFD1048576',{'type':'no_blanks', 'format': format_border}) #整个工作表，根据条件来设置格式
    # writer.save() #save()方法已经弃用？使用close()方法既是保存退出。

    worksheet.freeze_panes(1,0)   # 冻结首行，不冻结首列
    worksheet.autofilter(0,0,max_rows,(max_columes-1))   # 添加筛选
    
    writer.close()  #保存\退出
    file_log.write("End.")
    file_log.close()  #关闭日志文件
    return New_File_Name
def browse_file(entry):
    filename = filedialog.askopenfilename()
    entry.delete(0, tk.END)
    entry.insert(0, filename)

def Import():        
    output_text.insert(tk.END, "\n开始合并......\n") 
    file1 = entry_file1.get()
    # file2 = entry_file2.get()
    cmp_result_file = import_excel(file1)
    output_text.insert(tk.END, "\n相同型号元器件已经导入到新清单:\n")
    output_text.insert(tk.END, f"{cmp_result_file}\n") 

root = tk.Tk()
root.title("转换位号格式到新清单")

# 设置行和列的权重以使其可以拉伸，但保持间距不变
for i in range(4):
    root.grid_rowconfigure(i, weight=1, minsize=50)
root.grid_columnconfigure(1, weight=1)

# 设置窗口初始大小为原来的两倍
root.geometry("800x600")

label_file1 = tk.Label(root, text="原始清单:")
label_file1.grid(row=0, column=0, sticky="w", padx=10, pady=(10, 0))

entry_file1 = tk.Entry(root)
entry_file1.grid(row=0, column=1, padx=10, pady=(10, 0), sticky="ew")

browse_button1 = tk.Button(root, text="Browse", command=lambda: browse_file(entry_file1))
browse_button1.grid(row=0, column=2, padx=10, pady=(10, 0))

# label_file2 = tk.Label(root, text="参考清单:")
# label_file2.grid(row=1, column=0, sticky="w", padx=10, pady=(0, 5))

# entry_file2 = tk.Entry(root)
# entry_file2.grid(row=1, column=1, padx=10, pady=(0, 5), sticky="ew")

# browse_button2 = tk.Button(root, text="Browse", command=lambda: browse_file(entry_file2))
# browse_button2.grid(row=1, column=2, padx=10, pady=(0, 5))

# submit_button = tk.Button(root, text="Compare", command=submit)
# submit_button.grid(row=2, column=1, padx=10, pady=5)

Import_button = tk.Button(root, text="转换位号格式", command=Import)
Import_button.grid(row=2, column=1, padx=15, pady=5)

output_text = tk.Text(root, height=30, width=50)  # 增加了输出文本框的高度
output_text.grid(row=3, columnspan=3, padx=10, pady=(0, 10), sticky="nsew")

root.mainloop()


