#!/usr/bin/env python
# !encoding:utf-8

#  注意：pandas使用的是Dataframe结构，行、列索引从0开始，和Excel索引值不同！！！！

import time
import pandas as pd
pd.set_option('display.max_rows', None)   #pandas数据显示所有行，否则只显示前5行和后5行
pd.set_option('display.max_columns', None) #pandas数据显示所有列
pd.set_option('display.width', 500) #设置显示宽度，宽度要够大，否则一行显示不全

def compare_func(item):
    # 提取每个元素后面的数值部分并转换成int类型
    num = int(''.join([char for char in item if char.isdigit()]))
    return num

# 存储错误行的序号
error_row_num = bom_components_begin_row_num
# 位号不允许重复 存储位号
reference_ls = []
# 检查BOM表是否出错，前两列应全为空或全为数字，第三列应为序号，应是字母+数字的形式且不能重复，第四列为Value和第五列PCB Footprint应不为空
is_valid_reference_pattern = r"^[A-Za-z]+\d+$"
for row in origin_sheet.iter_rows(min_row=bom_components_begin_row_num, values_only=True):
    if any(row):  # 判断整行是否存在非空值，为空则跳过
        if ((type(row[0]) == int and type(row[1]) == int) or (row[0] == None and row[1] == None)) == False:
            pyautogui.alert(f'第{error_row_num}行={row}的前两列格式错误，非空，也非数字', '提示')
            sys.exit()
        if (re.match(is_valid_reference_pattern, row[2]) == False):
            pyautogui.alert(f'第{error_row_num}行={row}的前三列格式错误，并非位号', '提示')
            sys.exit()
        elif (row[2].strip() in reference_ls):
            pyautogui.alert(f'第{error_row_num}行={row}的位号重复', '提示')
            sys.exit()
        else:
            reference_ls.append(row[2].strip())
        if row[3] == None:
            pyautogui.alert(f'第{error_row_num}行={row}的前四列格式错误, Value缺失', '提示')
            sys.exit()
        if row[4] == None:
            pyautogui.alert(f'第{error_row_num}行={row}的前五列格式错误, PCB Footprint缺失', '提示')
            sys.exit()
        error_row_num += 1

# pyautogui.alert('此BOM表格式正确!', '确认')

#   标准化电容值方法       
def convert_capacitance_unit(value):
    cap_unit_ls = ['mF', 'uF', 'nF', 'pF', 'fF']
    cap_value = 0.0
    index = 0
    cap_other = ''
    for j in range(len(cap_unit_ls)):
        for i in range(len(value)): 
            if value[i : i+2] == cap_unit_ls[j]:
                cap_value = float(value[:i])
                index = j
                cap_other = value[i+2:]
                break
        if cap_value != 0.0:
            break
    if cap_value >= 1000:
        cap_value = cap_value / 1000
        index = index - 1
    elif cap_value < 1:
        cap_value = cap_value * 1000
        index = index + 1
    else:
        pass
    if cap_value.is_integer():
        cap_value = int(cap_value)
    value = str(cap_value) + cap_unit_ls[index] + cap_other
    return value


current_time_struct = time.localtime()  #获取当前时间
# 分别获取当前年、月、日、时、分、秒
current_year = current_time_struct.tm_year
current_month = current_time_struct.tm_mon
current_day = current_time_struct.tm_mday
current_hour = current_time_struct.tm_hour
current_minute = current_time_struct.tm_min
current_second = current_time_struct.tm_sec

INITIAL_LINE_NUM = 1  #初始行号

ITEM_COLUMN =  0    #清单Ietm所在的列号
QUANTITY_COLUMN = 1   #清单元器件数量所在的列号
REFERENCE_COLUMN = 2   #清单元器件位号所在的列号
MODEL_NUM_COLUMN = 8  #清单元器件型号所在的列号

part_count=0  #器件数量

first_loop_end_flag = False

File_Name='hongyun_V01导出清单_20240129_0130.xlsx'  #原始文件名
New_File_Name='hongyun_V01清单_pd.xlsx'  #输出的文件名
File_Log_Name = 'BOM_Excel_pd.log'  #记录的日志文件
file_log = open(File_Log_Name,'w')  #打开记录文件
file_log.write("时间："+f"{current_year}"+"年"+ f"{current_month}"+"月"+f"{current_day}"+"日"
               +f"{current_hour}"+"时"+f"{current_minute}"+"分"+f"{current_second}"+"秒"+"\n")
file_log.write("读出的文件："+File_Name+"\n")
file_log.write("写入的文件："+New_File_Name+"\n")

df = pd.read_excel(File_Name)
# df.reset_index()

max_rows = df.shape[0]  #获取最大行数
max_columes = df.shape[1]  #获取最大列数
print(f"原始最大行数：{max_rows}")
file_log.write(f"原始最大行数：{max_rows}"+"\n")
print(f"原始最大列数：{max_columes}")
file_log.write(f"原始最大列数：{max_columes}"+"\n")

dup=df.duplicated("客户型号",keep=False)
print("重复数据：\n",df[dup])
print("重复的行：\n",dup)
# print("重复的行号：\n",dup)
# file_log.write("重复的数据：\n")
# file_log.write(str(df[dup]) + "\n")
file_log.write("重复的行号：\n")
file_log.write(str(dup) + "\n")

# for i in range(initial_line_num,max_rows,1):
i=INITIAL_LINE_NUM 
while i<max_rows :
    repetition_flag = False  #重复标志
    if i > INITIAL_LINE_NUM:
        first_loop_end_flag = True

    part_count = df.iloc[i,QUANTITY_COLUMN]
    reference_num_list = df.iloc[i,REFERENCE_COLUMN]
    # for j in range(i+1,max_rows,1):
    j=i+1
    while j<max_rows :
        if pd.isnull(df.iloc[j,MODEL_NUM_COLUMN]):  #如果型号单元格为空则跳过
            if first_loop_end_flag == False:
                print(f"第{j}行没有型号")
                file_log.write(f"单元格：{i},{j} 没有型号"+"\n")
        elif df.iloc[i,MODEL_NUM_COLUMN] == df.iloc[j,MODEL_NUM_COLUMN]:  #如果型号相同则合并相同的行
                reference_num_list=reference_num_list+','+ df.iloc[j,REFERENCE_COLUMN]  #合并位号单元格内容
                part_count = part_count + df.iloc[j,QUANTITY_COLUMN]  #元器件数量相加
                df.drop(index=j,axis=0,inplace=True)  #删除后面相同的行,并且重排索引
                df.reset_index(drop=True,inplace=True)  #重排索引并更新
                max_rows -= 1  #更新最大行数
                print(f"重复的行号:{i},{j}")
                print(f"max_rows:{max_rows}")
                file_log.write(f"重复的行号:{i},{j}"+"\n")
                repetition_flag = True
        j+=1        
    if(repetition_flag):
        print(f"重复单元格的内容:{reference_num_list}")
        file_log.write(f"重复单元格的内容:\n{reference_num_list}"+"\n")
        reference_num_list=sorted(reference_num_list.split(','),key=compare_func)  #排序，默认升序排序
        print(f"合并单元格并排序完的内容:{reference_num_list}")
        file_log.write(f"合并单元格并排序完的内容:\n{reference_num_list}"+"\n")
        reference_num_list=','.join(reference_num_list)  #将列表转换成以逗号分隔的字符串
        df.iloc[i,REFERENCE_COLUMN]=reference_num_list
        df.iloc[i,QUANTITY_COLUMN]=part_count
    i+=1
for i in range(INITIAL_LINE_NUM,max_rows+1,1):
    df.iloc[i-1,ITEM_COLUMN] = i

# print(f"内容：{df.iloc[41,8]} \n")
# if pd.isnull(df.iloc[11,8]):
# j=5
# print("第1\n")
# df.drop(index=j,axis=0,inplace=True)
# df.reset_index(drop=True,inplace=True)
# print(f"总行数：{df.shape[0]}")
# df.set_index(['Item'],inplace=True) #设置原清单中Item为索引，否则pandas会自动添加一列索引
df.to_excel(New_File_Name)  #写入新的excel文件

file_log.write("End.")
file_log.close()  #关闭日志文件

