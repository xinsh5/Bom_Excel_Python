#!/usr/bin/env python
# !encoding:utf-8

#  注意：pandas使用的是Dataframe结构，行、列索引从0开始，和Excel索引值不同！！！！

import time
import pandas as pd
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



#   标准化电容值方法       
def convert_capacitance_unit(value):
    cap_unit_ls = ['mF', 'uF', 'nF', 'pF', 'fF'] 
    cap_unit_ls_capital = ['MF', 'UF', 'NF', 'PF', 'FF']
    cap_value = 0.0
    index = 0
    cap_other = ''
    value_list = ['','','']
    value=str(value)  #强制转成成字符串，便于后面处理
    if value[0].isdigit():        
        if value.isdigit() : #如果没写单位，默认是pF
            value = int(value[:-1]) * (10**int(value[-1])) #按照科学计数法计算值
            value = str(value) + 'pF' 
        #以下判断不同分隔符，将value内容拆分成列表
        value_list = value.split('-') #以‘-’分隔字符串
        if value_list[0]== value:  #如果不能以‘-’分隔，则返回整个字符串
            value_list = value.split('/')  #以‘/’分隔字符串  
        else:
            if value_list[0].isdigit():
                value = value_list[1] + '/' + value_list[2]
        #以下循环重新修正容值表示方法
        for j in range(len(cap_unit_ls)): #循环时跳过大写
            for i in range(len(value)): 
                if value[i : i+2] == cap_unit_ls[j] or value[i : i+2] == cap_unit_ls_capital[j]: #需要判断是大写和小写
                    cap_value = float(value[:i])
                    index = j
                    cap_other = value[i+2:]
                    break
            if cap_value != 0.0:
                break
        if cap_value >= 1000000:
            cap_value = cap_value / 1000000
            index = index - 2
        elif cap_value >= 1000:
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

# 标准化电阻值方法       
def convert_resistance_unit(value):
    res_unit_ls = ['M', 'K', 'R', 'm'] 
    res_unit_ls_lower = ['M', 'k', 'r', 'm']
    res_value = 0.0
    position = -1
    index = 0
    res_other = ''
    value_list = ['','','']
    value=str(value)  #强制转成成字符串，便于后面处理
    if value[0].isdigit():        
        # if value.isdigit() : #如果没写单位，默认是R（Ω）
        #     value = str(value) + 'R' 
        #以下判断不同分隔符，将value内容拆分成列表
        value_list = value.split(' ') #以‘ ’(空格)分隔字符串
        if value_list[0]== value:  #如果不能以‘ ’(空格)分隔，则返回整个字符串
            value_list = value.split('/')  #以‘/’分隔字符串  
        if is_number(value_list[0]):
            value = value_list[0] + 'R'
        else:
            value = value_list[0]
        #以下循环重新修正阻值表示方法
        print("resistance value is:",value)
        for j in range(len(res_unit_ls)):
            position = value.find( res_unit_ls[j])  #查找单位的位置，先查找大写字母 
            if position == -1:
                position = value.find(res_unit_ls_lower[j])  #查找单位的位置，查找小写字母
                if  position == -1:  #如果没找到则查找下一个单位
                    pass
                else:
                    value.replace(res_unit_ls_lower[j],res_unit_ls[j],1) #单位转换为大写                        
            if position == -1:
                pass
            elif position != (len(value)-1):  #如果单位不是最后一个字符，则说明是小数点位置，替换成"."
                print(" position:",position)
                value = value.replace(res_unit_ls[j],'.',1) + res_unit_ls[j]
                print(" RES Value:",value)
                res_value = float(value[:len(value)-1])
                index = j               
            elif position == (len(value)-1):  
                res_value = float(value[:len(value)-1]) 
                index = j
            if res_value != 0.0:
                pass        
        print(print("res_Value:",res_value))
        if res_value >= 1000000:
            res_value = res_value / 1000000
            index = index - 2
        elif res_value >= 1000:
            res_value = res_value / 1000
            index = index - 1
        # elif res_value < 1:
        #     res_value = res_value * 1000
        #     index = index + 1
        else:
            pass
        if res_value.is_integer():
            res_value = int(res_value)
        value = str(res_value) + res_unit_ls[index]  
        for i in range(1,len(value_list),1):
            value = value + '/' + value_list[i]         
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
VALUE_COLUMN = 3   #清单元器件值所在的列号
FOOTPRINT_COLUMN = 4 #清单元器件封装所在的列号
REVISED_VALUE_COLUMN =5 #清单新增修正后的元器件型号所在的列号
MANUFACTORY_PART_NUM_COLUMN =6 #清单制造商型号所在列号
MANUFACTORY_COLUMN =7 #清单“厂家”所在列号 
COMMENT_COLUMN = 8 #清单“备注”注释所在列号
MODEL_NUM_COLUMN = 8  #清单元器件型号所在的列号

REF_REFERENCE_COLUMN = 2   #参考清单元器件位号所在的列号
REF_VALUE_COLUMN = 8   #参考清单元器件值所在的列号
REF_FOOTPRINT_COLUMN = 9 #参考清单元器件封装所在的列号
REF_REVISED_VALUE_COLUMN =10 #参考清单新增修正后的元器件型号所在的列号
REF_MANUFACTORY_PART_NUM_COLUMN =4 #参考清单制造商型号所在列号
REF_MANUFACTORY_COLUMN =6 #参考清单“厂家”所在列号 

search_result = False

part_count=0  #器件数量
first_loop_end_flag = False

File_Name='./Bom/hongyun导出清单_20240212.xlsx'  #原始文件名,
New_File_Name='./Bom/hongyun_V01清单_value.xlsx'  #输出的文件名
Ref_File_Name = "./Bom/2100215381-料况-天路元器件件清单_焊接20231211.xlsx" #参考清单，用于读取制造商型号
File_Log_Name = './Bom/BOM_Excel_pd_value.log'  #记录的日志文件
file_log = open(File_Log_Name,'w')  #打开记录文件
file_log.write("时间："+f"{current_year}"+"年"+ f"{current_month}"+"月"+f"{current_day}"+"日"
               +f"{current_hour}"+"时"+f"{current_minute}"+"分"+f"{current_second}"+"秒"+"\n")
file_log.write("读出的文件："+File_Name+"\n")
file_log.write("写入的文件："+New_File_Name+"\n")

df = pd.read_excel(File_Name,sheet_name=0) #读第一个sheet内容
ref_df = pd.read_excel(Ref_File_Name,sheet_name=0) #读参考清单第一个sheet内容
# print(df)
# df.reset_index()

max_rows = df.shape[0]  #获取最大行数
max_columes = df.shape[1]  #获取最大列数
ref_max_rows = ref_df.shape[0]  #获取参考清单最大行数
ref_max_columes = ref_df.shape[1]  #获取参考清单最大列数
print(f"原始最大行数：{max_rows}")
file_log.write(f"原始最大行数：{max_rows}"+"\n")
print(f"原始最大列数：{max_columes}")
file_log.write(f"原始最大列数：{max_columes}"+"\n")
print(f"参考清单最大行数：{ref_max_rows}")
file_log.write(f"参考清单最大行数：{ref_max_rows}"+"\n")
print(f"参考清单最大列数：{ref_max_columes}")
file_log.write(f"参考清单最大列数：{ref_max_columes}"+"\n")

# 存储错误行的序号
# error_row_num = bom_components_begin_row_num
# 位号不允许重复 存储位号
# reference_ls = []
# 检查BOM表是否出错，前两列应全为空或全为数字，第三列应为序号，应是字母+数字的形式且不能重复，第四列为Value和第五列PCB Footprint应不为空
# is_valid_reference_pattern = r"^[A-Za-z]+\d+$"
# for row in origin_sheet.iter_rows(min_row=bom_components_begin_row_num, values_only=True):
#     if any(row):  # 判断整行是否存在非空值，为空则跳过
#         if ((type(row[0]) == int and type(row[1]) == int) or (row[0] == None and row[1] == None)) == False:
#             pyautogui.alert(f'第{error_row_num}行={row}的前两列格式错误，非空，也非数字', '提示')
#             sys.exit()
#         if (re.match(is_valid_reference_pattern, row[2]) == False):
#             pyautogui.alert(f'第{error_row_num}行={row}的前三列格式错误，并非位号', '提示')
#             sys.exit()
#         elif (row[2].strip() in reference_ls):
#             pyautogui.alert(f'第{error_row_num}行={row}的位号重复', '提示')
#             sys.exit()
#         else:
#             reference_ls.append(row[2].strip())
#         if row[3] == None:
#             pyautogui.alert(f'第{error_row_num}行={row}的前四列格式错误, Value缺失', '提示')
#             sys.exit()
#         if row[4] == None:
#             pyautogui.alert(f'第{error_row_num}行={row}的前五列格式错误, PCB Footprint缺失', '提示')
#             sys.exit()
#         error_row_num += 1
# pyautogui.alert('此BOM表格式正确!', '确认')



# dup=df.duplicated("客户型号",keep=False)
# print("重复数据：\n",df[dup])
# print("重复的行：\n",dup)
# # print("重复的行号：\n",dup)
# # file_log.write("重复的数据：\n")
# # file_log.write(str(df[dup]) + "\n")
# file_log.write("重复的行号：\n")
# file_log.write(str(dup) + "\n")

# for i in range(initial_line_num,max_rows,1):
# i=INITIAL_LINE_NUM 
# while i<max_rows :
#     repetition_flag = False  #重复标志
#     if i > INITIAL_LINE_NUM:
#         first_loop_end_flag = True

#     part_count = df.iloc[i,QUANTITY_COLUMN]
#     reference_num_list = df.iloc[i,REFERENCE_COLUMN]
#     # for j in range(i+1,max_rows,1):
#     j=i+1
#     while j<max_rows :
#         if pd.isnull(df.iloc[j,MODEL_NUM_COLUMN]):  #如果型号单元格为空则跳过
#             if first_loop_end_flag == False:
#                 print(f"第{j}行没有型号")
#                 file_log.write(f"单元格：{i},{j} 没有型号"+"\n")
#         elif df.iloc[i,MODEL_NUM_COLUMN] == df.iloc[j,MODEL_NUM_COLUMN]:  #如果型号相同则合并相同的行
#                 reference_num_list=reference_num_list+','+ df.iloc[j,REFERENCE_COLUMN]  #合并位号单元格内容
#                 part_count = part_count + df.iloc[j,QUANTITY_COLUMN]  #元器件数量相加
#                 df.drop(index=j,axis=0,inplace=True)  #删除后面相同的行,并且重排索引
#                 df.reset_index(drop=True,inplace=True)  #重排索引并更新
#                 max_rows -= 1  #更新最大行数
#                 print(f"重复的行号:{i},{j}")
#                 print(f"max_rows:{max_rows}")
#                 file_log.write(f"重复的行号:{i},{j}"+"\n")
#                 repetition_flag = True
#         j+=1        
#     if(repetition_flag):
#         print(f"重复单元格的内容:{reference_num_list}")
#         file_log.write(f"重复单元格的内容:\n{reference_num_list}"+"\n")
#         reference_num_list=sorted(reference_num_list.split(','),key=compare_func)  #排序，默认升序排序
#         print(f"合并单元格并排序完的内容:{reference_num_list}")
#         file_log.write(f"合并单元格并排序完的内容:\n{reference_num_list}"+"\n")
#         reference_num_list=','.join(reference_num_list)  #将列表转换成以逗号分隔的字符串
#         df.iloc[i,REFERENCE_COLUMN]=reference_num_list
#         df.iloc[i,QUANTITY_COLUMN]=part_count
#     i+=1
# for i in range(INITIAL_LINE_NUM,max_rows+1,1):
#     df.iloc[i-1,ITEM_COLUMN] = i

# print(f"内容：{df.iloc[41,8]} \n")
# if pd.isnull(df.iloc[11,8]):
# j=5
# print("第1\n")
# df.drop(index=j,axis=0,inplace=True)
# df.reset_index(drop=True,inplace=True)
# print(f"总行数：{df.shape[0]}")
# df.set_index(['Item'],inplace=True) #设置原清单中Item为索引，否则pandas会自动添加一列索引

#插入新列，用于记录整理后的元器件型号
df.insert(REVISED_VALUE_COLUMN,'Revised Value','')
#插入新列，用于记录制造商的元器件型号
df.insert(MANUFACTORY_PART_NUM_COLUMN,'Manufactory Part Num','')
#插入新列，用于记录厂家
df.insert(MANUFACTORY_COLUMN,'Manufactory','')
#插入新列，用于记录备注信息
df.insert(COMMENT_COLUMN,'Comment','')

max_rows = df.shape[0]  #获取最大行数
max_columes = df.shape[1]  #获取最大列数
#以下循环整理规范Value列里的型号，存放在新列Revised Value里
for i in range(0,max_rows,1):
    ref_result= df.iloc[i,2][0] #读取元器件位号首字母
    NC_result= str(df.iloc[i,3])[-2:]  #读取元器件值的最后2位字母，用于判断是否为/NC
    if NC_result=='NC':
        print(f"这些器件不焊接：{i}")
        df.iloc[i,REVISED_VALUE_COLUMN]=df.iloc[i,VALUE_COLUMN]  #NC器件的型号不变，保持后缀带NC
    else:
        if ref_result=='C': 
            df.iloc[i,REVISED_VALUE_COLUMN]=convert_capacitance_unit(df.iloc[i,VALUE_COLUMN]) #修订电容值，写到新列里   
        elif ref_result=='R':
            df.iloc[i,REVISED_VALUE_COLUMN]=convert_resistance_unit(df.iloc[i,VALUE_COLUMN]) #修订电阻值，写到新列里

#以下循环从参考清单里导入相同元器件的厂家和型号
for rows in range(0,max_rows,1):
    search_result = False
    for ref_rows in range(0,ref_max_rows,1):
        if (((df.iloc[rows,VALUE_COLUMN]==ref_df.iloc[ref_rows,REF_VALUE_COLUMN]) or \
            (df.iloc[rows,REVISED_VALUE_COLUMN]==ref_df.iloc[ref_rows,REF_VALUE_COLUMN])) and \
            (df.iloc[rows,FOOTPRINT_COLUMN]==ref_df.iloc[ref_rows,REF_FOOTPRINT_COLUMN])):
            df.iloc[rows,MANUFACTORY_PART_NUM_COLUMN]=ref_df.iloc[ref_rows,REF_MANUFACTORY_PART_NUM_COLUMN]
            df.iloc[rows,MANUFACTORY_COLUMN]=ref_df.iloc[ref_rows,REF_MANUFACTORY_COLUMN]
            search_result = True
    if search_result==False:
        df.iloc[rows,COMMENT_COLUMN]="没有相应的型号"

writer = pd.ExcelWriter(New_File_Name,engine='xlsxwriter') #使用ExcelWriter需要安装xlsxwriter模块：pip install xlsxwriter 
df.to_excel(writer, sheet_name='Sheet1', index=False)



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

# worksheet.set_row(1,None,row_even_format)
#   以下循环将Excel表格背景颜色设置为棋盘形式，间隔灰色背景。
# for column_num in range(0,max_columes+1,1):
#     for row_num in range(0,max_rows+1,1):
#         if column_num % 2 == 0:
#             if row_num % 2 == 0:
#                 # worksheet.set_row(i,None,row_even_format)
#                 worksheet.conditional_format(row_num,column_num,row_num,column_num, {'type':'no_errors','format': while_format})
#             else:
#                 worksheet.conditional_format(row_num,column_num,row_num,column_num, {'type':'no_errors','format': gray_format})
#         else:
#             if row_num % 2 == 0:
#                 # worksheet.set_row(i,None,row_odd_format)
#                 worksheet.conditional_format(row_num,column_num,row_num,column_num, {'type':'no_errors','format': gray_format})
#             else:
#                 worksheet.conditional_format(row_num,column_num,row_num,column_num, {'type':'no_errors','format': while_format})

#   以下循序将Excel表格背景颜色隔行设置为灰色
for row_num in range(0,max_rows+1,1):
    if row_num % 2 == 0:
        # worksheet.set_row(i,None,row_even_format)
        worksheet.conditional_format(row_num,0,row_num,(max_columes-1), {'type':'no_errors','format': gray_format})
    else:
        worksheet.conditional_format(row_num,0,row_num,(max_columes-1), {'type':'no_errors','format': while_format})
    

# worksheet.conditional_format(0,0,0,0, {'type':'no_errors','format': gray_format})

worksheet.set_column("A:A", 5, header_format) #设置A列宽度为10，格式为:垂直中信对齐；水平中心对齐
worksheet.set_column("B:B", 5, header_format1)
worksheet.set_column("C:C", 50,header_format2)
worksheet.set_column("D:D", 40,header_format3)
worksheet.set_column("E:I", 25,header_format4)
# worksheet.set_column("F:F", 30,header_format5)
# worksheet.set_column("G:G", 30,header_format5)
# worksheet.set_default_row(30)# 设置所有行高
# worksheet.set_row(0,15,header_format)#设置指定行

format_border = workbook.add_format({'border':5})   # 设置边框格式
worksheet.conditional_format('A1:XFD1048576',{'type':'no_blanks', 'format': format_border}) #整个工作表，根据条件来设置格式
# writer.save() #save()方法已经弃用？使用close()方法既是保存退出。

worksheet.freeze_panes(1,1)   # 冻结首行
worksheet.autofilter(0,0,max_rows,(max_columes-1))   # 添加筛选

writer.close()  #保存\退出


file_log.write("End.")
file_log.close()  #关闭日志文件

