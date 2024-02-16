#!/usr/bin/env python
# !encoding:utf-8

#  注意：pandas使用的是Dataframe结构，行、列索引从0开始，和Excel索引值不同！！！！

import time
import pandas as pd
import re
# from pandas import options
#####################################
#此文件使用了xlsxwriter 库，这个库不引用pandas是自动引用了，但是安装时没有自动一起安装，需要手动安装
# import xlsxwriter 此文件使用了
pd.set_option('display.max_rows', None)   #pandas数据显示所有行，否则只显示前5行和后5行
pd.set_option('display.max_columns', None) #pandas数据显示所有列
pd.set_option('display.width', 500) #设置显示宽度，宽度要够大，否则一行显示不全


# 自定义排序函数
def custom_sort_key(item):
    # 使用正则表达式提取列表元素中的数字部分
    num_part = re.findall(r'\d+', item)
    if num_part:
        # 如果存在数字部分，则返回数字部分的整数值和原始字符串组成的元组
        return (int(num_part[0]), item)
    else:
        # 如果没有数字部分，则返回一个元组，其中整数部分为无穷大，原始字符串为本身
        return (float('inf'), item)

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
record_file = False

part_count=0  #器件数量
first_loop_end_flag = False

# File_Name='./Bom/hongyun导出清单_20240213.xlsx'  #要比较文件名
# Ref_File_Name = "./Bom/hongyun_V01导出清单_20240126.xls" #被比较的清单
# File_Name='./Bom/hongyun_V01导出清单_20240126.xls'  #要比较文件名
# Ref_File_Name = "./Bom/hongyun导出清单_20240213.xlsx" #被比较的清单
File_Name='./Bom/hongyun导出清单_20240213.xlsx'  #要比较文件名
Ref_File_Name = "./Bom/hongyun导出清单_20240213_M.xlsx" #被比较的清单
New_File_Name='./Bom/hongyun_V01清单_compare.xlsx'  #输出的文件名

File_Log_Name = './Bom/BOM_Excel_pd_compare.log'  #记录的日志文件
if record_file:
    file_log = open(File_Log_Name,'w')  #打开记录文件
    file_log.write("时间："+f"{current_year}"+"年"+ f"{current_month}"+"月"+f"{current_day}"+"日"
                +f"{current_hour}"+"时"+f"{current_minute}"+"分"+f"{current_second}"+"秒"+"\n")
    file_log.write("读出的文件："+File_Name+"\n")
    file_log.write("写入的文件："+New_File_Name+"\n")

df = pd.read_excel(File_Name,sheet_name=0) #读第一个sheet内容
ref_df = pd.read_excel(Ref_File_Name,sheet_name=0) #读参考清单第一个sheet内容

max_rows = df.shape[0]  #获取最大行数
max_columes = df.shape[1]  #获取最大列数
ref_max_rows = ref_df.shape[0]  #获取参考清单最大行数
ref_max_columes = ref_df.shape[1]  #获取参考清单最大列数

if record_file:
    print(f"原始最大行数：{max_rows}")
    file_log.write(f"原始最大行数：{max_rows}"+"\n")
    print(f"原始最大列数：{max_columes}")
    file_log.write(f"原始最大列数：{max_columes}"+"\n")
    print(f"参考清单最大行数：{ref_max_rows}")
    file_log.write(f"参考清单最大行数：{ref_max_rows}"+"\n")
    print(f"参考清单最大列数：{ref_max_columes}")
    file_log.write(f"参考清单最大列数：{ref_max_columes}"+"\n")

compare_max_columns= max(max_columes,ref_max_columes)
compare_max_rows= max(max_rows,ref_max_rows)
print("最大列：",compare_max_columns)
print("最大行：",compare_max_rows)


diff_df=pd.DataFrame(columns=df.columns) #创建一个空表，记录元器件差异，列索引和读取的Excel一致
diff_df['diff refenrence']=[] #插入新列，记录不同的元器件位号

# print("diff_df插入列后初值: ",diff_df)      

########################################################
# 将两个 DataFrame 连接起来，并标记它们来自于哪个文件
df['file']='cmp_file' #添加一列，用于标识是哪个文件的数据
ref_df['file']='ref_file' #添加一列，用于标识是哪个文件的数据
df_concat = pd.concat([df, ref_df], ignore_index=True)

df_diff = df_concat.reset_index(drop=True) #重排索引，并去掉索引
ignore_columns = ['Item Number','file'] #去重复行时要忽略的列名

###########################################################
#去掉重复的行，比较时忽略ignore_columns列表里的列
df_diff = df_concat.drop_duplicates(subset=df_concat.columns.difference(ignore_columns), keep=False)
# print("df_diff \n",df_diff)

######################################################
# 根据某列数据相同的行将 DataFrame 拆分为不同的 DataFrame
grouped = df_diff.groupby('file')
# 将每个分组存储为不同的 DataFrame
grouped_dataframes = [group for _, group in grouped]
# 打印每个分组的 DataFrame
if len(grouped_dataframes[0])>=len(grouped_dataframes[1]): #比较行的多少
    # 下面查找差异时，df在外循序，因此要分配行数多的表
    df = grouped_dataframes[0].reset_index(drop = True) 
    ref_df = grouped_dataframes[1].reset_index(drop = True)
else:
    ref_df = grouped_dataframes[0].reset_index(drop = True) 
    df = grouped_dataframes[1].reset_index(drop = True)
# for i, group_df in enumerate(grouped_dataframes):
#     print(f"DataFrame {i+1}:\n{group_df}\n")
print(f"DataFrame {1}:\n{df}\n")
print(f"DataFrame {2}:\n{ref_df}\n")

max_rows = df.shape[0]  #获取最大行数
max_columes = df.shape[1]  #获取最大列数
ref_max_rows = ref_df.shape[0]  #获取参考清单最大行数
ref_max_columes = ref_df.shape[1]  #获取参考清单最大列数
print("max_rows\n",max_rows)
print("ref_max_rows\n",ref_max_rows)

###########################################################
#以下比较2个差异表里的具体单元格差异，主要是元器件位号的差异
empty_row = pd.Series([pd.NA, pd.NA, pd.NA, pd.NA, pd.NA, pd.NA], index=diff_df.columns) #空行，插入空行时使用
for rows in range(0,max_rows,1):  #比较表的行循环
    search_result = False
    for ref_rows in range(0,ref_max_rows,1):  #被比较表的行循环
         #此循环查找Value和PCB Footprint里相同的行,由于删除了"item"列,因此列号需要减一
        if df.iloc[rows,VALUE_COLUMN]== ref_df.iloc[ref_rows,VALUE_COLUMN] and \
           df.iloc[rows,FOOTPRINT_COLUMN]== ref_df.iloc[ref_rows,FOOTPRINT_COLUMN] :
            #如果型号相同，将本行数据保存到差异表里。注意需要to_frame()和.T转置，否则不是按行追加
            diff_df=pd.concat([diff_df,df.iloc[rows].to_frame().T],axis=0,ignore_index=False) 
            diff_df=pd.concat([diff_df,ref_df.iloc[ref_rows].to_frame().T],axis=0,ignore_index=False)      
            reference_num=df.iloc[rows,REFERENCE_COLUMN]
            reference_num_list=reference_num.split(',') #字符串分割为列表            
            ref_reference_num = ref_df.iloc[ref_rows,REFERENCE_COLUMN]
            ref_reference_num_list = ref_reference_num.split(',') #字符串分割为列表
            #######################################################################
            #利用集合找出差异的位号，
            reference_diff_list = list(set(reference_num_list)-set(ref_reference_num_list)) # 利用集合找出列表reference_num_list中独有的元素
            # reference_diff_list =','.join(list(reference_diff_list)) #将列表转换成以逗号分隔的字符串
            reference_diff_list= sorted(reference_diff_list,key=custom_sort_key) #排序，默认升序排序
            # print("reference_diff_list:\n",reference_diff_list)            
            reference_num_diff =','.join(list(reference_diff_list)) #将列表转换成以逗号分隔的字符串
            diff_df.iloc[(diff_df.shape[0]-2),diff_df.shape[1]-2]=reference_num_diff #差异的位号写在倒数第2行
            reference_diff_list = list(set(ref_reference_num_list)-set(reference_num_list)) # 利用集合找出列表ref_reference_num_list中独有的元素
            # reference_diff_list =','.join(list(reference_diff_list)) #将列表转换成以逗号分隔的字符串
            reference_diff_list= sorted(reference_diff_list,key=custom_sort_key) #排序，默认升序排序
            # print("reference_diff_list:\n",reference_diff_list)            
            reference_num_diff =','.join(list(reference_diff_list)) #将列表转换成以逗号分隔的字符串
            diff_df.iloc[(diff_df.shape[0]-1),diff_df.shape[1]-2]=reference_num_diff #差异的位号写在本行，也就是最后一行                   
            search_result = True
            break
        else:
            # print("差异的行号: ",rows)
            # print("差异的列号: ",VALUE_COLUMN-1)
            pass
    if search_result== False:
        diff_df=pd.concat([diff_df,df.iloc[rows].to_frame().T],axis=0,ignore_index=False)
        # diff_df=pd.concat([diff_df,diff_df.columns.to_frame().T],axis=0,ignore_index=False)
        diff_df=pd.concat([diff_df,empty_row.to_frame().T],axis=0,ignore_index=False)

 
# print("diff_df最终值\n",diff_df)
compare_result=False
###############################################
#以下用xlsxwriter设置Excel的格式和颜色
if compare_result==False:
    max_rows = diff_df.shape[0]  #获取最大行数
    max_columes = diff_df.shape[1]  #获取最大列数
    
    writer = pd.ExcelWriter(New_File_Name,engine='xlsxwriter') #使用ExcelWriter需要安装xlsxwriter模块：pip install xlsxwriter 
    diff_df.to_excel(writer, sheet_name='Sheet1', index=False)

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
        'text_wrap': True,  #自动换行
        'font_color': '#FF0000', #设置前景(字体)颜色为灰色
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

    red_format = workbook.add_format({
        'font_color': '#FF0000',
    })

# 遍历 DataFrame，并检查空行，给上一行设置红色背景颜色
    for row_num in range(1, max_rows):
        if diff_df.iloc[row_num].isnull().all(axis=0):
        #    print("空行：\n",row_num)
           worksheet.conditional_format(row_num,0,row_num,(max_columes-1), {'type':'no_blanks','format': red_format})

#   以下循序将Excel表格背景颜色隔行设置为灰色
    for row_num in range(0,max_rows+1,1):
        if row_num % 2 == 0:
            # worksheet.set_row(i,None,row_even_format)
            worksheet.conditional_format(row_num,0,row_num,(max_columes-1), {'type':'no_errors','format': gray_format})
        else:
            worksheet.conditional_format(row_num,0,row_num,(max_columes-1), {'type':'no_errors','format': while_format})
   
    worksheet.set_column("A:A", 10, header_format) #设置A列宽度为10，格式为:垂直中信对齐；水平中心对齐
    worksheet.set_column("B:B", 10, header_format1)
    worksheet.set_column("C:C", 50,header_format2)
    worksheet.set_column("D:D", 30,header_format3)
    worksheet.set_column("E:E", 25,header_format4) 
    worksheet.set_column("F:F", 25,header_format5)
    worksheet.set_column("G:G", 25,header_format4)

    format_border = workbook.add_format({'border':5})   # 设置边框格式
    worksheet.conditional_format('A1:XFD1048576',{'type':'no_blanks', 'format': format_border}) #整个工作表，根据条件来设置格式
    # writer.save() #save()方法已经弃用？使用close()方法既是保存退出。

    worksheet.freeze_panes(1,1)   # 冻结首行
    worksheet.autofilter(0,0,max_rows,(max_columes-1))   # 添加筛选

    writer.close()  #保存\退出
if record_file:
    file_log.write("End.")
    file_log.close()  #关闭日志文件

