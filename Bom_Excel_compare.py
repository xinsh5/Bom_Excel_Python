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

File_Name='./Bom/hongyun导出清单_20240213.xlsx'  #要比较文件名
Ref_File_Name = "./Bom/hongyun_V01导出清单_20240126.xls" #被比较的清单
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
print("最大行：",compare_max_columns)
#以下循环逐个比较2个dataframe的单元格差异；

diff_df=df.iloc[0:0,0:compare_max_columns-1] #赋值给差异表首行，作为列索引
diff_df=diff_df.rename(columns={'Item Number':'diff refenrence'}) #重命名首列的名称
# diff_refenrence= pd.Series(['diff refenrence'])
# diff_df=pd.concat([diff_df,diff_refenrence.to_frame().T],axis=1,ignore_index=False)
print("diff_df初值: ",diff_df)      

compare_result = True
# compare_max_rows= 2
df.drop('Item Number',axis=1,inplace=True)  #删除Item Number列,
ref_df.drop('Item Number',axis=1,inplace=True)  #删除Item Number列,
df_same_list=[]
ref_df_same_list=[]
# pair_s=pd.Series()
for rows in range(0,max_rows-1,1):  #比较表的行循环
    tmp_df_row= df.iloc[rows]
    search_result = False
    for ref_rows in range(0,ref_max_rows-1,1):  #被比较表的行循环
        tmp_ref_df_row= ref_df.iloc[ref_rows]
        if tmp_df_row.compare(tmp_ref_df_row).empty:
            #"=="不能判断空值，因此需要单独判断空值
            print(f"相同的行：{rows}和{ref_rows}\n")
            df_same_list.append(rows) 
            ref_df_same_list.append(ref_rows)           
            search_result = True
            break

print("df_same_list",df_same_list)
print("ref_df_same_list",ref_df_same_list)
#将不同的行号提取出来，存储在df_diff_list列表里
df_complete_list=[]
for i in range(max_rows): 
    df_complete_list.append(i)
df_diff_list = list(set(df_complete_list)-set(df_same_list))
df_diff_list.sort() #需要排序

#将不同的行号提取出来，存储在ref_df_diff_list列表里
ref_df_complete_list=[]
for i in range(ref_max_rows): 
    ref_df_complete_list.append(i)
ref_df_diff_list = list(set(ref_df_complete_list)-set(ref_df_same_list))
ref_df_diff_list.sort()
print("df_diff_list",df_diff_list)
print("ref_df_diff_list",ref_df_diff_list)


for i in df_same_list:
    df.drop(index=i,axis=0,inplace=True)  #删除相同的行,
df.reset_index(drop=True, inplace=True)  # 重排索引并更新
for i in ref_df_same_list:
    ref_df.drop(index=i,axis=0,inplace=True)  #删除相同的行,
ref_df.reset_index(drop=True, inplace=True)  # 重排索引并更新

print("df不同行\n",df)
print("ref_df不同行\n",ref_df)

max_rows = df.shape[0]  #获取最大行数
max_columes = df.shape[1]  #获取最大列数
ref_max_rows = ref_df.shape[0]  #获取参考清单最大行数
ref_max_columes = ref_df.shape[1]  #获取参考清单最大列数
print("max_rows\n",max_rows)
print("ref_max_rows\n",ref_max_rows)

# compare_result = True
# # compare_max_rows= 2
for rows in range(0,max_rows-1,1):  #比较表的行循环
    search_result = False
    for ref_rows in range(0,ref_max_rows-1,1):  #被比较表的行循环
         #此循环查找Value和PCB Footprint里相同的行,由于删除了"item"列,因此列号需要减一
        if df.iloc[rows,VALUE_COLUMN-1]== ref_df.iloc[ref_rows,VALUE_COLUMN-1] and \
           df.iloc[rows,FOOTPRINT_COLUMN-1]== ref_df.iloc[ref_rows,FOOTPRINT_COLUMN-1] :
            #如果型号相同，将本行数据保存到差异表里。注意需要to_frame()和.T转置，否则不是按行追加
            diff_df=pd.concat([diff_df,df.iloc[rows].to_frame().T],axis=0,ignore_index=False) 
            diff_df=pd.concat([diff_df,ref_df.iloc[ref_rows].to_frame().T],axis=0,ignore_index=False)      
            reference_num=df.iloc[rows,REFERENCE_COLUMN-1]
            reference_num_list=reference_num.split(',') #字符串分割为列表            
            ref_reference_num = ref_df.iloc[ref_rows,REFERENCE_COLUMN-1]
            ref_reference_num_list = ref_reference_num.split(',') #字符串分割为列表
            if len(reference_num_list)>=len(ref_reference_num_list):
                reference_diff_list = list(set(reference_num_list)-set(ref_reference_num_list))
                print("reference_diff_list:\n",reference_diff_list)            
                reference_num_diff =','.join(list(reference_diff_list)) #将列表转换成以逗号分隔的字符串
                print("reference_num_diff:\n",reference_num_diff)
                diff_df.iloc[(diff_df.shape[0]-2),0]=reference_num_diff #差异的位号写在多出来的那一行
            else:
                reference_diff_list = list(set(ref_reference_num_list)-set(reference_num_list))
                print("reference_diff_list:\n",reference_diff_list)            
                reference_num_diff =','.join(list(reference_diff_list)) #将列表转换成以逗号分隔的字符串
                print("reference_num_diff:\n",reference_num_diff)
                diff_df.iloc[(diff_df.shape[0]-1),0]=reference_num_diff  #差异的位号写在多出来的那一行
            search_result = True
            break
        else:
            # print("差异的行号: ",rows)
            # print("差异的列号: ",VALUE_COLUMN-1)
            pass
    if search_result== False:
        diff_df=pd.concat([diff_df,df.iloc[rows].to_frame().T],axis=0,ignore_index=False)
        diff_df=pd.concat([diff_df,diff_df.columns.to_frame().T],axis=0,ignore_index=False)
 
print("diff_df最终值\n",diff_df)
compare_result=False
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
        'text_wrap': True,  #自动换行
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
        
    worksheet.set_column("A:A", 40, header_format) #设置A列宽度为10，格式为:垂直中信对齐；水平中心对齐
    worksheet.set_column("B:B", 10, header_format1)
    worksheet.set_column("C:C", 50,header_format2)
    worksheet.set_column("D:D", 30,header_format3)
    worksheet.set_column("E:I", 25,header_format4) 

    format_border = workbook.add_format({'border':5})   # 设置边框格式
    worksheet.conditional_format('A1:XFD1048576',{'type':'no_blanks', 'format': format_border}) #整个工作表，根据条件来设置格式
    # writer.save() #save()方法已经弃用？使用close()方法既是保存退出。

    worksheet.freeze_panes(1,1)   # 冻结首行
    worksheet.autofilter(0,0,max_rows,(max_columes-1))   # 添加筛选

    writer.close()  #保存\退出
if record_file:
    file_log.write("End.")
    file_log.close()  #关闭日志文件

