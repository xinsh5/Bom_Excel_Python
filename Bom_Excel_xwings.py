#!/usr/bin/env python
# !encoding:utf-8

import time
import xlwings as xw
current_time_struct = time.localtime()  #获取当前时间
# 分别获取当前年、月、日、时、分、秒
current_year = current_time_struct.tm_year
current_month = current_time_struct.tm_mon
current_day = current_time_struct.tm_mday
current_hour = current_time_struct.tm_hour
current_minute = current_time_struct.tm_min
current_second = current_time_struct.tm_sec
print(current_time_struct)  #输出当前时间
app=xw.App(visible=False,add_book=False)
app.display_alerts=False               #关闭各种提示信息，可以提高运行速度
File_Name='hongyun_V01导出清单_20240129_0130.xlsx'  #原始文件名
New_File_Name='hongyun_V01清单.xlsx'  #输出的文件名
File_Log_Name = 'BOM_Excel.log'  #记录的日志文件
file_log = open(File_Log_Name,'w')  #打开记录文件
file_log.write("时间："+f"{current_year}"+"年"+ f"{current_month}"+"月"+f"{current_day}"+"日"
               +f"{current_hour}"+"时"+f"{current_minute}"+"分"+f"{current_second}"+"秒"+"\n")
file_log.write("读出的文件："+File_Name+"\n")
file_log.write("写入的文件："+New_File_Name+"\n")

def compare_func(item):
    # 提取每个元素后面的数值部分并转换成int类型
    num = int(''.join([char for char in item if char.isdigit()]))
    return num

#打开要处理的Excel文件名
Work_Book=app.books.open(File_Name)      
#打开要处理的Excel文件中的工作簿
Work_Sheet=Work_Book.sheets[0]      

max_row=Work_Sheet.used_range.shape[0]   #获取最大行数
max_colume=Work_Sheet.used_range.shape[1]  #获取最大列数
part_count=0  #器件数量
initial_line_num = 2  #初始行号
# X=3
# Y=4
# Z=5
# reference_num_list=Work_Sheet.range((Y,X)).value
print(f"原始最大行数：{max_row}")
file_log.write(f"原始最大行数：{max_row}"+"\n")
print(f"原始最大列数：{max_colume}")
file_log.write(f"原始最大列数：{max_colume}"+"\n")
# print(f"单元格内容:{reference_num_list}")
# Work_Sheet.range(f'{Z}:{Z}').delete()
first_loop_end_flag = False  #首次循环结束标志
for i in range(initial_line_num,max_row,1):
    repetition_flag = False  #重复标志
    if i > initial_line_num:
        first_loop_end_flag = True
    
    part_count=Work_Sheet.range(i,2).value
    reference_num_list=Work_Sheet.range(i,3).value
    for j in range(i+1,max_row,1):
        if Work_Sheet.range(j,9).value == None:  #如果型号单元格为空则跳过
            if first_loop_end_flag == False:
                print(f"第{j}行没有型号")
                file_log.write(f"单元格：{i},{j} 没有型号"+"\n")                
        elif Work_Sheet.range(i,9).value == Work_Sheet.range(j,9).value:  #如果型号相同则合并相同的行
                reference_num_list=reference_num_list+','+Work_Sheet.range(j,3).value  #合并位号单元格内容
                part_count=part_count+Work_Sheet.range(j,2).value  #元器件数量相加
                Work_Sheet.range(f'{j}:{j}').delete()  #删除后面相同的行
                max_row=max_row-1  #行总是减一
                print(f"重复的行号:{i},{j}")
                file_log.write(f"重复的行号:{i},{j}"+"\n")
                repetition_flag = True
        
    if(repetition_flag):
        print(f"重复单元格的内容:{reference_num_list}")
        file_log.write(f"重复单元格的内容:\n{reference_num_list}"+"\n")
        reference_num_list=sorted(reference_num_list.split(','),key=compare_func)  #排序，默认升序排序
        print(f"合并单元格并排序完的内容:{reference_num_list}")
        file_log.write(f"合并单元格并排序完的内容:\n{reference_num_list}"+"\n")
        reference_num_list=','.join(reference_num_list)  #将列表转换成以逗号分隔的字符串
        Work_Sheet.range(i,3).value=reference_num_list
        Work_Sheet.range(i,2).value=part_count
for i in range(initial_line_num,max_row+1,1):
    Work_Sheet.range(i,1).value=i-1


#reference_num_list=Work_Sheet.range('C8').value+','+Work_Sheet.range('C9').value
#print(f"C9单元格内容:{reference_num_list}")
#reference_num_list=sorted(reference_num_list.split(','),key=compare_func)
#print(f"C9单元格排序内容:{reference_num_list}")
#reference_num_list=','.join(reference_num_list)#将列表转换成以逗号分隔的字符串
#Work_Sheet.range('C9').value=reference_num_list           
#Work_Sheet.range('B9').value=Work_Sheet.range('B8').value+Work_Sheet.range('B9').value

Work_Book.save(New_File_Name)
#保存改动的工作簿。若无保存，则上述操作会随着工作簿的关闭而作废不保存。
Work_Book.close()
#关闭工作簿。
app.quit()
#退出Office软件，不驻留后台。
file_log.write("End.")
file_log.close()  #关闭日志文件

