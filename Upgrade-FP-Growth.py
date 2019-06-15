# -*- coding: utf-8 -*-
"""
Created on Tue April 9 21:06:37 2019
@author: liyun
"""
"""
导入数据包
"""
import xlwt#xlwt用于写入excel
import xlrd#xlrd用于读excel
import pyfpgrowth#pyfpgrowth.py必须和Upgrade-FP-Growth在同一路径下。
"""
从excel读取数据
"""
wb = xlrd.open_workbook(r'C:\Users\liyun\Desktop\毕设\数据\programming\New Data.xls')
#自己定义路径，没有原excel也没关系。它会自动新建一个文件。但是只能生成xls格式。
sheet1 = wb.sheet_by_index(0)
#设定一个sheet，名字叫sheet1
originalrow=sheet1.nrows
#读取excel原始最大行数
"""
excel数据导入空列表
"""
agvnum=[]
#新建空列表 for agv number
armgnum=[]
#新建空列表 for armg number
task=[]
#新建空列表 for task number
policy=[]
#新建空列表 for 调度规则
key=[]
#新建空列表 for 0/1 
indicator_one=[]
#新建空列表 for the first indicator

for rows in range(1,originalrow):#excel列第一个为0
    agvnum.append(int(sheet1.cell(rows,1).value*1000))
    armgnum.append(int(sheet1.cell(rows,2).value*10000))
    task.append(int(sheet1.cell(rows,3).value))
    policy.append(int(sheet1.cell(rows,4).value))
    key.append(int(sheet1.cell(rows,5).value+1)*100)
    indicator_one.append(round(float(sheet1.cell(rows,16).value)*10))
#数值标记法，详见论文。
"""
合并生成“输入项”and“输出项”。并保存到excel方便查看。
"""
"""
设置excel表头
"""
book = xlwt.Workbook() 
#定义写入excel
sheet = book.add_sheet('test', cell_overwrite_ok=True) 
#定义写入的sheet，以及是否能覆盖原数据
colume_name=['agv数量','armg数量','task数量','指派策略','key','indicator 1 list']
row = 0
for item in range(len(colume_name)):
    sheet.write(row, item, colume_name[item])
#写入第0行的列名称
"""
合并并写入数据
"""
trans_combine=[]
#新建空列表 for 输入项
trans_indi_one=[]
#新建空列表 for 输出项
i=0
location=0
row=0
for rowNum in range(1,originalrow): 
    if [agvnum[rowNum-1],armgnum[rowNum-1],task[rowNum-1],policy[rowNum-1],key[rowNum-1]] not in trans_combine:
        trans_combine.append([agvnum[rowNum-1],armgnum[rowNum-1],task[rowNum-1],policy[rowNum-1],key[rowNum-1]])
        trans_indi_one.insert(i,[indicator_one[rowNum-1]])
        i=i+1
    else:
        location=trans_combine.index([agvnum[rowNum-1],armgnum[rowNum-1],task[rowNum-1],policy[rowNum-1],key[rowNum-1]])
        trans_indi_one[location].append(indicator_one[rowNum-1])
#合并输入项 & 输出项
for row in range(1,i+1):
    sheet.write(row,0,trans_combine[row-1][0])#list的第一个序号为0
    sheet.write(row,1,trans_combine[row-1][1])
    sheet.write(row,2,trans_combine[row-1][2])
    sheet.write(row,3,trans_combine[row-1][3])
    sheet.write(row,4,trans_combine[row-1][4])
    sheet.write(row,5,str(trans_indi_one[row-1]))
#将合并前的数据写入excel
book.save(r'C:\Users\liyun\Desktop\毕设\数据\programming\test.xls')
#保存excel。没有原文件没关系，会自动生成一个对应名字的xls文件。但是xlwt包只能保存xls格式表格。

"""
Upgrade-FP-Growth
"""
"""
调参transactions/support_threshold/confidence_threshold/second_support_threshold/second_confidence_threshold
"""
transactions=trans_indi_one
#transaction为要挖掘的“输出项”列表
support_threshold=1
#挖掘输出项的支持度参数
confidence_threshold=1
#挖掘输出项的置信参数
patterns = pyfpgrowth.find_frequent_patterns(transactions,support_threshold)
#调用pyfpgrowth.py的 find_frequent_patterns函数挖掘频繁模式。
#可通过：print(patterns),查看挖掘结果
rules = pyfpgrowth.generate_association_rules(patterns,confidence_threshold)
#调用pyfpgrowth.py的 generate_association_rules函数挖掘关联规则。
#可通过：print(rules),查看挖掘结果
print("一次关联规则分析:",rules)
print("一次关联规则结果数量:",len(rules))
#输出第一次挖掘的关联规则以及关联规则数量
for i in rules.keys():
    if len(rules[i][0])==0:
        pass
    else:
        list_i=[]
        trans_output=[]
        for t in range(0,len(i)):
            list_i.append(i[t])
        for transaction_items in transactions:
            Intersection = [k for k in i if k in transaction_items]
            if Intersection==list_i:
                trans_output.append(trans_combine[transactions.index(transaction_items)])
#遍历关联规则，找到对应的所有“输入项”数据
        """
        二次反向挖掘
        """
        second_support_threshold = 20
        #挖掘输入项的支持度参数
        second_confidence_threshold = 1
        #挖掘输入项的置信参数
        input_patterns = pyfpgrowth.find_frequent_patterns(trans_output, second_support_threshold)
        #调用pyfpgrowth.py的 find_frequent_patterns函数挖掘频繁模式。
        #可通过：print(input_patterns),查看挖掘结果
        input_rules = pyfpgrowth.generate_association_rules(input_patterns, second_confidence_threshold)
        #调用pyfpgrowth.py的 generate_association_rules函数挖掘关联规则。
        #可通过：print(input_rules),查看挖掘结果
        print("二次关联规则分析:",input_rules)    
        print("二次关联规则结果数量:",len(input_rules))
        #输出第一次挖掘的关联规则以及关联规则数量
        for n in input_rules.keys():
            list_n=[]
            trans_input=[]
            if len(input_rules[n][0])==0:
                if input_rules[n][0]==1:
                    print("输出数据",i,"和",rules[i][0],"与输入数据",n,"有强关联")
                else:
                    pass
            else:
                medium_list=[]
                medium_intersection=[]
                for m in range(0,len(n)):
                    medium_list=[n[m],input_rules[n][0][0]]
                    medium_intersection = [k for k in medium_list if k in list_n]
                    if medium_intersection == medium_list:
                        pass
                    else:
                        list_n.append(medium_list)
                        if input_rules[n][1] == 1:
                            print("输出数据",i,"和",rules[i][0],"与输入数据",medium_list,"有强关联")
                        else:
                            pass
            #筛选无用的规则，并显示剩余数据。
                  


