import psycopg2

conn = psycopg2.connect(
    host="localhost",
    database="material",
    user="postgres",
    password="aman")

cur = conn.cursor()

cur.execute('SELECT version()')

for i in range(1000):

    cur.execute("INSERT INTO facture.aman (name, id) VALUES ('{a}', 'b');")
    conn.commit()

cur.close()
conn.close()

# dbv = cur.fetchone()
# print(dbv)



# import threading
# import time
# import openpyxl
# from openpyxl import load_workbook, Workbook
# from pathlib import Path
# import re
# import os
# from multiprocessing import Pool
# import pandas as pd
# import xlwings as xw


# import win32com
# print(win32com.__gen_path__)

# unsorted_list=[1, 5]
# # print(unsorted_list)
# sorted_list=unsorted_list.sort()
# # print(unsorted_list)

# def decide_if_queue(number_list):
#     """
#         number_list: it must be list of numbers only
#     """
#     global is_queue
#     is_queue=True
#     old_number=0
#     number_list.sort()
#     for number in number_list:
#         difference=number-old_number
#         if not old_number==0 and not difference==1:
#             is_queue=False
        
#         old_number=number
    

# decide_if_queue(number_list=unsorted_list)
# print(is_queue)

# # file_path=Path.cwd() / 'test material table11111.xlsm'

# # wb=xw.Book(str(file_path))
# # sheet=wb.sheets['ALL FACTURES']

# # xlist=[[1,2,3,4,5,6,7,8,9], [1,2,3,4,5,6,7,8,9], [1,2,3,4,5,6,7,8,9], [1,2,3,4,5,6,7,8,9], [1,2,3,4,5,6,7,8,9], [1,2,3,4,5,6,7,8,9], [1,2,3,4,5,6,7,8,9], [1,2,3,4,5,6,7,8,9]]
# # for i in range(1,10):
# #     for j in range(1,10):
# #         sheet.range(i,j).value=1

# # sheet.range(15,1).value=xlist




# # # print(len("экономия3r4".encode("ascii", "ignore")) > len ("экономия3r4"))

# # xx='экономия3r4'
# # tm1='aman'
# # yy=xx.encode()
# # tm2=tm1.encode()
# # print(tm2)

# # class Test:
# #     def __init__(self):
# #         self.dict={
# #             1: self.func1,
# #             2: self.func2
# #         }

# #     def func1(self, msg):
# #         print('func1 is on: ', msg)

# #     def func2(self, msg):
# #         print('func2 is on: ', msg)

# # if __name__=="__main__":
# #     test=Test()
# #     dict[1]("hey")
# #     test.dict[1]("serh")
# #     test.dict[1]("ryj")

# # # xpath=Path(r'D:\BYTK_Facturation\7. MT\1-FACTURE')
# # xpath=Path(r'D:\BYTK_Facturation\7. MT\xxx')
# # wb=Workbook()
# # wb.create_sheet(title='test')
# # sh=wb.active
# # sh.cell(row=1, column=5, value='Expected last column \'Prix Total\'')
# # # sh.cell(row=1, column=11, value='Expected facture no')
# # # sh.cell(row=1, column=12, value='Expected to be empty')
# # sh.cell(row=1, column=12, value='Description')
# # sh.cell(row=1, column=15, value='total sum')
# # sh.cell(row=1, column=16, value='total row no')

# # for i, xl in enumerate(xpath.glob('*.xls*'), start=1):
# #     print (xl.name)
# #     sh.cell(row=i, column=1, value=xl.name)
# #     try:
# #         opxl=load_workbook(xl)
# #         xsh=opxl['Annexe']
# #         value1=xsh.cell(row=1, column=19).value
# #         value2=xsh.cell(row=1, column=20).value
# #         value3=xsh.cell(row=1, column=21).value
# #         value4=xsh.cell(row=1, column=22).value
# #         value5=xsh.cell(row=1, column=23).value
# #         value6=xsh.cell(row=1, column=24).value
# #         value7=xsh.cell(row=1, column=25).value
# #         value8=xsh.cell(row=1, column=26).value

# #         # for description columns
# #         value9=xsh.cell(row=1, column=12).value
# #         value10=xsh.cell(row=1, column=13).value
# #         value11=xsh.cell(row=1, column=14).value
# #         value12=xsh.cell(row=1, column=15).value
# #         value13=xsh.cell(row=1, column=16).value

# #         last_total_sum_cell_value=''
# #         last_total_sum_cell_color=''
# #         last_total_word_cell_value=xsh.cell(row=len(xsh['N']) , column=14)
# #         last_total_word_cell_color=xsh.cell(row=len(xsh['N']) , column=14).fill.start_color.index
# #         # desc_value_list=[value9, value10, value11, value12, value13]
# #         total_sum_list=[value3,value4,value5]
# #         # for i , xval in enumerate(desc_value_list, start=1):
# #         #     if str(xval).lower().__contains__('gnation'):
# #         #         last_total_word_cell_value=xsh.cell(row=len(xsh['N']) , column=14)
# #         #         last_total_word_cell_color=xsh.cell(row=len(xsh['N']) , column=14).fill.start_color.index

# #         for i, xval in enumerate(total_sum_list, start=1):
# #             if str(xval).lower().__contains__('prix'):
# #                 last_total_sum_cell_value=xsh.cell(row=len(xsh['V']), column=20+i).value
# #                 # last_total_sum_cell_value=xsh.cell(row=len(xsh['V']), column=20+i).fill.start_color.index



# #         sh.cell(row=i, column=2, value=value1)
# #         sh.cell(row=i, column=3, value=value2)
# #         sh.cell(row=i, column=4, value=value3)
# #         sh.cell(row=i, column=5, value=value4)
# #         sh.cell(row=i, column=6, value=value5)
# #         sh.cell(row=i, column=7, value=value6)
# #         sh.cell(row=i, column=8, value=value7)
# #         sh.cell(row=i, column=9, value=value8)

# #         # for description columns
# #         sh.cell(row=i, column=10, value=value9)
# #         sh.cell(row=i, column=11, value=value10)
# #         sh.cell(row=i, column=12, value=value11)
# #         # sh.cell(row=i, column=12).fill=last_total_sum_cell_value
# #         sh.cell(row=i, column=13, value=value12)
# #         sh.cell(row=i, column=14, value=value13)
# #         sh.cell(row=i, column=15, value=last_total_sum_cell_value)
# #         sh.cell(row=i, column=16, value=len(xsh['V']))

        


# #         # ysh=opxl['Facture']
# #         # value11=ysh.cell(row=2, column=6).value
# #         # value12=ysh.cell(row=2, column=7).value
        
# #         # sh.cell(row=i, column=11, value=value11)
# #         # sh.cell(row=i, column=12, value=value12)

# #     except:
# #         pass
    
# #     # if i>10:
# #     #     break
# #     # try:
# #     #     pxl=pd.ExcelFile(xl)
# #     #     sh.cell(row=i, column=2, value=xl.name)
# #     #     for j, psh in enumerate(pxl.sheet_names, start=1):
            
# #     #         sh.cell(row=i, column=2+j, value=psh)
# #     # except:
# #     #     sh.cell(row=i, column=2, value=xl.name)      

# # wb.save(Path.cwd() / 'all facture names.xlsx')



# # # iter_list_for_pooling=list(enumerate(xpath.glob('*.xls*'), start=1))
# # # # print(iter_list_for_pooling)
# # # p=Pool(3)
# # # result_of_pool_as_list=p.map(poolthis, iter_list_for_pooling)
# # # p.close()
# # # p.join()

# # # for i, xl in enumerate(xpath.glob('*.xls*'), start=1):
# # #     print(i, xl)
# #     # poolthis(xl=xl)
    


# # # for s in wb.worksheets:
# # #     print(wb.)

# # # wb.save(Path.cwd() / 'all facture names.xlsx')

# # # print(os.path.basename(wb))




# # # def poolthis(*args):
# # #     print(args)
# # #     try:
# # #         temp_wb=load_workbook(args[1])
# # #         sh.cell(row=args[0], column=2, value=args[1].name)
# # #         sh.cell(row=args[0], column=3, value='Success')
# # #         for j, shtemp in enumerate(temp_wb.worksheets, start=1):
# # #             sh.cell(row=args[0], column=3+j, value=shtemp.title)
# # #     except:
# # #         sh.cell(row=args[0], column=2, value=args[1].name)
# # #         sh.cell(row=args[0], column=3, value='Error')