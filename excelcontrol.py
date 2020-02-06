import os

import openpyxl
import xlrd,xlwt,xlutils
from xlutils.copy import copy

class ExcelControl():
    """"""

    #
    def OpenXlxs(self,xlsxName,sheet_name):
        readOpenXlsx = xlrd.open_workbook (xlsxName)

        readXlsxSheet = readOpenXlsx.sheet_by_name (sheet_name)
        # copy管道作用
        writeOpenXlsx = copy (readOpenXlsx)

        return readXlsxSheet, writeOpenXlsx, xlsxName


    #获取文件某列数据
    def get_sheet_col_info(self,xlsxName,sheet_name,col_number):
        readOpenXlsx = xlrd.open_workbook (xlsxName)
        readXlsxSheet = readOpenXlsx.sheet_by_name(sheet_name)
        # 获取行数
        rowMax = readXlsxSheet.nrows
        # 获取列数
        colMax = readXlsxSheet.ncols
        result = []
        for r in range (0,rowMax):
            alldata = readXlsxSheet.row_values (r)  # 循环输出excel表中每一行，即所有数据
            cloValue = alldata[col_number]  # 取出表中某列数据
            result.append(cloValue)
        return result

    #给文件某列写入数据
    def write_info_into_row(self,filename,datalist,clo_number):
        book = xlrd.open_workbook (filename)  # 打开excel
        work_book = copy (book)  # 复制excel
        sheet = work_book.get_sheet(0)  # 获取表格的数据
        rowMax = len(datalist)
        for r in range(0,rowMax):
            sheet.write (r, clo_number,datalist[r])  # 修改行列的数据
            work_book.save (filename)  # 保存excel


    #获取两列数据生成字典
    def excle_generate_dict(self,xlsxName,sheet_name,keycol_number,col_number):
        readOpenXlsx = xlrd.open_workbook (xlsxName)
        readXlsxSheet = readOpenXlsx.sheet_by_name(sheet_name)
        # 获取行数
        rowMax = readXlsxSheet.nrows
        # 获取列数
        colMax = readXlsxSheet.ncols
        dic = {}
        for r in range (0,rowMax):
            alldata = readXlsxSheet.row_values (r)  # 循环输出excel表中每一行，即所有数据
            clokeyValue = alldata[keycol_number]  # 取出表中某列数据
            cloValue = alldata[col_number]  # 取出表中某列数据
            dic[clokeyValue] = cloValue
        return dic

    def write_excel_xlsx(path, sheet_name, value):
        index = len(value)
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = sheet_name
        for i in range(0, index):
            for j in range(0, len(value[i])):
                sheet.cell(row=i + 1, column=j + 1, value=str(value[i][j]))
        workbook.save(path)
        print("xlsx格式表格写入数据成功！")


    def write_excel_xls_append(path, value):
        index = len(value)  # 获取需要写入数据的行数
        workbook = xlrd.open_workbook(path)  # 打开工作簿
        sheets = workbook.sheet_names()  # 获取工作簿中的所有表格
        worksheet = workbook.sheet_by_name(sheets[0])  # 获取工作簿中所有表格中的的第一个表格
        rows_old = worksheet.nrows  # 获取表格中已存在的数据的行数
        new_workbook = copy(workbook)  # 将xlrd对象拷贝转化为xlwt对象
        new_worksheet = new_workbook.get_sheet(0)  # 获取转化后工作簿中的第一个表格
        for i in range(0, index):
            for j in range(0, len(value[i])):
                new_worksheet.write(i+rows_old, j, value[i][j])  # 追加写入数据，注意是从i+rows_old行开始写入
        new_workbook.save(path)  # 保存工作簿
        print("xls格式表格【追加】写入数据成功！")


if __name__ == '__main__':
    filepath = "./banche.xlsx"
    sheet_name = "banche"
    e = ExcelControl()
    # result = e.get_sheet_col_info(filepath,sheet_name,0)
    # print(result)
    # e.write_info_into_row(filepath,datalist=result,clo_number=5)
    dic = e.excle_generate_dict(filepath,sheet_name,0,1)
    print(type(dic))
    for i in dic.items():
        print(i)




