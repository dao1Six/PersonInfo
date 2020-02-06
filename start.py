import xlwt,xlrd
from xlutils.copy import copy
import GenerateName
import GeneratePhone
import GenerateID
from excelcontrol import ExcelControl
import threading
import time

class PoetrySpider:
    # 写入excel

    def generateValue(self):
        #生成数据
        zidian = {}
        for i in range(18):
            value = []
            #生成姓名，拼音
            o = GenerateName.aa()
            name = o[0]
            value.append(name)
            py = o[1]
            value.append(py)
            #生成手机号
            p = GeneratePhone.create_phone()
            value.append(p)
            #生成身份证
            s = GenerateID.genneratorID()
            value.append(s)
            zidian[i] = value
        return zidian

    # 装饰器，计算插入10000条数据需要的时间
    def timer(func):
        def decor(*args):
            start_time = time.time()
            func(*args)
            end_time = time.time()
            d_time = end_time - start_time
            print("这次插入数据耗时 : ", d_time)

        return decor

    @timer
    def wirt(self,path):
        book_name_xlsx = path
        data = xlrd.open_workbook(book_name_xlsx, formatting_info=True)
        excel = copy(wb=data)  # 完成xlrd对象向xlwt对象转换
        excel_table = excel.get_sheet(0)  # 获得要操作的页
        table = data.sheets()[0]
        nrows = table.nrows  # 获得行数
        print("从第"+str(nrows)+"行开始写入")
        ncols = table.ncols  # 获得列数
        for i in range(60000):
            value = []
            #生成姓名，拼音
            o = GenerateName.aa()
            name = o[0]
            value.append(name)
            py = o[1]
            value.append(py)
            #生成手机号
            p = GeneratePhone.create_phone()
            value.append(p)
            #生成身份证
            s = GenerateID.genneratorID()
            value.append(s)

            #写姓名
            excel_table.write(nrows, 0, value[0])
            #写账号
            excel_table.write(nrows, 1, value[1])
            #显示名
            excel_table.write(nrows, 3, value[0])
            #手机号码
            excel_table.write(nrows, 7, value[2])
            #证件号码
            excel_table.write(nrows, 11, value[3])
            nrows = nrows+1
            print("正在写入第"+str(nrows)+"行")

        excel.save(path)

    def main(self):

        c_url03 = threading.Thread(target=self.wirt,args=('人员3.xls',))
        c_url04 = threading.Thread(target=self.wirt,args=('人员4.xls',))

        # 启动所有线程

        c_url03.start()
        c_url04.start()




if __name__ == "__main__":
    flt = PoetrySpider()
    flt.main()







