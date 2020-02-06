import requests

import queue
import threading
import csv
import time

class PoetrySpider:

    def run(self):
        for i in range(10):
            print(threading.current_thread().name+"处理"+str(i))
            time.sleep(500)

    # 主函数
    def main(self):

        # 创建4个线程解析页面
        c_url01 = threading.Thread(target=self.run)
        c_url02 = threading.Thread(target=self.run)
        c_url03 = threading.Thread(target=self.run)
        c_url04 = threading.Thread(target=self.run)

        # 启动所有线程

        c_url01.start()
        c_url02.start()
        c_url03.start()
        c_url04.start()



if __name__ == "__main__":
    flt = PoetrySpider()
    flt.main()

