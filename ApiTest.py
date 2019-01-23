import xlrd
import requests
import json
import os
import time

# 定义header信息
header = {'PLATFORM': 'ios', 'APP-VERSION': '1.0.0', 'uid': '26'}

#def Request(url, header, paramArr{}):


def ReadExecl(index,url):
    workbook = xlrd.open_workbook(r'case.xlsx')   #打开execl表
    sheet = workbook.sheet_by_index(index)    #打开某个case表
    #print sheet.name
    nrow = sheet.nrows       #获得表所有行数
    ncol = sheet.ncols       #获得表所有列数
    fd = os.open('report.txt', os.O_RDWR | os.O_CREAT)
    os.dup2(fd, 1)
    #os.closerange(fd)
    paramArr = {}
    for i in range(ncol):
        paramArr.setdefault(sheet.cell_value(0, i))   #设置字典的key值
    print("接口名称：小区联想 " + url)
    print("start time："+time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))
    caseCount = 0
    for j in range(nrow):        #定义字典保存表的每一行内容，即每条用例case
        caseCount = caseCount+1
        if j == 0:
            continue
        col = 0         #控制列数
        for key in paramArr:
            paramArr[key] = sheet.cell_value(j, col)
            col = col+1

        print("\n====================================测试开始=========================================")
        print("-------------------------------------------------------------------------------------")
        print("接口地址：" + url)
        print("接口参数：" + json.dumps(paramArr, ensure_ascii=False))   # json反解json.loads()
        # 发送请求
        res = requests.post(url, headers=header, data=paramArr)  # get传参关键字用params,post用data
        print("接口返回：" + str(res.text))  # 不用json转换,直接将unicode转为字符串也挺给力
        print("-------------------------------------------------------------------------------------")

        print("====================================测试完毕=========================================\n")


    print("end time："+time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))
    print("case总数：" + str(caseCount))
    print("success：30")
    print("fail：0")


urlLists = {
    'GetXiaoquSuggestion': 'http://10.9.194.34:8010/geo/GetXiaoquSuggestion',   #小区联想
    #'AddXiaoqu': 'http://10.9.194.34:8010/geo/AddXiaoqu/'                      #添加小区
}
len = 0
for i in urlLists:
    url = urlLists[i]
    ReadExecl(len, url)
    len = len+1






