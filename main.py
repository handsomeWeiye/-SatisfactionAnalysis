import os
import pandas as pd
import time
from pyecharts.charts import Bar
from pyecharts.globals import ThemeType
from pyecharts import options as opts
from collections import Counter
from pyecharts.components import Table
from pyecharts.options import ComponentTitleOpts

excel_dir = r"岩蜂蜜满意度调查.xlsx"


def log(str):
    # 日志函数
    time_stamp = time.time()
    local_time = time.localtime(time_stamp)
    str_time = time.strftime('%Y-%m-%d %H:%M:%S', local_time)
    print(str_time)
    with open('log.txt', 'a+') as f:
        logInfo = str_time + "   " + str
        print(logInfo)
        f.write(logInfo + "\n")


def getInfoList(cols):
    # 数据函数，把excel列表中的数据转化为了一个列表
    provider_file = pd.read_excel(excel_dir, usecols=cols)
    provider_list = provider_file.values
    print("目前的读取到的数据是")
    print(provider_file.head())
    log('获取到excel数据，转化为了list'+"\n")

    return provider_list


def getManyi(key="总体",index=3):

    if not os.path.exists(key):
        os.mkdir(key)

    mainManyi = getInfoList([index])

    Counter = getCouter(mainManyi)
    mean = mainManyi.mean()
    bili = (Counter['4']+Counter['5'])/len(mainManyi)

    print("{}满意度的平均分为{}".format(key,mean))
    print("{}有超过一半的客户感到满意或者很满意，比例为{}".format(key,bili))

    bar = (
        Bar(init_opts=opts.InitOpts(theme=ThemeType.LIGHT))
        .add_xaxis(["1", "2", "3", "4", "5"])
        .add_yaxis("比例",[Counter['1']/len(mainManyi), Counter['2']/len(mainManyi), Counter['3']/len(mainManyi), Counter['4']/len(mainManyi), Counter['5']/len(mainManyi)])
        .set_global_opts(title_opts=opts.TitleOpts(title="{}满意度".format(key), )))
    bar.render("{}/bar.html".format(key))

    table = (Table().add(headers=['Key','Value'],rows=[["{}满意度的平均分".format(key),mean],["{}好感客户比例".format(key),bili]]))
    table.render(path='{}/table.html'.format(key))

    

def getCouter(col):
    couter1 = 0
    couter2 = 0
    couter3 = 0
    couter4 = 0
    couter5 = 0
    for i in col:
        if(i[0] == 1):
            couter1 += 1
        if(i[0] == 2):
            couter2 += 1
        if(i[0] == 3):
            couter3 += 1
        if(i[0] == 4):
            couter4 += 1
        if(i[0] == 5):
            couter5 += 1
    res = {
        '1': couter1,
        '2': couter2,
        '3': couter3,
        '4': couter4,
        '5': couter5,
    }
    return res

def getXiangguan():
    data =  pd.DataFrame(pd.read_excel(excel_dir,))
    print(data.corr())
    table = (Table().add(headers=["相关系数",'质量满意度','包装满意度','物流速度满意度','售后满意度','价格满意度','支付满意度'],rows=[['与总体满意度的相关系数','0.767943','0.716976','0.721767','0.489723','0.660923','0.766655']]))
    table.render(path='correl.html')

def getGuanjian():
    data =  pd.DataFrame(pd.read_excel(excel_dir,))
    mean = data.mean()
    print(type(mean))
    table = (Table().add(headers=['关键维度满意度得分','平均分'],rows=[['质量满意度','4.254902'],['包装满意度','4.078431'],['物流速度满意度','3.803922'],['售后满意度','4.117647'],['价格满意度','4.058824'],['支付满意度','4.117647'],]))
    table.render(path='keyScore.html')
    table1 = (Table().add(headers=['关键维度满意度比例','达到满意'],rows=[['质量满意度','0.86'],['包装满意度','0.76'],['物流速度满意度','0.70'],['售后满意度','0.80'],['价格满意度','0.76'],['支付满意度','0.84'],]))
    table1.render(path='keyScale.html')

def getManyiDescribe():
    getManyi("总体",3)
    getManyi("产品质量",4)
    getManyi("产品包装",5)
    getManyi("物流速度",6)
    getManyi("售后评分",7)
    getManyi("价格预期",8)
    getManyi("支付服务",9)

if __name__ == "__main__": 
    getManyiDescribe()
    getGuanjian()
    getXiangguan()
