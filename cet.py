"""
Created on Thu Aug 25 14:47:53 2022

@author: Xu
"""


from time import sleep
from requests import get
from pandas import read_excel
from easygui import diropenbox,fileopenbox,choicebox

msg = '请选择用来查询的数据文件（xls文件）'
data =read_excel(fileopenbox(msg))

msg ="你想要查询什么？"
title = "考试类型"
choices = ["全国大学英语四级考试(CET4)","全国大学英语六级考试(CET6)","全国大学日语四级考试(CJT4)","全国大学日语六级考试(CJT6)","全国大学德语四级考试(PHS4)","全国大学德语六级考试(PHS6)","全国大学俄语四级考试(CRT4)","全国大学俄语六级考试(CRT6)","全国大学法语四级考试(TFU4)"]
ks=choicebox(msg, title, choices)
numb={'全国大学英语四级考试(CET4)':1,'全国大学英语六级考试(CET6)':2,'全国大学日语四级考试(CJT4)':3,'全国大学日语六级考试(CJT6)':4,'全国大学德语四级考试(PHS4)':5,'全国大学德语六级考试(PHS6)':6,'全国大学俄语四级考试(CRT4)':7,'全国大学俄语六级考试(CRT6)':8,'全国大学法语四级考试(TFU4)':9}
level=numb['{}'.format(ks)]
# 添加几个字段，用于存储爬取到的信息，
# 注意从7（索引从0开始）开始，因为我的表中前面已经有7个字段，这里需要根据自己的信息表调整
data.insert(2, '准考证号', '')
data.insert(4, ks+"成绩", '')
data.insert(5, '听力部分', '')
data.insert(6, '阅读部分', '')
data.insert(7, '写作部分', '')
 
# 添加请求头
headers = {
    'Accept': '*/*',
    'Accept-Encoding': 'gzip, deflate',
    'Accept-Language': 'zh-CN,zh;q=0.9',
    'Connection': 'keep-alive',
    'Host': 'cachecloud.neea.cn',
    'Origin': 'http://cet.neea.cn',
    'Referer': 'http://cet.neea.cn/',
    'Cookie': 'Hm_lvt_dc1d69ab90346d48ee02f18510292577=1629958972,1629961007,1629989056,1629992054; community=Home; language=1; http_waf_cookie=f27d75db-23f7-4aa938445a95fb8092bcad0ba73630744bd3',
    'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) '
                  'Chrome/98.0.4758.102 Safari/537.36 '
}
 
# 由于是批量，用一个列表存储所有考生的参数
params_list = []
for i in range(len(data)):
    params_list.append({
        'e': 'CET_202206_DANGCI',
        'km': level,
        'xm': data.iloc[i, 0],
        'no': data.iloc[i, 1]
    })
 
# 四六级成绩查询网址
url = "http://cachecloud.neea.cn/api/latest/results/cet"
 
# 遍历每一个考生的数据
for i in range(len(data)):
    # requests.get(url,params,**args) **args可传入请求头headers
    response = get(url, params_list[i], headers=headers)
    print("*******************************************************************")
    print("开始爬取%s的成绩......" % params_list[i]['xm'])
    # 由于response是json格式，可用response.json()接受
    json_data = response.json()
 
    # 判断学生是否参加了考试
    if json_data['code'] == 200:
        # 是，便开始存储成绩信息
        data.iloc[i, 2] = json_data['zkzh']  # 准考证号
        data.iloc[i, 4] = json_data['score']  # 六级分数
        data.iloc[i, 5] = json_data['sco_lc']  # 听力部分
        data.iloc[i, 6] = json_data['sco_rd']  # 阅读部分
        data.iloc[i, 7] = json_data['sco_wt']  # 写作部分
        info = '%d：%s查询成功，成绩为：%s' % ((i + 1), params_list[i]['xm'], json_data['score'])
        print(info)
    else:
        # 否，输出相关信息
        print('%d：%s未参加考试...' % ((i + 1), params_list[i]['xm']))
    print("*******************************************************************")
 
    # 休息0.3秒，防止爬取过快IP被封
    sleep(0.3)
 
print("已全部查询完毕，正在导出到csv......")
# 导出至csv文件中
msg = '请选择用来保存的数据文件'
saveurl=diropenbox(msg)
filesave="{}\{}成绩.csv".format(saveurl,ks)
data.to_csv(filesave,encoding='utf_8_sig')
print("导出完毕.")
 