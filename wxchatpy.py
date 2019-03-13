import datetime
import re
import time

import requests
import xlrd
import xlwt
from apscheduler.schedulers.blocking import BlockingScheduler  # 定时框架
from bs4 import BeautifulSoup
from wxpy import Bot, Friend, SentMessage

#from datetime import datetime


bot = Bot(cache_path=True)

def oneday():
    url='http://wufazhuce.com/one/'#每一期的链接共同的部分
    #ran=(datetime.today()-datetime.date()).days+2376
    ran=(datetime.date.today()-datetime.date(2019,3,11)).days+2376
    currenturl=url+str(ran)#当前期的链接
    try:
        res=requests.get(currenturl)
        res.raise_for_status()
    except requests.RequestException as e:#处理异常
        print(e)
    else:
        html=res.text#页面内容
        soup = BeautifulSoup(html,'html.parser')
        b=soup.select('.one-cita')#查找“每日一句”所在的标签
        print(b[0].string.split())
        words=str(b[0].string.split())
        words = words.replace("['", "[").replace("']", "]")
        print(words)
        return words

#tuling = Tuling(api_key=你的api')#机器人api
def send_weather(location):
#准备url地址
    path ='http://api.map.baidu.com/telematics/v3/weather?location=%s&output=json&ak=TueGDhCvwI6fOrQnLM0qmXxY9N0OkOiQ&callback=?'
    url = path % location
    response = requests.get(url)
    result = response.json()
    #如果城市错误就按照濮阳发送天气
    if result['error'] !=0:
        location ='北京'
        url = path % location
        response = requests.get(url)
        result = response.json()
    day=datetime.datetime.now().date().strftime('%Y-%m-%d')
    str0 = ('    早上好！这是今天的天气预报！\n')
    if datetime.datetime.strptime(day+'5:30', '%Y-%m-%d%H:%M') <= datetime.datetime.now() \
        and datetime.datetime.strptime(day+'12:30', '%Y-%m-%d%H:%M')> datetime.datetime.now():
        str0 = ('    早上好！这是今天的天气预报！\n')
    elif datetime.datetime.strptime(day+'12:30', '%Y-%m-%d%H:%M') <= datetime.datetime.now() \
        and datetime.datetime.strptime(day+'18:00', '%Y-%m-%d%H:%M')> datetime.datetime.now():
        str0 = ('    下午好！这是今天的天气预报！\n')
    else:
        str0 = ('    晚上好！这是今天的天气预报！\n')
    results = result['results']
    # 取出数据字典
    data1 = results[0]
    # 取出城市
    city = data1['currentCity']
    str1 ='    城市    : %s\n' % city
    # 取出pm2.5值
    pm25 = data1['pm25']
    str2 ='    Pm值    : %s\n' % pm25
    # 将字符串转换为整数 否则无法比较大小
    if pm25 =='':
        pm25 =0
    pm25 =int(pm25)
    # 通过pm2.5的值大小判断污染指数
    if 0 <= pm25 <35:   
        pollution ='优'
    elif 35 <= pm25 <75:
        pollution ='良'
    elif 75 <= pm25 <115:
        pollution ='轻度污染'
    elif 115 <= pm25 <150:
        pollution ='中度污染'
    elif 150 <= pm25 <250:
        pollution ='重度污染'
    elif pm25 >=250:
        pollution ='严重污染'
    str3 ='    污染指数: %s\n' % pollution
    result1 = results[0]
    weather_data = result1['weather_data']
    data = weather_data[0]
    temperature_now = data['date']
    str4 ='    当前温度: %s\n' % temperature_now
    wind = data['wind']
    str5 ='    风向    : %s\n' % wind
    weather = data['weather']
    str6 ='    天气    : %s\n' % weather
    str7 ='    温度    : %s\n' % data['temperature']
    message = data1['index']
    str8 ='    穿衣    : %s\n' % message[0]['des']
    str9 ='    我很贴心: %s\n' % message[2]['des']
    str10 ='    运动    : %s\n' % message[3]['des']
    str11 ='    紫外线 : %s\n' % message[4]['des']
    str = str0 + str1 + str2 + str3 + str4 + str5 + str6 + str7 + str8 + str9 + str10 + str11 +oneday() +'\n ……One fine day'
    return str

#发送函数
def send_message():
    file_path = r'users.xlsx'
    #文件路径的中文转码
    #file_path = file_path.decode('utf-8')

    #获取数据
    data = xlrd.open_workbook(file_path)

    #获取sheet
    table = data.sheet_by_name('Sheet1')

    #获取总行数
    nrows = table.nrows
    #获取总列数
    #ncols = table.ncols


    #获取一行的数值，例如第5行
    #rowvalue = table.row_values(5)

    #获取一列的数值，例如第6列
    #col_values = table.col_values(6)

    #获取一个单元格的数值，例如第5行第6列
    #cell_value = table.cell(5,6).value
    '''
    workbook = xlwt.Workbook(encoding = 'utf-8')
    worksheet = workbook.add_sheet('My Worksheet')
    my_friends = bot.friends()
    i=0
    for friend in my_friends:
        worksheet.write(i, 0, label = friend.name)
        worksheet.write(i, 1, label = friend.nick_name)
        i=i+1
    workbook.save(r'answer.xls')
    '''
    print(nrows)
    for vari in range(nrows-1):
        print(vari)
        #friend = bot.friends().search('Martine')[0]#好友的微信昵称，或者你存取的备注
        print(table.cell(vari+1,0).value)
        print(table.cell(vari+1,2).value)
        friend = bot.friends().search(table.cell(vari+1,0).value)[0]#好友的微信昵称，或者你存取的备注
        friend.send(send_weather(table.cell(vari+1,2).value))    
    
    #friend.send(send_weather(friend.city))
    #friend.send(send_weather('武汉'))
#给全体好友发送
#     for friend in my_friends:
#         friend.send(send_weather(friend.city))
#发送成功通知我
    #bot.file_helper.send(send_weather('濮阳'))
    bot.file_helper.send('发送完毕')
#定时器
print('star')
sched = BlockingScheduler()
#sched.add_job(send_message, 'interval', seconds=10)
sched.add_job(send_message,'cron',month='1-12',day='1-31',hour=7,minute =00)
sched.start()
