#爬取思路：
#创建header池，伪装成浏览器在访问所爬取的网址，每次在header池中随机选择一个header
#设置延时爬取，每次爬取延时0.1s；设置重试(解决因网络波动引起的爬取中断问题)，retry(tries=5, delay=2)
#爬取两个网址（可自行扩展）
#根据该网站的特点：连续爬取30个左右之后，暂封ip 解决方法：设置计数，每爬取20个url之后sleep 5s（可自行调节参数）


#函数：getFromExcel
#功能：读取excel表格获取数据，将数据返回一个列表col_data
#参数：excel_url 为excel表格的路径
def getFromExcel(excel_url):
    # 打开指定路径下的excel表格并读取数据到data_excel
    data_excel = xlrd.open_workbook(excel_url)
    # 通过索引顺利获取sheet；获取第一个sheet
    table = data_excel.sheet_by_index(sheetx=0)
    # 获取某列中所有单元格的数据组成的列表,第一个参数为第几列
    col_data = table.col_values(0, start_rowx=0, end_rowx=None)
    return col_data

#函数：seek_1   第一个可爬取的网址
#功能：输入url，输出公司名、网站
#参数：url_2 url的后半段不同的部分
@retry(tries = 5, delay = 2)    #为该函数设置重试机制
def seek_1(url_2):
    t = 0.1
    #设置1s的延时
    sleep(t)
    #代理ip池，代理网站：http://www.goubanjia.com/
    #ps:下面的代理ip池还未应用到本项目中
    '''
    proxy = [
        {
            'http': 'http://27.151.29.32:8080',
            'https': 'http://27.151.29.32:8080',
        },
        {
            'http': 'http://222.249.238.138:8080',
            'https': 'http://222.249.238.138:8080',
        },
        {
            'http': 'http://124.205.155.146:9090',
            'https': 'http://124.205.155.146:9090',
        },
        {
            'http': 'http://114.249.117.200:9000',
            'https': 'http://114.249.117.200:9000',
        },
        {
            'http': 'http://124.232.133.199:3128',
            'https': 'http://124.232.133.199:3128',
        },
        {
            'http': 'http://112.95.27.253:8088',
            'https': 'http://112.95.27.253:8088',
        },
        {
            'http': 'http://119.57.105.25:53281',
            'https': 'http://119.57.105.25:53281',
        },
        {
            'http': 'http://101.36.160.87:3128',
            'https': 'http://101.36.160.87:3128',
        },
        {
            'http': 'http://123.57.210.164:3128',
            'https': 'http://123.57.210.164:3128',
        },
        {
            'http': 'http://42.59.87.91:1133',
            'https': 'http://42.59.87.91:1133',
        },
        {
            'http': 'http://220.249.149.59:9999',
            'https': 'http://220.249.149.59:9999',
        },
        {
            'http': 'http://219.239.142.253:3128',
            'https': 'http://219.239.142.253:3128',
        },
        {
            'http': 'http://119.57.108.53:53281',
            'https': 'http://119.57.108.53:53281',
        },
        {
            'http': 'http://115.221.241.222:9999',
            'https': 'http://115.221.241.222:9999',
        },
        {
            'http': 'http://171.13.137.108:9999',
            'https': 'http://171.13.137.108:9999',
        },
        {
            'http': 'http://123.55.101.146:9999',
            'https': 'http://123.55.101.146:9999',
        },
        {
            'http': 'http://119.39.112.125:8118',
            'https': 'http://119.39.112.125:8118',
        },
        {
            'http': 'http://123.55.98.212:9999',
            'https': 'http://123.55.98.212:9999',
        },
        {
            'http': 'http://218.93.119.165:9002',
            'https': 'http://218.93.119.165:9002',
        },
    ]
'''
    user_agent = [
        "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1; SV1; AcooBrowser; .NET CLR 1.1.4322; .NET CLR 2.0.50727)",
        "Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 6.0; Acoo Browser; SLCC1; .NET CLR 2.0.50727; Media Center PC 5.0; .NET CLR 3.0.04506)",
        "Mozilla/4.0 (compatible; MSIE 7.0; AOL 9.5; AOLBuild 4337.35; Windows NT 5.1; .NET CLR 1.1.4322; .NET CLR 2.0.50727)",
        "Mozilla/5.0 (Windows; U; MSIE 9.0; Windows NT 9.0; en-US)",
        "Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; Win64; x64; Trident/5.0; .NET CLR 3.5.30729; .NET CLR 3.0.30729; .NET CLR 2.0.50727; Media Center PC 6.0)",
        "Mozilla/5.0 (compatible; MSIE 8.0; Windows NT 6.0; Trident/4.0; WOW64; Trident/4.0; SLCC2; .NET CLR 2.0.50727; .NET CLR 3.5.30729; .NET CLR 3.0.30729; .NET CLR 1.0.3705; .NET CLR 1.1.4322)",
        "Mozilla/4.0 (compatible; MSIE 7.0b; Windows NT 5.2; .NET CLR 1.1.4322; .NET CLR 2.0.50727; InfoPath.2; .NET CLR 3.0.04506.30)",
        "Mozilla/5.0 (Windows; U; Windows NT 5.1; zh-CN) AppleWebKit/523.15 (KHTML, like Gecko, Safari/419.3) Arora/0.3 (Change: 287 c9dfb30)",
        "Mozilla/5.0 (X11; U; Linux; en-US) AppleWebKit/527+ (KHTML, like Gecko, Safari/419.3) Arora/0.6",
        "Mozilla/5.0 (Windows; U; Windows NT 5.1; en-US; rv:1.8.1.2pre) Gecko/20070215 K-Ninja/2.1.1",
        "Mozilla/5.0 (Windows; U; Windows NT 5.1; zh-CN; rv:1.9) Gecko/20080705 Firefox/3.0 Kapiko/3.0",
        "Mozilla/5.0 (X11; Linux i686; U;) Gecko/20070322 Kazehakase/0.4.5",
        "Mozilla/5.0 (X11; U; Linux i686; en-US; rv:1.9.0.8) Gecko Fedora/1.9.0.8-1.fc10 Kazehakase/0.5.6",
        "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/535.11 (KHTML, like Gecko) Chrome/17.0.963.56 Safari/535.11",
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_7_3) AppleWebKit/535.20 (KHTML, like Gecko) Chrome/19.0.1036.7 Safari/535.20",
        "Opera/9.80 (Macintosh; Intel Mac OS X 10.6.8; U; fr) Presto/2.9.168 Version/11.52",
    ]
    session = HTMLSession() #创建会话
    url_1 = 'http://icp.chinaz.com/' #网址中公共的前半部分
    url = url_1 + url_2 #输入待爬取网址http://icp.chinaz.com/
    r = session.get(url, headers = {'User-Agent': random.choice(user_agent)})
    #利用session的get方法将链接的整个网页爬取回来，在爬取时随机从header池选择一个header进行浏览器伪装爬取

    # print(r.html.text) #将网页预处理之后，仅读取文本部分
    # print(r.html.absolute_links )#获取网页中的全部链接;r.html.absolute_links为绝对链接；r.html.links为相对链接
    company_name = '#first > li:nth-child(1) > p > a' #设置目标区域selector（目标爬取区域）
    company_url = '#first > li:nth-child(6) > p' #设置目标区域selector（目标爬取区域）

    results_1 = r.html.find(company_name)
    results_2 = r.html.find(company_url)

    #判断是否爬取到目标值
    if len(results_1)>0 and len(results_2)>0:
        results = (results_1[0].text, results_2[0].text)
    else:
        results = 0 #此时出现异常，本网址爬取失败
    # print(results_1[0].text, results_2[0].text) #输出公司名和网址的文本信息；例如，北京字节跳动科技有限公司 www.toutiao.com
    return results #返回公司名称和网址，且results为tuple类型

#函数：seek_2   第二个可爬取的网址
#功能：输入url，输出公司名、网站
#参数：url_2 url的后半段不同的部分
@retry(tries = 5, delay = 2)
def seek_2(url_2):
    user_agent = [
        "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1; SV1; AcooBrowser; .NET CLR 1.1.4322; .NET CLR 2.0.50727)",
        "Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 6.0; Acoo Browser; SLCC1; .NET CLR 2.0.50727; Media Center PC 5.0; .NET CLR 3.0.04506)",
        "Mozilla/4.0 (compatible; MSIE 7.0; AOL 9.5; AOLBuild 4337.35; Windows NT 5.1; .NET CLR 1.1.4322; .NET CLR 2.0.50727)",
        "Mozilla/5.0 (Windows; U; MSIE 9.0; Windows NT 9.0; en-US)",
        "Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; Win64; x64; Trident/5.0; .NET CLR 3.5.30729; .NET CLR 3.0.30729; .NET CLR 2.0.50727; Media Center PC 6.0)",
        "Mozilla/5.0 (compatible; MSIE 8.0; Windows NT 6.0; Trident/4.0; WOW64; Trident/4.0; SLCC2; .NET CLR 2.0.50727; .NET CLR 3.5.30729; .NET CLR 3.0.30729; .NET CLR 1.0.3705; .NET CLR 1.1.4322)",
        "Mozilla/4.0 (compatible; MSIE 7.0b; Windows NT 5.2; .NET CLR 1.1.4322; .NET CLR 2.0.50727; InfoPath.2; .NET CLR 3.0.04506.30)",
        "Mozilla/5.0 (Windows; U; Windows NT 5.1; zh-CN) AppleWebKit/523.15 (KHTML, like Gecko, Safari/419.3) Arora/0.3 (Change: 287 c9dfb30)",
        "Mozilla/5.0 (X11; U; Linux; en-US) AppleWebKit/527+ (KHTML, like Gecko, Safari/419.3) Arora/0.6",
        "Mozilla/5.0 (Windows; U; Windows NT 5.1; en-US; rv:1.8.1.2pre) Gecko/20070215 K-Ninja/2.1.1",
        "Mozilla/5.0 (Windows; U; Windows NT 5.1; zh-CN; rv:1.9) Gecko/20080705 Firefox/3.0 Kapiko/3.0",
        "Mozilla/5.0 (X11; Linux i686; U;) Gecko/20070322 Kazehakase/0.4.5",
        "Mozilla/5.0 (X11; U; Linux i686; en-US; rv:1.9.0.8) Gecko Fedora/1.9.0.8-1.fc10 Kazehakase/0.5.6",
        "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/535.11 (KHTML, like Gecko) Chrome/17.0.963.56 Safari/535.11",
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_7_3) AppleWebKit/535.20 (KHTML, like Gecko) Chrome/19.0.1036.7 Safari/535.20",
        "Opera/9.80 (Macintosh; Intel Mac OS X 10.6.8; U; fr) Presto/2.9.168 Version/11.52",
    ]
    session = HTMLSession() #创建会话
    url_1 = 'https://www.aizhan.com/' #网址中公共的前半部分
    url = url_1 + url_2 #输入待爬取网址http://icp.chinaz.com/
    r = session.get(url, headers = {'User-Agent': random.choice(user_agent)})
    #利用session的get功能将链接的整个网页爬取回来，在爬取时随机从header池选择一个header进行浏览器伪装

    # print(r.html.text) #将网页预处理之后，仅读取文本部分
    # print(r.html.absolute_links )#获取网页中的全部链接;r.html.absolute_links为绝对链接；r.html.links为相对链接
    company_name = '#icp-table > table > tbody > tr:nth-child(1) > td:nth-child(2)' #设置目标区域selector（目标爬取区域）
    company_url = '#icp-table > table > tbody > tr:nth-child(5) > td:nth-child(2)' #设置目标区域selector（目标爬取区域）

    results_1 = r.html.find(company_name)
    results_2 = r.html.find(company_url)

    # 判断是否爬取到目标值
    if len(results_1)>0 and len(results_2)>0:
        results = (results_1[0].text, results_2[0].text)
    else:
        results = 0 #此时出现异常，爬取失败
    # print(results_1[0].text, results_2[0].text) #输出公司名和网址的文本信息；例如，北京字节跳动科技有限公司 www.toutiao.com
    return results #返回公司名称和网址，且results为tuple类型

#主函数
excel_url = 'E:/pycode/crawler/829-1.xls' #待处理表格所在路径
excelOpened = xlrd.open_workbook(excel_url)  #打开待写入的表格
excelToWrite = copy(excelOpened)
tableOpened = excelToWrite.get_sheet('Sheet1') #取待写入表格的第一个sheet
dataFromExcel = getFromExcel(excel_url) #调用getFromExcel函数，将数据列表赋值给dataFromExcel
sucNum = 0         #爬取成功数
failNum = 0        #爬取失败数
count =0
for i in range(1, len(dataFromExcel)):
    count = count + 1
    if  count == 20:   #每抓取20个网址，就暂停5s
        sleep(5)
        count = 0

    url_2 = dataFromExcel[i]
    dataFromWeb = seek_1(url_2) #调用seek_1函数，将公司名和网站的文本信息列表赋值给dataFromWeb，且为tuple类型
    if dataFromWeb == 0:        #seek_1函数的返回值为0时，爬取失败
        dataFromWeb = seek_2(url_2)  # seek_1 爬取失败时，跳转到seek_2进行爬取；即第一个网址为爬取到相关信息，跳转到第二个网址进行爬取
    #dataFromWeb = seek_2(url_2) #调用seek_2函数，将公司名和网站的文本信息列表赋值给dataFromWeb，且为tuple类型

    ##爬取结束，开始判断、写入excel
    if dataFromWeb == 0:      #等于0时，爬取失败
        print("第",int(i+1),"条出错了！")
        failNum = failNum + 1
        tableOpened.write(i, 4, "未找到相关信息")  # 爬取失败时，在备注中填写“未找到相关信息”
        continue
    else:
        print("第",int(i+1),"条正常！")
        sucNum = sucNum + 1
        tableOpened.write(i, 4, dataFromWeb[0])  # 将网址写入第i+1行，第5列
        tableOpened.write(i, 3, dataFromWeb[1])  # 将公司名写入第i+1行，第4列
print("爬取成功", sucNum, "条")
print("爬取失败", failNum, "条")
excelToWrite.save('829-11.xls') #保存表格





