import requests
import json
from lxml import etree
from selenium import webdriver
from selenium.webdriver import ActionChains
import xlwt
import xlrd
cookies = {
    'qgqp_b_id': 'e1e1d321af868898085e9617154d2b78',
    'EMFUND1': 'null',
    'EMFUND2': 'null',
    'EMFUND3': 'null',
    'Eastmoney_Fund': '001158_000001_000011',
    'kforders': '0%3B-1%3B%3B%3B0%2C2%2C24%2C25%2C18%2C19%2C22%2C23%2C21%2C3',
    'EMFUND0': 'null',
    'EMFUND4': '06-30%2021%3A51%3A19@%23%24%u5E7F%u53D1%u6539%u9769%u6DF7%u5408@%23%24001468',
    'EMFUND5': '06-30%2021%3A53%3A43@%23%24%u5BCC%u56FD%u4E2D%u56FD%u4E2D%u5C0F%u76D8%u6DF7%u5408%28QDII%29%u4EBA%u6C11%u5E01@%23%24100061',
    'EMFUND6': '06-30%2021%3A54%3A09@%23%24%u534E%u5B89%u5FB7%u56FD%28DAX%29%u8054%u63A5%28QDII%29A@%23%24000614',
    'EMFUND7': '06-30%2021%3A55%3A11@%23%24%u534E%u5B89%u5FB7%u56FD%28DAX%29ETF%28QDII%29@%23%24513030',
    'EMFUND8': '07-01%2001%3A42%3A42@%23%24%u5DE5%u94F6%u65B0%u6750%u6599%u65B0%u80FD%u6E90%u80A1%u7968@%23%24001158',
    'st_si': '41093717328617',
    'st_asi': 'delete',
    'ASP.NET_SessionId': 'rpanpzvjfie54e2vufu42myq',
    '_adsame_fullscreen_18503': '1',
    'st_pvi': '49393371217988',
    'st_sp': '2022-06-30%2021%3A51%3A19',
    'st_inirUrl': 'http%3A%2F%2Ffund.eastmoney.com%2F000614.html',
    'st_sn': '5',
    'st_psi': '20220706230220372-112200312945-8091357456',
}
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/103.0.0.0 Safari/537.36'}
# url='http://fund.eastmoney.com'
data=xlrd.open_workbook("D:/360data/重要数据/桌面/客户基金分析0701.xlsx")
table=data.sheet_by_index(0)
row=table.nrows
shuju=table.col_values(0)
print(shuju)
wb = xlwt.Workbook(encoding="utf-8")
ws = wb.add_sheet("sheetname")
for j in range(35):
    print(j)
    # url='http://fund.eastmoney.com/'+shuju[j]+'.html?spm=search'
    # url='http://fund.eastmoney.com/'+shuju[j]+'.html?spm=search'
    url = 'http://fundf10.eastmoney.com/gmbd_'+shuju[j]+'.html'
    response = requests.get(url, headers=headers)
    response = requests.get(url, headers=headers)
    txt=response.content.decode('utf-8')
    # txt=response.content.decode("UTF-8")
    # html = etree.HTML(txt)
    # response.encoding='utf-8'

    html = etree.HTML(response.content)
    # result = etree.tostring(html)
    # response.encoding="utf-8"
    # print(response.cookies())
    browser=driver = webdriver.Chrome('D:\迅雷下载\chromedriver_win32\chromedriver.exe')
    browser.get(url)
    # //*[@id="highcharts-26"]/svg/g[5]/g[1]/rect[1]
    # above=browser.find_element_by_xpath('//*[@id="highcharts-26"]/*[name()="svg"]/*[name()="g"][5]/*[name()="g"][1]/*[name()="rect"][1]')
    # print(above)
    # ActionChains(driver).move_to_element(above).perform()
    # above=browser.find_element_by_xpath('//*[@id="body"]/div[15]/div/div/div[1]/div[4]/div[1]/div[2]/h3/a')
    # print(above)
    # ActionChains(driver).click(above).perform()
    # above=browser.find_element_by_xpath('//*[@id="highcharts-28"]/*[name()="svg"]/*[name()="g"][5]/*[name()="g"][3]/*[name()="rect"][1]')
    # # above=browser.find_element_by_xpath('//*[@id="highcharts-28"]/*[name()="svg"]/*[name()="g"][5]/*[name()="g"][3]/*[name()="rect"][1]')
    # print(above)
    # ActionChains(driver).click(above).perform()
    # ret=browser.find_element_by_xpath('//*[@id="highcharts-26"]/svg/g[8]').text
    # print(ret)
    # data2 基金名称
    # ws.write(j,0,shuju[j])
    i=1
    # print('基金名称')
    # ret=browser.find_element_by_xpath('//*[@id="body"]/div[11]/div/div/div[1]/div[1]/div').text
    # print(ret)
    # ws.write(j,i,ret)
    # i=i+1
    # # data3 基金经理
    # print('基金经理')
    # ret=browser.find_element_by_xpath('//*[@id="fundManager"]/div[2]/ul/li[1]/div/div/div[2]/div[1]/a').text
    # print(ret)
    # ws.write(j,i,ret)
    # i=i+1
    # # data4 从业年限
    # print('从业年限')
    # ret=browser.find_element_by_xpath('//*[@id="fundManager"]/div[2]/ul/li[1]/div/div/div[2]/div[4]').text
    # print(ret)
    # ws.write(j,i,ret)
    # i=i+1
    # # data5 3月盈亏
    # print('三月盈亏')
    # ret=browser.find_element_by_xpath('//*[@id="body"]/div[11]/div/div/div[3]/div[1]/div[1]/dl[2]/dd[2]/span[2]').text
    # print(ret)
    # ws.write(j,i,ret)
    # i=i+1
    # # data6 6月盈亏
    # print('6月盈亏')
    # ret=browser.find_element_by_xpath('//*[@id="body"]/div[11]/div/div/div[3]/div[1]/div[1]/dl[3]/dd[2]/span[2]').text
    # print(ret)
    # ws.write(j,i,ret)
    # i=i+1
    # # data7 12月盈亏
    # print('12月盈亏')
    # ret=browser.find_element_by_xpath('//*[@id="body"]/div[11]/div/div/div[3]/div[1]/div[1]/dl[1]/dd[3]/span[2]').text
    # print(ret)
    # ws.write(j,i,ret)
    # i=i+1
    # # data8 成立时间
    # print('成立时间')
    # ret=browser.find_element_by_xpath('//*[@id="body"]/div[11]/div/div/div[3]/div[1]/div[2]/table/tbody/tr[2]/td[1]').text
    # print(ret)
    # ws.write(j,i,ret)
    # i=i+1
    # # data9 基金规模
    # ret=browser.find_element_by_xpath('//*[@id="body"]/div[11]/div/div/div[3]/div[1]/div[2]/table/tbody/tr[1]/td[2]').text
    # print(ret)
    # ws.write(j,i,ret)
    # i=i+1
    # # data10 资产规模变化
    # ret=browser.find_element_by_xpath('//*[@id="highcharts-26"]/*[name()="svg" ]//*[name()="g"][8]/*[name()="text"] ').text
    # print(ret)
    # ws.write(j,i,ret)
    # i=i+1
    # 份额
    ret=browser.find_element_by_xpath('//*[@id="gmbdtable"]/table/tbody/tr[1]/td[4]').text
    print(ret)
    ws.write(j,i,ret)
    i=i+1
    # 上一年
    ret=browser.find_element_by_xpath('//*[@id="gmbdtable"]/table/tbody/tr[2]/td[4]').text
    print(ret)
    ws.write(j,i,ret)
    i=i+1
    # 规模
    ret=browser.find_element_by_xpath('//*[@id="gmbdtable"]/table/tbody/tr[1]/td[5]').text
    print(ret)
    ws.write(j,i,ret)
    i=i+1
    # 上一年
    ret=browser.find_element_by_xpath('//*[@id="gmbdtable"]/table/tbody/tr[5]/td[5]').text
    print(ret)
    ws.write(j,i,ret)

    # i=i+1
    browser.quit()

# data11 持有人结构
# //*[@id="highcharts-28"]/div/span/table/tbody
# ret=browser.find_element_by_xpath('//*[@id="highcharts-28"]/div/span/table/tbody/tr[1]/td[1]')
# print(ret)
# browser.quit()
wb.save("test0805.xls")









