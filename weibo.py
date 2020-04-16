import requests
from lxml import etree
import xlwt

#获取html页面
def gethtmltext(url):
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.149 Safari/537.36"
    }
    res = requests.get(url,headers=headers)
    res.encoding = 'utf-8'
    html = res.text
    return html

#解析页面：1、转化为etree对象；2、定位所有tr标签；3、在每一个tr标签中提取信息
def infolist(lis,html):
    tree = etree.HTML(html)
    trs = tree.xpath('//tbody/tr[position()>1]')   #position指定从哪一行开始的tr标签
    for tr in trs:
        data = []
        keyword = tr.xpath('.//a/text()')[0]
        data.append(keyword)
        num = tr.xpath('.//span/text()')[0]
        data.append(num)
        lis.append(data)


#保存：1、创建wb、ws；2、定义标头；3、将标题写进excel中；4、将信息写到指定位置的表格中
def saveinfolist(lis,path):
    wb = xlwt.Workbook(encoding='utf-8')
    ws = wb.add_sheet("sheet1")
    col_name = ('关键词','指数')
    for i in range(2):
        ws.write(0, i, col_name[i])
    for r in range(len(lis)):
        case = lis[r]
        for c in range(2):
            ws.write(r+1, c, case[c])
    wb.save(path)

def main():
    url = 'https://s.weibo.com/top/summary'
    gethtmltext(url)
    html = gethtmltext(url)
    info = []
    infolist(info,html)
    savepath = "热搜.xls"
    saveinfolist(info,savepath)

main()

