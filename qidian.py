import requests
from lxml import etree
import xlwt

#获取html页面
def get_htmltext(url):
    headers = {
        "User-Agent":"Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.149 Safari/537.36"
    }
    try:
        res = requests.get(url,headers=headers,timeout=60)
        res.encoding = 'utf-8'
        return res.text
    except:
        return ""


#解析提取信息
def parsing_htmltext(infolist,html):
    #构建etree对象，定位所有小说所在的标签
    tree = etree.HTML(html)
    lis = tree.xpath('//ol/li')

    #循环提取信息
    for li in lis:
        info = []
        name = li.xpath('.//h4[@class="book-title"]/text()')[0]
        info.append(name)
        type = li.xpath('.//p/em[1]/text()')[0]
        info.append(type)
        num = li.xpath('.//p/em[2]/text()')[0]
        info.append(num)
        author = li.xpath('.//span[@class="book-author"]/text()')[0]
        info.append(author)
        link = 'https://m.qidian.com' + li.xpath('.//a/@href')[0]
        info.append(link)
        infolist.append(info)

#保存到文件
def save_data(infolist,path):
    wb = xlwt.Workbook(encoding='utf-8')
    ws = wb.add_sheet('sheet1')
    #写表头
    col_name = ("小说名","类型","字数","作者","链接")
    for i in range(len(col_name)):
        ws.write(0,i,col_name[i])
    #写数据
    for r in range(len(infolist)):
        case = infolist[r]
        for c in range(len(col_name)):
            ws.write(r+1,c,case[c])
    wb.save(path)


#主函数
def main():
    url = 'https://m.qidian.com/rank/readIndex/male'
    get_htmltext(url)
    data = []
    html = get_htmltext(url)
    parsing_htmltext(data,html)
    path = '起点排行.xls'
    save_data(data,path)

main()