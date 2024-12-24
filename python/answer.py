import requests
from lxml import html
import xlwt

def main():
    data_list = getdata()
    savedata(data_list)

def getdata():
    print('开始读取')
    data_list = []
    ###以下i,str(i+1)需要修改
    for i in range(4,28):
        ##https://www.patenthub.cn/s?p=2&q2=&q=%E4%BA%AC%E5%BE%AE%E9%BD%90%E5%8A%9B&ps=20&s=ad&dm=mix&m=none&fc=%5B%7B%22type%22%3A%22applicant%22%2C%22op%22%3A%22include%22%2C%22values%22%3A%5B%22%E4%BA%AC%E5%BE%AE%E9%BD%90%E5%8A%9B%28%E5%8C%97%E4%BA%AC%29%E7%A7%91%E6%8A%80%E6%9C%89%E9%99%90%E5%85%AC%E5%8F%B8%22%2C%22%E4%BA%AC%E5%BE%AE%E9%BD%90%E5%8A%9B%28%E5%8C%97%E4%BA%AC%29%E7%A7%91%E6%8A%80%E8%82%A1%E4%BB%BD%E6%9C%89%E9%99%90%E5%85%AC%E5%8F%B8%22%2C%22%E4%BA%AC%E5%BE%AE%E9%BD%90%E5%8A%9B%28%E6%B7%B1%E5%9C%B3%29%E7%A7%91%E6%8A%80%E6%9C%89%E9%99%90%E5%85%AC%E5%8F%B8%22%2C%22%E4%BA%AC%E5%BE%AE%E9%BD%90%E5%8A%9B%28%E4%B8%8A%E6%B5%B7%29%E4%BF%A1%E6%81%AF%E7%A7%91%E6%8A%80%E6%9C%89%E9%99%90%E5%85%AC%E5%8F%B8%22%5D%7D%5D&ds=cn
        #https://www.patenthub.cn/s?p=2&q2=&q=%E6%88%90%E9%83%BD%E5%8D%8E%E5%BE%AE&ps=20&s=ad&dm=mix&m=none&fc=%5B%7B%22type%22%3A%22applicant%22%2C%22op%22%3A%22include%22%2C%22values%22%3A%5B%22%E6%88%90%E9%83%BD%E5%8D%8E%E5%BE%AE%E7%94%B5%E5%AD%90%E7%A7%91%E6%8A%80%E6%9C%89%E9%99%90%E5%85%AC%E5%8F%B8%22%2C%22%E6%88%90%E9%83%BD%E5%8D%8E%E5%BE%AE%E7%94%B5%E5%AD%90%E7%A7%91%E6%8A%80%E8%82%A1%E4%BB%BD%E6%9C%89%E9%99%90%E5%85%AC%E5%8F%B8%22%2C%22%E6%88%90%E9%83%BD%E5%8D%8E%E5%BE%AE%E7%94%B5%E5%AD%90%E7%B3%BB%E7%BB%9F%E6%9C%89%E9%99%90%E5%85%AC%E5%8F%B8%22%2C%22%E6%88%90%E9%83%BD%E5%8D%8E%E5%BE%AE%E7%A7%91%E6%8A%80%E6%9C%89%E9%99%90%E5%85%AC%E5%8F%B8%22%5D%7D%5D&ds=cn
        #https://www.patenthub.cn/s?p=2&q2=&q=%E6%99%BA%E5%A4%9A%E6%99%B6&ps=20&s=ad&dm=mix&m=none&fc=%5B%7B%22type%22%3A%22applicant%22%2C%22op%22%3A%22include%22%2C%22values%22%3A%5B%22%E8%A5%BF%E5%AE%89%E6%99%BA%E5%A4%9A%E6%99%B6%E5%BE%AE%E7%94%B5%E5%AD%90%E6%9C%89%E9%99%90%E5%85%AC%E5%8F%B8%22%2C%22%E5%8E%A6%E9%97%A8%E6%99%BA%E5%A4%9A%E6%99%B6%E7%A7%91%E6%8A%80%E6%9C%89%E9%99%90%E5%85%AC%E5%8F%B8%22%2C%22%E6%B5%8E%E5%8D%97%E6%99%BA%E5%A4%9A%E6%99%B6%E5%BE%AE%E7%94%B5%E5%AD%90%E6%9C%89%E9%99%90%E5%85%AC%E5%8F%B8%22%2C%22%E6%88%90%E9%83%BD%E6%99%BA%E5%A4%9A%E6%99%B6%E7%A7%91%E6%8A%80%E6%9C%89%E9%99%90%E5%85%AC%E5%8F%B8%22%5D%7D%5D&ds=cn
        #https://www.patenthub.cn/s?p=2&q2=&q=%E5%AE%89%E8%B7%AF%E4%BF%A1%E6%81%AF%E7%A7%91%E6%8A%80&ps=20&s=ad&dm=mix&m=none&fc=%5B%7B%22type%22%3A%22applicant%22%2C%22op%22%3A%22include%22%2C%22values%22%3A%5B%22%E4%B8%8A%E6%B5%B7%E5%AE%89%E8%B7%AF%E4%BF%A1%E6%81%AF%E7%A7%91%E6%8A%80%E8%82%A1%E4%BB%BD%E6%9C%89%E9%99%90%E5%85%AC%E5%8F%B8%22%2C%22%E4%B8%8A%E6%B5%B7%E5%AE%89%E8%B7%AF%E4%BF%A1%E6%81%AF%E7%A7%91%E6%8A%80%E6%9C%89%E9%99%90%E5%85%AC%E5%8F%B8%22%5D%7D%5D&ds=cn
        #https://www.patenthub.cn/s?p=2&q2=&q=%E7%B4%AB%E5%85%89%E5%90%8C%E5%88%9B&ps=20&s=ad&dm=mix&m=none&fc=%5B%7B%22type%22%3A%22applicant%22%2C%22op%22%3A%22include%22%2C%22values%22%3A%5B%22%E6%B7%B1%E5%9C%B3%E5%B8%82%E7%B4%AB%E5%85%89%E5%90%8C%E5%88%9B%E7%94%B5%E5%AD%90%E6%9C%89%E9%99%90%E5%85%AC%E5%8F%B8%22%5D%7D%5D&ds=cn
        #https://www.patenthub.cn/s?p=2&q2=&q=%E7%B4%AB%E5%85%89%E5%90%8C%E5%88%9B&ps=20&s=ad&dm=mix&m=none&fc=%5B%7B%22type%22%3A%22applicant%22%2C%22op%22%3A%22include%22%2C%22values%22%3A%5B%22%E6%B7%B1%E5%9C%B3%E5%B8%82%E7%B4%AB%E5%85%89%E5%90%8C%E5%88%9B%E7%94%B5%E5%AD%90%E6%9C%89%E9%99%90%E5%85%AC%E5%8F%B8%22%2C%22%E6%B7%B1%E5%9C%B3%E5%B8%82%E5%90%8C%E5%88%9B%E5%9B%BD%E8%8A%AF%E7%94%B5%E5%AD%90%E6%9C%89%E9%99%90%E5%85%AC%E5%8F%B8%22%5D%7D%5D&ds=cn
        #https://www.patenthub.cn/s?p=2&q2=&q=%E4%B8%AD%E7%A7%91%E4%BA%BF%E6%B5%B7%E5%BE%AE&ps=20&s=ad&dm=mix&m=none&fc=&ds=cn
        #https://www.patenthub.cn/s?p=2&q2=&q=%E5%A4%8D%E6%97%A6%E5%BE%AE&ps=20&s=ad&dm=mix&m=none&fc=%5B%7B%22type%22%3A%22applicant%22%2C%22op%22%3A%22include%22%2C%22values%22%3A%5B%22%E4%B8%8A%E6%B5%B7%E5%A4%8D%E6%97%A6%E5%BE%AE%E7%94%B5%E5%AD%90%E9%9B%86%E5%9B%A2%E8%82%A1%E4%BB%BD%E6%9C%89%E9%99%90%E5%85%AC%E5%8F%B8%22%2C%22%E4%B8%8A%E6%B5%B7%E5%A4%8D%E6%97%A6%E5%BE%AE%E7%94%B5%E5%AD%90%E8%82%A1%E4%BB%BD%E6%9C%89%E9%99%90%E5%85%AC%E5%8F%B8%22%5D%7D%5D&ds=cn
        url = 'https://www.patenthub.cn/s?p=' + str(i + 1) + '&q2=&q=%E5%A4%8D%E6%97%A6%E5%BE%AE&ps=20&s=ad&dm=mix&m=none&fc=%5B%7B%22type%22%3A%22applicant%22%2C%22op%22%3A%22include%22%2C%22values%22%3A%5B%22%E4%B8%8A%E6%B5%B7%E5%A4%8D%E6%97%A6%E5%BE%AE%E7%94%B5%E5%AD%90%E9%9B%86%E5%9B%A2%E8%82%A1%E4%BB%BD%E6%9C%89%E9%99%90%E5%85%AC%E5%8F%B8%22%2C%22%E4%B8%8A%E6%B5%B7%E5%A4%8D%E6%97%A6%E5%BE%AE%E7%94%B5%E5%AD%90%E8%82%A1%E4%BB%BD%E6%9C%89%E9%99%90%E5%85%AC%E5%8F%B8%22%5D%7D%5D&ds=cn'
        print('访问网页',i+1)
        response = askurl(url)
        tree = html.fromstring(response)
        div_contents = tree.xpath('//div[@class="content"]')
        for contents in div_contents:
            title = contents.xpath(".//span[@data-property = 'title']/text()") #返回的结果是一个列表。
            if title:
                title = ''.join(title) #如果找到了标题，使用 ''.join() 将列表中的所有文本拼接成一个字符串
                label = ''.join(contents.xpath(".//span[contains(@class, 'ui') and contains(@class, 'horizontal') and contains(@class, 'label')]/text()")) #查找 contents 中符合条件的 <span> 元素，并将其文本内容拼接成字符串，赋值给 label
                applicationNumber = ''.join(contents.xpath(".//span[@data-property = 'applicationNumber']/text()"))
                applicant = ''.join(contents.xpath(".//span[@data-property = 'applicant']/text()"))
                applicationDate = ''.join(contents.xpath(".//span[@data-property = 'applicationDate']/text()"))
                summary = ''.join(contents.xpath(".//span[@data-property = 'summary']/text()"))
                summary = summary.lstrip('\n ') #移除摘要前面的换行符和空格，以确保数据整洁。
                inventor = '. '.join(contents.xpath(".//span[@data-property = 'inventor']/text()"))
                dn = ''.join(contents.xpath(".//span[@class = 'dn']/text()"))
                state = ''.join(tree.xpath('.//a[@data-udn = "{}" ]/text()'.format(dn)))
                number = ','.join(contents.xpath(".//span[@data-property = 'ipc' and @class = 'ipc']/text()"))
                data = {
                    'title': title,
                    'applicationNumber': applicationNumber,
                    'applicant': applicant,
                    'applicationDate': applicationDate,
                    'summary': summary,
                    'inventor': inventor,
                    'dn': dn,
                    'state': state,
                    'label': label,
                    'number': number
                }
                data_list.append(data)
    return data_list

def askurl(url):
    #章家伟（已用)
    # header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36 Edg/130.0.0.0'}
    # cookie = 'T_ID=20241106174735lSRTwJUaglCLGrWrQp; s=ZFEqHi19W28BXko/AlQ4PAcXQzAwGjc8Ng0QIg4HAEAYFy4FVBYZEQ1VBg4ZH09JGBIFCERJTiQRXTMGDQM1SjdmfAYAWlZRDB4bRUFWGDoAUVBfEkJWUU0e; pref="ds:cn,s:score!,dm:mix_10"; Qs_lvt_241723=1730886457%2C1731574355; Qs_pv_241723=3917051758148824000%2C4401580521962136600%2C4494255528788256000%2C3302848100549516000%2C258331566790701470; _nxid=431275; l=1'
    #毛博（已用）
    # header = {'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36 Edg/131.0.0.0'}
    # cookie = 'T_ID=20241204170103FGSrVrIVkIxZwGpuxY; source=b64:ZGlyZWN0CW51bGwJL3M/cD0yJnEyPSZxPSVFNiU4OCU5MCVFOSU4MyVCRCVFNSU4RCU4RSVFNSVCRSVBRSZwcz0yMCZzPXNjb3JlJTIxJmRtPW1peCZtPW5vbmUmZmM9JmRzPWNuCTEyNC42NC4xNy4yMg==; l=1; U_TOKEN=1d0916796ec9f37b2d4b255567e2d192f0bf291f; s=aws8NDFrC1FHZUcyODEVGiFlJigVFSELKC1YBghWRxsUEhReBwUZPyhwQQ4ZH09JHBobIAMFCSwBHnAIA1QpR3F+cgELX11VFRBPWUIRUBUFHxkOXA==; _nxid=438493; pref="ds:cn,s:score!,dm:mix_20"'
    #晓凯（已用）
    # header = {'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36 Edg/130.0.0.0'}
    # cookie = 'T_ID=20231102195317VRWdZMHzmhJUzwYzpy; source=b64:ZGlyZWN0CW51bGwJLwkxMjQuNjQuMTcuMjE2; U_TOKEN=c518de05f92f9cd7b46c14ef18fa3c66b294c3e5; s=SwU5DhdJRktaZFY8BQpBBCZCGhUXT1o0IlJCKg8fFjM8Ng0pJzVYXh9Jfw4ZH09JHBobIAMFCSwBHnAIA1QpR3F+dQECU1ZfFRBPWUIRUBUFHxkOXA==; pref="ds:cn,s:score!,dm:mix_10"; _nxid=331829; l=1; Qs_lvt_241723=1728352769%2C1729042456%2C1729480695%2C1732673070%2C1733392337; Qs_pv_241723=3513467371739894000%2C2249848499473341700%2C2994277435884987000%2C3961877584471993000%2C490048726646521500'
    #文冲
    # header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36 Edg/131.0.0.0'}
    # cookie = 'T_ID=20241106174735lSRTwJUaglCLGrWrQp; _nxid=431275; C_UER_ID=; source=deleted; l=1; s=YHheUlV4VWh8e2ktATY8BRFmIDQoNhFOAQ8LDD0BCSpLMBFAJVgnAiliXQ4ZH09JGBIFCERJTjVIQy0GDQM1SjdmfAYAU1xTAR4bRUFWGDoAUVBfEktcU0Ae; Qs_lvt_241723=1730886457%2C1731574355%2C1733391460; Qs_pv_241723=4401580521962136600%2C4494255528788256000%2C3302848100549516000%2C258331566790701470%2C3788980280880377300; mediav=%7B%22eid%22%3A%22521176%22%2C%22ep%22%3A%22%22%2C%22vid%22%3A%228C2cnR*sr%5B%3Dru%24E6DUXt%22%2C%22ctn%22%3A%22%22%2C%22vvid%22%3A%228C2cnR*sr%5B%3Dru%24E6DUXt%22%2C%22_mvnf%22%3A1%2C%22_mvctn%22%3A0%2C%22_mvck%22%3A0%2C%22_refnf%22%3A0%7D; pref="ds:cn,s:score!,dm:mix_20"'
    # 何鑫
    header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36 Edg/131.0.0.0'}
    cookie = 'T_ID=20241211155141xQJOjasKVurTcSWcBc; source=deleted; l=1; s=c0VbLylBWVV4RQFBQQwnCztwPzEiSTQqMD8iUBseIyU1Sy9EVis1UFcTBg4ZH09JGBIFCERJTjYaRyMGDQM1SjdmfAYHW1ZTCh4bRUFWGDoAUVBfFUNWU0se; _nxid=440253'
    # 王凡硕
    # header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36 Edg/131.0.0.0'}
    # cookie = 'T_ID=20241106174735lSRTwJUaglCLGrWrQp; Qs_lvt_241723=1730886457%2C1731574355%2C1733391460; Qs_pv_241723=4401580521962136600%2C4494255528788256000%2C3302848100549516000%2C258331566790701470%2C3788980280880377300; _nxid=438858; pref="ds:cn,s:score!,dm:mix_20"; C_UER_ID=; source=b64:d3d3LnBhdGVudGh1Yi5jbglodHRwczovL3d3dy5wYXRlbnRodWIuY24vcz9wPTE5JnEyPSZxPSVlOSVhYiU5OCVlNCViYSU5MSVlNSU4ZCU4YSVlNSVhZiViYyVlNCViZCU5MyZwcz0yMCZzPXNjb3JlJTIxJmRtPW1peCZtPW5vbmUmZmM9JmRzPWNuCS9zP3A9MTkmcTI9JnE9JUU5JUFCJTk4JUU0JUJBJTkxJUU1JThEJThBJUU1JUFGJUJDJUU0JUJEJTkzJnBzPTIwJnM9c2NvcmUlMjEmZG09bWl4Jm09bm9uZSZmYz0mZHM9Y24JMzkuMTQ0LjEzOS4zOQ==; l=1; U_TOKEN=e2c9d447a738aa2697ed1a28ad340c0163ef2279; s=CmQjICxDfV5hf3BBPFIEGyxrGhYDMBEyUVoQChAsP1IAJlNbWRIjHlRRcw4ZH09JHBobIAMFCSwBHnAIA1QpR3F+cgYDWVFSFRBPWUIRUBUFHxkOXA=='

    cookie_dict = {i.split("=", 1)[0]: i.split("=", 1)[-1] for i in cookie.split("; ")}
    response = requests.get(url, headers=header, cookies=cookie_dict).content.decode()
    return response

def savedata(data_list):
    print("save.......")
    book = xlwt.Workbook(encoding="utf-8", style_compression=0)  # 创建workbook对象
    sheet = book.add_sheet('复旦微2', cell_overwrite_ok=True)  # 创建工作表
    # 创建一个居中对齐的样式
    style = xlwt.XFStyle()
    alignment = xlwt.Alignment()
    alignment.horz = xlwt.Alignment.HORZ_CENTER
    alignment.vert = xlwt.Alignment.VERT_CENTER
    style.alignment = alignment
    #写入标题
    col = ('序号','申请人','标题','标签','状态','申请日期','DN','申请编号','发明人','摘要','分类号')
    for i in range(len(col)):
        sheet.write(0, i, col[i],style)
    #从第一行开始写入数据
    for i, data in enumerate(data_list, start=1):
        sheet.write(i, 0, str(i),style)  # 序号
        sheet.write(i, 1, data.get('applicant', ''),style)  # 使用 get 方法防止 KeyError
        sheet.write(i, 2, data.get('title', ''),style)
        sheet.write(i, 3, data.get('label', ''),style)
        sheet.write(i, 4, data.get('state', ''),style)
        sheet.write(i, 5, data.get('applicationDate', ''),style)
        sheet.write(i, 6, data.get('dn', ''),style)
        sheet.write(i, 7, data.get('applicationNumber', ''),style)
        sheet.write(i, 8, data.get('inventor', ''),style)
        sheet.write(i, 9, data.get('summary', ''))
        sheet.write(i, 10, data.get('number', ''))
        # numbers = data.get('number', '').split(',')
        # for j, number in enumerate(numbers):
        #     if number.strip():  # 只写入非空的number
        #         sheet.write(i, 10 + j, number.strip(), style)
    book.save("专利-复旦微2.xls")

if __name__ == "__main__":
    print("程序start")
    main()
    print("爬取完毕！")