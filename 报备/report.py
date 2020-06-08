import six
import requests
import xlrd
from bs4 import BeautifulSoup
from urllib import parse

def main():
    excelFile = 'xh.xlsx'
    data = read_xlrd(excelFile)
    for i in range(0,len(data)):
        xh=int(data[i][0])
        print(xh+'：开始报备')
        report(xh,data[i][1])
    
def report(xh,address):
    url = 'http://www.hngczy.cn/jkreport/jkbb.aspx'
    headers = {'content-type': 'application/x-www-form-urlencoded',
               'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/83.0.4103.61 Safari/537.36', 'Upgrade-Insecure-Requests': '1',
               'Referer': 'http://www.hngczy.cn/jkreport/jkbb.aspx','Cache-Control': 'max-age=0'}

    html_str = (requests.get(url)).content
    soup = BeautifulSoup(html_str, features='html.parser')
    status = (soup.find(name="input", attrs={"id": "__VIEWSTATE"})).attrs['value']
    generator = (soup.find(name="input", attrs={"id": "__VIEWSTATEGENERATOR"})).attrs['value']
    validation = (soup.find(name="input", attrs={"id": "__EVENTVALIDATION"})).attrs['value']

    form_data = {'__VIEWSTATE': status, '__VIEWSTATEGENERATOR': generator, '__EVENTVALIDATION': validation,
             'txtXgh': xh, 'submitok': '查   询', 'txtaddress': address}
    data = parse.urlencode(form_data)

    html_str = (requests.post(url, data=data, headers=headers)).content
    soup = BeautifulSoup(html_str, features='html.parser')
    status = (soup.find(name="input", attrs={"id": "__VIEWSTATE"})).attrs['value']
    generator = (soup.find(name="input", attrs={"id": "__VIEWSTATEGENERATOR"})).attrs['value']
    validation = (soup.find(name="input", attrs={"id": "__EVENTVALIDATION"})).attrs['value']

    form_data = {'__VIEWSTATE': status, '__VIEWSTATEGENERATOR': generator, '__EVENTVALIDATION':  validation,
                                                   'txtXgh': xh, 'txtJkinfo': 'njk', 'txtaddress': address,
                                                   'txtcomecs': '', 'txttraffic': '', 'txtHbinfo': 'nhb', 'txtChbinfo': 'nchb',
                                                   'txtCdoubtinfo': 'ncdoubt', 'txtDoubtinfo': 'ndoubt', 'txtWork': 'ncschool',
                                                   'txtBz': '', 'submitbb': '报   备'}
    data = parse.urlencode(form_data)

    # 报备
    html_str = (requests.post(url, data=data, headers=headers)).content

    soup = BeautifulSoup(html_str, features='html.parser')
    name = (soup.find(name="script", attrs={"language": "javascript"}))

    # 结果
    print(name)


def read_xlrd(excelFile):
    data = xlrd.open_workbook(excelFile)
    table = data.sheet_by_index(0)
    dataFile = []

    for rowNum in range(table.nrows):
        # if 去掉表头
        if rowNum > 0:
            dataFile.append(table.row_values(rowNum))

    return dataFile

if __name__ == '__main__':
    main()