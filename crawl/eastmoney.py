from datetime import date

import requests
from lxml import html


def get_net_worth(fid):
    headers = {
        'User-Agent': ('Mozilla/5.0 (Macintosh; Intel Mac OS X 10_14_6) '
                       'AppleWebKit/537.36 (KHTML, like Gecko) '
                       'Chrome/85.0.4183.83 Safari/537.36'),
        'Referer': 'http://fundf10.eastmoney.com/'
    }
    url = ('http://api.fund.eastmoney.com/f10/lsjz?'
           'fundCode={}&pageIndex=1&pageSize=20'
           '&startDate=&endDate=&_=1605596748694').format(fid)
    resp = requests.get(url, headers=headers)
    if resp.status_code == 200:
        print(resp)
        items = resp.json().get('Data', {}).get('LSJZList')
        result = []
        for item in items:
            y, m, d = [int(i) for i in item.get('FSRQ').split('-')]
            if date(y, m, d).weekday() == 4:
                _d = [item.get('FSRQ'), item.get('DWJZ'), item.get('LJJZ')]
                result.append(_d)
        return result
    else:
        print('failed')


def get_phase_increases(fid):
    headers = {
        'User-Agent': ('Mozilla/5.0 (Macintosh; Intel Mac OS X 10_14_6) '
                       'AppleWebKit/537.36 (KHTML, like Gecko) '
                       'Chrome/85.0.4183.83 Safari/537.36'),
        'Referer': 'http://fundf10.eastmoney.com/jdzf_110011.html'
    }
    url = ('http://fundf10.eastmoney.com/FundArchivesDatas.aspx?'
           'type=jdzf&code={}&rt=0.6509032210719485').format(fid)
    resp = requests.get(url, headers=headers)
    if resp.status_code == 200:
        content = resp.text[resp.text.find('<div'):resp.text.find('"};')]
        doc = html.document_fromstring(content)
        result = []
        items = doc.xpath('//*/ul')[2:7]
        for item in items:
             result.append(item.xpath('li/text()')[1:6])
        return result
    else:
        print('failed')


if __name__ == '__main__':
    urls = [
        ('http://api.fund.eastmoney.com/f10/lsjz?'
         'fundCode=110011&pageIndex=1&pageSize=20'
         '&startDate=&endDate=&_=1605596748694')
    ]
    # for url in urls:
    #     get_net_worth(url)

    phase_increases_urls = [
        'http://fundf10.eastmoney.com/FundArchivesDatas.aspx?type=jdzf&code=110011&rt=0.6509032210719485'
    ]
    for url in phase_increases_urls:
        get_phase_increases(url)
