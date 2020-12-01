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
        try:
            items = resp.json().get('Data', {}).get('LSJZList')
        except Exception as e:
            return
        result = []
        for item in items:
            y, m, d = [int(i) for i in item.get('FSRQ').split('-')]
            if date(y, m, d).weekday() == 4:
                _d = [item.get('FSRQ'), item.get('DWJZ'), item.get('LJJZ')]
                result.append(_d)
        return result
    else:
        print('failed')
        return


def get_phase_increases(fid):
    headers = {
        'User-Agent': ('Mozilla/5.0 (Macintosh; Intel Mac OS X 10_14_6) '
                       'AppleWebKit/537.36 (KHTML, like Gecko) '
                       'Chrome/85.0.4183.83 Safari/537.36'),
        'Referer': 'http://fundf10.eastmoney.com/jdzf_{}.html'.format(fid)
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
            r = item.xpath('li/text()')[1:6]
            rank = '/'.join(r[-2:])
            r = r[:-2]
            r.append(rank)
            result.extend(r)
        increase = [result[i] for i in range(0, len(result), 4)]
        same_avg = [result[i] for i in range(1, len(result), 4)]
        hs300 = [result[i] for i in range(2, len(result), 4)]
        rank = [result[i] for i in range(3, len(result), 4)]
        real = []
        real.extend(increase)
        real.extend(same_avg)
        real.extend(hs300)
        real.extend(rank)
        return real
    else:
        print('failed')
        return


if __name__ == '__main__':
    urls = [
        ('http://api.fund.eastmoney.com/f10/lsjz?'
         'fundCode=110011&pageIndex=1&pageSize=20'
         '&startDate=&endDate=&_=1605596748694')
    ]
    # for url in urls:
    #     get_net_worth(url)

    phase_increases_urls = [
        '110011',
        '000875'
    ]
    for url in phase_increases_urls:
        print(get_phase_increases(url))
        print('\n')
