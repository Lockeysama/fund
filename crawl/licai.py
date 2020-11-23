import requests
from lxml import html


def login(user_name, password):
    headers = {
        'User-Agent': ('Mozilla/5.0 (Macintosh; Intel Mac OS X 10_14_6) '
                       'AppleWebKit/537.36 (KHTML, like Gecko) '
                       'Chrome/85.0.4183.83 Safari/537.36')
    }
    url = 'https://www.licai.com/api/v1/auth/login/pass'
    s = requests.session()
    s.headers = headers
    resp = s.put(url, json={'username': user_name, 'password': password})
    print(resp)

    url2 = 'https://www.licai.com/api/v1/auth/login'
    resp2 = s.post(url)
    print(resp2)
    return s


def get_net_worth(session, fid):
    url = 'https://www.licai.com/simu/product/{}'.format(fid)
    resp = session.get(url)
    print(resp)
    if resp.status_code == 200:
        table_xpath = ('//*[@class="safe-padding"]/div/div'
                       '/div[@class="table-core-container"]'
                       '/div/div/table/tbody/tr')
        doc = html.document_fromstring(resp.content)

        earnings = doc.xpath('//*[@class="rate-num"]/text()')
        if len(earnings):
            earnings = earnings[1:]
            earnings = [e.strip() for e in earnings]

        sharpe_ratio = doc.xpath('//*[@class="table-item table-item-rowspan-1"][5]/div[2]/div/div[1]/span/text()')
        if len(sharpe_ratio):
            sharpe_ratio = [sharpe_ratio[0].strip()]

        itmes = doc.xpath(table_xpath)
        result = []
        if len(itmes):
            for item in itmes:
                values = item.xpath('td/div/span/text()')
                if len(values) == 5:
                    result.append(values)
        return earnings, sharpe_ratio, result
    else:
        print('failed')


if __name__ == '__main__':
    s = login('18757878001', '22020319')
    print(get_net_worth(s, 'P138112'))
