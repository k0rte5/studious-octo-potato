# -*- coding: utf-8 -*-

import requests
url='https://rguk.ru/upload/iblock/6f3/ib8g9uy36c6mbkadfgw042skzoum14sg/%D0%98%D0%A5%D0%A2%D0%B8%D0%9F%D0%AD-%D0%BE%D1%87%D0%BD%D0%BE-2%D0%BA%D1%83%D1%80%D1%81.xlsx'
filename=url.split('/')[-1]

r=requests.get(url)
with open(filename, 'wb') as f:
    f.write(r.content)
    print(f"{filename} saved")