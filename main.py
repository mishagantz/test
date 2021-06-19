import sys
from io import StringIO
import json
import time
import os
import requests
import pandas as pd

from xlsxwriter import exceptions

PATH = '~/tinkoff_api/sbp_dictionary/'


def sbp_api(url="https://api.tinkoff.ru/v1/sbp_dictionary", quiet=False):
    headers = {
        'Accept': "*/*",
        'Cache-Control': "no-cache",
        'Host': "api.tinkoff.ru",
        'Accept-Encoding': "gzip, deflate",
        'Connection': "keep-alive",
        'cache-control': "no-cache"
    }
    r = requests.get('https://api.tinkoff.ru/v1/sbp_dictionary')
    if r.ok:
        if not quiet:
            print('API доступен')
        response = requests.request("GET", url, headers=headers).json()
        json_str = json.dumps(response["payload"], ensure_ascii=False)
        json_str_df = pd.read_json(StringIO(json_str))
        del json_str_df['brand']
        column_dtype_chng = {
            'bankMemberId': 'str',
            'name': 'str',
            'engName': 'str',
            'isMe2meSupported': 'bool'
        }
        json_str_df = json_str_df.astype(dtype=column_dtype_chng)
        r.close()
        return json_str_df
    else:
        print(r.raise_for_status())


def w_to_excel(attempts=0):
    if os.path.exists(PATH):
        os.remove(PATH)
        print('Existing file deleted')
    try:
        pd.io.formats.excel.header_style = None
        writer = pd.ExcelWriter(path=PATH, date_format='%d.%m.%d %H:%M:%S', engine='xlsxwriter')
        sbp_api().to_excel(writer, sheet_name='banks list', index=False)
        workbook = writer.book
        worksheet = writer.sheets['banks list']
        header_format = workbook.add_format({
            'bold': False,
            'border': 0
        })
        for col_num, value in enumerate(sbp_api(quiet=True).columns.values):
            worksheet.write(0, col_num, value, header_format)
        writer.save()
        sys.exit()
    except exceptions.FileCreateError:
        time.sleep(10)
        print('Please close "mergedfile.xlsx" so program can replace it')
        attempts += 1
        w_to_excel(attempts)


w_to_excel()
