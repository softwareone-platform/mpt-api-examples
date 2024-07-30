#!/usr/bin/env python

import httpx
import json
import openpyxl

from config import Config

cfg = Config()

def download_rest_collection(name, query):

    offset = 0
    limit = 50
    total = 0
    objects = []

    print(f'downloading {name} collection: {query}')
    idx = 0

    client = httpx.Client(http2=True)

    while True:

        print(f'  {idx}. iterating {name} offset={offset} total={total} ...')

        if '?' in query:
            url = cfg.mpt_platform_url + f'{query}&offset={offset}&limit={limit}'
        else:
            url = cfg.mpt_platform_url + f'{query}?offset={offset}&limit={limit}'

        result = client.get(
            url = url,
            headers = {
                "Authorization": f"Bearer {cfg.mpt_platform_token}",
            }
        )

        if result.status_code != 200:
            raise Exception(f'failed to read, please check that your API Token ({cfg.mpt_platform_token}) is valid.')

        ret = json.loads(result.text)

        total = ret['$meta']['pagination']['total']
        data = ret['data']
        for object in data:
            objects.append(object)

        offset += limit
        if offset > total:
            print('collection has been retreived successfully.')
            break

        idx += 1
    
    objects_count = len(objects)

    print(f'total objects downloaded: {objects_count}')
    return objects


def save_rest_collection(collection, fname):

    print(f'writing {fname} ...')
    txt = json.dumps(collection, indent=4)
    with open(fname, "w") as f:
        f.write(txt)


def convert_json_to_excel(collection, mapping, fname):

    wb = openpyxl.Workbook()
    sheet = wb.active

    def assign(row, col, val):
        cell = sheet.cell(row=row, column=col)
        cell.value = val

    col = 1
    for h in mapping:
        assign(1, col, h)
        col += 1

    row = 2
    for object in collection:
        col = 1
        for h in mapping:
            keys = h.split('.')
            val = object
            for k in keys:
                if k in val:
                    val = val[k]
                else:
                    val = None
                    break
            assign(row, col, val)
            col += 1
        row += 1
        
    print(f'saving {fname} ...')
    wb.save(fname)
    wb.close()
    

def main():

    orders = download_rest_collection('orders', '/commerce/orders?select=audit&order=-audit.created.at')

    save_rest_collection(orders, f'./output/orders.json')

    json_to_excel_mapping = {
        'id': 'order id',
        'type': 'order type',
        'status': 'order status',
        'externalIds.client': 'order external id',
        'agreement.id': 'agreement id',
        'agreement.name': 'agreement name',
        'price.SPx1': 'price (one time)',
        'price.SPxM': 'price (monthly)',
        'price.SPxY': 'price (yearly)',
        'product.name': 'product name',
        'buyer.name': 'buyer name',
        'seller.name': 'seller name',
    }

    convert_json_to_excel(orders, json_to_excel_mapping, './output/orders.xlsx')


if __name__ == '__main__':
    main()
