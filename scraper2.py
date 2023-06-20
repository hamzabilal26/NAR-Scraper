import random
from multiprocessing import Pool
import requests
import json
import openpyxl

all_useragents = []
file = open("chrome_useragents.txt", 'r')
for each in file.readlines():
    each = each.replace("\n", '')
    all_useragents.append(each)
file.close()

def get_data(item_input_data):
    this_url = item_input_data[0]
    this_phone = item_input_data[1]
    for i in range(0, 10):

        headers = {
            'authority': 'nar.m1gateway.realtor',
            'accept': 'application/json, text/plain, */*',
            'accept-language': 'en-GB,en-US;q=0.9,en;q=0.8',
            'authorization': 'Basic bmFycmVhbHRvcmRpcmVjdG9yeTokV2Q/S0huN15Va3EtcWo1',
            'origin': 'https://directories.apps.realtor',
            'referer': 'https://directories.apps.realtor/',
            'sec-ch-ua': '"Google Chrome";v="105", "Not)A;Brand";v="8", "Chromium";v="105"',
            'sec-ch-ua-mobile': '?0',
            'sec-ch-ua-platform': '"macOS"',
            'sec-fetch-dest': 'empty',
            'sec-fetch-mode': 'cors',
            'sec-fetch-site': 'cross-site',
            'user-agent': random.choice(all_useragents),
            'withcredentials': 'true',
        }
        try:
            response = requests.get(f'https://nar.m1gateway.realtor/ext/office/{this_url}', headers=headers)
            # print(i)
            data2 = json.loads(response.text)
            name = data2['OfficeBusinessName']
            print(f"Name:{name}")

            street_add = data2['StreetAddressLine1']
            print(f"Street:{street_add}")

            city = data2['StreetCity']
            state = data2['StreetState']
            zip = data2['StreetZip']
            location = f"{city} {state}, {zip}"
            print(f"Location:{location}")

            d_realtor = data2['OfficeContactDrName']
            print(f"Designated Realtor:{d_realtor}")

            manager = data2['OfficeContactManagerName']
            print(f"Manager:{manager}")

            state_ass = data2['PrimaryStateAssociationName']
            print(f"State Association:{state_ass}")

            local_ass = data2['PrimaryLocalAssociationName']
            print(f"Local Association:{local_ass}")

            row = [name, street_add, location, d_realtor, manager, this_phone, state_ass, local_ass]
            return row



        except Exception as e:
            print("Page not found!")



if __name__ == "__main__":
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(['Name', 'Address', 'Location', 'Designated Realtor', 'Office Contact Manager', 'Phone Number',
               'State Association', 'Local Association'])

    headers = {
        'authority': 'nar.m1gateway.realtor',
        'accept': 'application/json, text/plain, */*',
        'accept-language': 'en-GB,en-US;q=0.9,en;q=0.8',
        'authorization': 'Basic bmFycmVhbHRvcmRpcmVjdG9yeTokV2Q/S0huN15Va3EtcWo1',
        # Already added when you pass json=
        # 'content-type': 'application/json',
        'origin': 'https://directories.apps.realtor',
        'referer': 'https://directories.apps.realtor/',
        'sec-ch-ua': '"Google Chrome";v="105", "Not)A;Brand";v="8", "Chromium";v="105"',
        'sec-ch-ua-mobile': '?0',
        'sec-ch-ua-platform': '"macOS"',
        'sec-fetch-dest': 'empty',
        'sec-fetch-mode': 'cors',
        'sec-fetch-site': 'cross-site',
        'user-agent': random.choice(all_useragents),
        'withcredentials': 'true',
    }

    json_data = {
        'OfficeStreetCountry': 'US',
        'StreetState': 'WY',
    }

    response = requests.post('https://nar.m1gateway.realtor/ext/Office/Search/Directory', headers=headers, json=json_data)
    print(response.status_code)
    data = json.loads(response.text)
    data_list = []
    for i in range(0, 820):
        url = data[i]['OfficeId']
        phone = data[i]['PhoneNumber']
        data_list.append([url, phone])

    p = Pool(100)
    results = p.map(get_data, data_list)
    p.terminate()
    p.join()

    for res in results:
        if res:
            ws.append(res)
    wb.save('54.xlsx')

