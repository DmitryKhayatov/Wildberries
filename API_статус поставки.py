import base64
import zipfile
from io import BytesIO
import os
import requests
import json
import pandas as pd
from datetime import datetime

df_seller_id = pd.read_excel(r"c:\Users\User\Парсеры_161024\seller_id.xlsx")
seller_id = df_seller_id['seller_id'][0]
#seller_id = 'c68111eb-15d1-5910-90ae-9666f334e5bd'


def resp_funct(filter_value=-2):
    url = f"https://seller-supply.wildberries.ru/ns/sm-supply/supply-manager/api/v1/supply/listSupplies"
    #header_value = f"BasketUID=d9b7509528994afd9e9f7416bd6ab9de; ___wbu=9d701a62-0fbe-4172-8a18-82d55c5d0962.1703672299; wb-pid=gYGeqrvXWLVKUbcNANNO2yONAAABkC9uSRUgkVneyuZwfD0yHkDbGKjUhjCJBNN6UFnh3J4w7SFNXw; wbx-validation-key=97547a22-46a4-497a-8fc7-32df72cd2c6d; external-locale=ru; _wbauid=5259289681721046210; WBTokenV3=eyJhbGciOiJSUzI1NiIsInR5cCI6IkpXVCJ9.eyJpYXQiOjE3MjEyMDE0NzMsInZlcnNpb24iOjIsInVzZXIiOiIxNDU4NTk5OSIsInNoYXJkX2tleSI6IjIiLCJjbGllbnRfaWQiOiJzZWxsZXItcG9ydGFsIiwic2Vzc2lvbl9pZCI6IjFhMmU1NTg1MDg5ZjRmMzFiNjkyNGJkYzI5NGFjNjFkIiwidXNlcl9yZWdpc3RyYXRpb25fZHQiOjE2NzE3MzIyNjMsInZhbGlkYXRpb25fa2V5IjoiODkxNzU3YTFmYWIwMDkyODU3NDc2NzEwNDZjY2E4MmJhMzM1ZWIxMjE5ZjBlZGE1ODgxZDk2YjM3MzUzMzM2NSIsInBob25lIjoiIn0.OWj-o7XwrOLxvqUxHlBb2KOwByH-w7TnxrKCWZyDjFqvHiSSOt5Bo4XEZx4YUuraQwpnyPHU1HEUJv06coF6kVVDSp_wTLco1hWKjnxdxOpop5OZbMNb1lJ-21Nhp9gquLW-Rnaeyc35oibuRcDoB2R1rzRCW9_sw6ApTYKnUBk1eMpYjK8RRr_M94qUCle4WML2ct0ASZLqbNdwJP2AhczMn3EOVAfHn9esGrehGNPPNDUqs3rx5X-WGEzuaI7V09ws3J5onNP29fVOX98Rc94qPyCdfF0DOP6Ek11BD_MR7u6UIXU2e2SJ88LDD3EqpjI-3-Al3ZrTg_4CCv2Xbg; wb-sid=dfd5cd3f-773f-4add-bcd9-daf19efe53a4; wb-id=gYHyqBiu2LdKc4qeE19d-CWBAAABkO97MNJpsqb2fDHmVqFtwJuFGZY2JmzvvQlYKm36DfLB2B4jlmRmZDVjZDNmLTc3M2YtNGFkZC1iY2Q5LWRhZjE5ZWZlNTNhNA; x-supplier-id-external={seller_id}; __zzatw-wb=MDA0dC0cTHtmcDhhDHEWTT17CT4VHThHKHIzd2UqP2olZFBiKDVRP0FaW1Q4NmdBEXUmCQg3LGBwVxlRExpceEdXeiwbE3drKVcLE19ERmllbQwtUlFRS19/Dg4/aU5ZQ11wS3E6EmBWGB5CWgtMeFtLKRZHGzJhXkZpdRVQOA0YQkZ1eClDblNjfVwgdVtWeylLRTJtLFM4PmE+dV9vG3siXyoIJGM1Xz9EaVhTMCpYQXt1J3Z+KmUzPGwfYUdZJUtZVX0sHw1pN2wXPHVlLwkxLGJ5MVIvE0tsP0caRFpbQDsyVghDQE1HFF9BWncyUlFRS2EQR0lrZU5TQixmG3EVTQgNND1aciIPWzklWAgSPwsmIBN5ayNVDw9jQUh0bxt/Nl0cOWMRCxl+OmNdRkc3FSR7dSYKCTU3YnAvTCB7SykWRxsyYV5GaXUVWA8QX25yb3smcWoeX0ReIUkRSjJaThp0cCxPCxNjb3UnfC1DVxlRDxZhDhYYRRcje0I3Yhk4QhgvPV8/YngiD2lIYCFFVU1/LR0ZfXApS3FPLH12X30beylOIA0lVBMhP05y7DAOcw==; cfidsw-wb=B7ELU/VKuLChoyApNoBtkLLGOp8zVhnt5hOkkoB53LrrWW5kjPThGwKtj2BP27p3H7mi3AH8MiBbMlqoJgmMOFP7RyECONybXpQevGxWEGXVFiyG3FMlSHKeJhVk4oXLOxBZimoDfb8Z3yYpyQz5HCnol7tm1eSo+N/LV1ud"
    # header_value = f"BasketUID=d9b7509528994afd9e9f7416bd6ab9de; ___wbu=9d701a62-0fbe-4172-8a18-82d55c5d0962.1703672299; wb-pid=gYGeqrvXWLVKUbcNANNO2yONAAABkC9uSRUgkVneyuZwfD0yHkDbGKjUhjCJBNN6UFnh3J4w7SFNXw; external-locale=ru; _wbauid=5259289681721046210; wb-id=gYF3igJwl0JMFZTk9QxQHl_LAAABkUcrIhGXjbEY1n4NqYN01a7n4_H6bj-5342MYfUW80RsRaQIkmYzZmMzYWUzLTU1YTEtNDM3ZS1hNmExLWMwYmVlNjE0ZGZkOA; wbx-validation-key=98594e2b-33c0-4905-bc17-b9c47cb0cd95; x-supplier-id-external={seller_id}; WBTokenV3=eyJhbGciOiJSUzI1NiIsInR5cCI6IkpXVCJ9.eyJpYXQiOjE3MjM3OTQ3MzEsInZlcnNpb24iOjIsInVzZXIiOiIxNDU4NTk5OSIsInNoYXJkX2tleSI6IjIiLCJjbGllbnRfaWQiOiJzZWxsZXItcG9ydGFsIiwic2Vzc2lvbl9pZCI6IjFhMmU1NTg1MDg5ZjRmMzFiNjkyNGJkYzI5NGFjNjFkIiwidXNlcl9yZWdpc3RyYXRpb25fZHQiOjE2NzE3MzIyNjMsInZhbGlkYXRpb25fa2V5IjoiMjQ3ZjhiYjUyNWI3YzI4MTQ0NTRlZTY5MDkyMzQ3OGNhN2UyZGY3ODhlYzg3ZGE2OTU1NmM0ZDk5M2IxNmE2MCIsInBob25lIjoiIn0.mpy9jyLAUWgEZrFYQ0Xav3LwJ8Y9ZNjoU6CdGDAmV2fBeu_yzIPswO5k_6KjIxvk5521hcfW1RjOAtaicISWMr4koxEuli6g-Ry1tou-0TCJJQYt0Ofemmn0CYP9cai4M2zo0TEmpj5HvLYpUiH2bpa5E5tmezMD94tJx7DEePchWyFLjJ2OQJ687HrFruQo-d9hmWIwEICW1bQtifHomYw3pbOBHJshP5DNWYPdRXb50JMDE9i3LUBmYslJD-tqiP0LkRVlGQlrttHz4uXeqDEyr9E5AUCaO2eYZSsVnISEkQVc6gxrg9bRql5lZ_cnkyqmxWHoiS0quw_wm5HxmQ; cfidsw-wb=rACZrSaafF0gCcjzZ49DOvYCk+FEEyWgufprr+WGWruohYTnDsVXziaoU5I56RO9DJH4xPYVluFF6Af5oTUAKYYX6rbxEW1sG7iesJRoCMJZ/xe5IOf8J/UGjHVKFixoaGlh6o7XCPGSsIEGePPK+nQzgJv3yvOPCcl+Sf8JMQ==; __zzatw-wb=MDA0dC0cTHtmcDhhDHEWTT17CT4VHThHKHIzd2UqP2olZFBiKDVRP0FaW1Q4NmdBEXUmCQg3LGBwVxlRExpceEdXeiwbFXdwKlEJDl9DSGllbQwtUlFRS19/Dg4/aU5ZQ11wS3E6EmBWGB5CWgtMeFtLKRZHGzJhXkZpdRVQOA0YQkZ1eClDblNjfVwgdVtWeylLRTJtLFM4PmE+dV9vG3siXyoIJGM1Xz9EaVhTMCpYQXt1J3Z+KmUzPGwfY0deJkVXUH0sGg1pN2wXPHVlLwkxLGJ5MVIvE0tsP0caRFpbQDsyVghDQE1HFF9BWncyUlFRS2EQR0lrZU5TQixmG3EVTQgNND1aciIPWzklWAgSPwsmIBN7ayhWCQ1eQUhxbxt/Nl0cOWMRCxl+OmNdRkc3FSR7dSYKCTU3YnAvTCB7SykWRxsyYV5GaXUVUz0/GERFdnYmPyNTaEReH0pcSgkoGhZ0dFUICg5bQHImLi1uVxlRDxZhDhYYRRcje0I3Yhk4QhgvPV8/YngiD2lIYCFHVVIIJxsUfHIrS3FPLH12X30beylOIA0lVBMhP05y7yhVyA=="
    header_value = f"BasketUID=d9b7509528994afd9e9f7416bd6ab9de; ___wbu=9d701a62-0fbe-4172-8a18-82d55c5d0962.1703672299; wb-pid=gYGeqrvXWLVKUbcNANNO2yONAAABkC9uSRUgkVneyuZwfD0yHkDbGKjUhjCJBNN6UFnh3J4w7SFNXw; external-locale=ru; _wbauid=5259289681721046210; wb-id=gYF3igJwl0JMFZTk9QxQHl_LAAABkUcrIhGXjbEY1n4NqYN01a7n4_H6bj-5342MYfUW80RsRaQIkmYzZmMzYWUzLTU1YTEtNDM3ZS1hNmExLWMwYmVlNjE0ZGZkOA; x-supplier-id-external={seller_id}; wbx-validation-key=34921d56-55ae-4bd5-958c-acb5f8d694df; WBTokenV3=eyJhbGciOiJSUzI1NiIsInR5cCI6IkpXVCJ9.eyJpYXQiOjE3MjkwNzU1NzQsInZlcnNpb24iOjIsInVzZXIiOiIxNDU4NTk5OSIsInNoYXJkX2tleSI6IjIiLCJjbGllbnRfaWQiOiJzZWxsZXItcG9ydGFsIiwic2Vzc2lvbl9pZCI6IjFhMmU1NTg1MDg5ZjRmMzFiNjkyNGJkYzI5NGFjNjFkIiwidXNlcl9yZWdpc3RyYXRpb25fZHQiOjE2NzE3MzIyNjMsInZhbGlkYXRpb25fa2V5IjoiNzA4ZDExNDQ4MWZhMGI2N2FlYmU0NmI0OTJjMThiZTIyODc3OWM1YmQxMTM3MjUwODFiNGFkN2Y2MzBkZGQ5MyJ9.AYSGY132eBxALObDxq8Ns1Vhi0SQKecyyKp6_wAYeRA0xA_6kf7hPoBEjmWHhL-GOqF3NKK7GDj8ho07QrYVOvgCj1nhTWFgex3Gi9GPmFxr0-qFc3XPA8bMGdZXN35wIQ1yP9JObQFP-Kb7xrsuq878svKnlIeMvTgMEyWCeQ1iXDOvSjESl6NONTRWBswZ7InhW7AQRbXs2WRNtJlxBLHy1RXc2PWEiBeugxWeYTG9JeUAiZnpCFQX1iCuSAGEmNAtvwAoT4U4L7PkWigYG0Kz1JytLkZx1Qak_zhTMdaP233t5IAf50WtnrzrnyLlXGkpffboEXDFEfstvg2J_Q; __zzatw-wb=MDA0dC0cTHtmcDhhDHEWTT17CT4VHThHKHIzd2UqP2olZFBiKDVRP0FaW1Q4NmdBEXUmCQg3LGBwVxlRExpceEdXeiwbGndyKFQQD2RCSWllbQwtUlFRS19/Dg4/aU5ZQ11wS3E6EmBWGB5CWgtMeFtLKRZHGzJhXkZpdRVQOA0YQkZ1eClDblNjfVwgdVtWeylLRTJtLFM4PmE+dV9vG3siXyoIJGM1Xz9EaVhTMCpYQXt1J3Z+KmUzPGwfaEdgJEheUQoqIQ1pN2wXPHVlLwkxLGJ5MVIvE0tsP0caRFpbQDsyVghDQE1HFF9BWncyUlFRS2EQR0lrZU5TQixmG3EVTQgNND1aciIPWzklWAgSPwsmIBMIaypUDBRfRkZ2bxt/Nl0cOWMRCxl+OmNdRkc3FSR7dSYKCTU3YnAvTCB7SykWRxsyYV5GaXUVTzg/W292JjEmPmlPaERdUENcSgpWTRd0JCpWOD5iQkEnMS1DVxlRDxZhDhYYRRcje0I3Yhk4QhgvPV8/YngiD2lIYCFMVVR+KiIWeG8oS3FPLH12X30beylOIA0lVBMhP05yqwtA/Q==; cfidsw-wb=u+ixTpo2ajfeRavcj985+bNEGpeiDgscD7SUa5d46DRE6DAZQkve+4uvTS66yVieOwKmnlJRk4fvlBxpCzvBB+gQh6bLrLpmpY3hDtEfisIY6Y2nBJk93Jehn1db6fGzhmfXFzpCbn1mP2adEHu+L0/kpUvoR1D1ozu4MBkI3w=="
    headers = {
        "Content-Type": "application/json",
        "User-Agent": "Chrome/115.0.0.0",
        "Accept": "*/*",
        "Accept-Encoding": "gzip, deflate, br",
        "Connection": "keep-alive",
        "cookie": f"{header_value}"
    }

    data = {"params":{"pageNumber":1,"pageSize":10000,"sortBy":"createDate","sortDirection":"desc","statusId":filter_value},"jsonrpc":"2.0","id":"json-rpc_23"}
    json_data = json.dumps(data)

    response = requests.post(url, headers=headers, data=json_data)
    print(response.status_code)
    response_result = response.json()
    data = response_result.get('result', {}).get('data', [])
    return data

def format_resp(data):
    df_data = []
    for item in data:
        #print(item)
        row = {
            'preorderId': item['preorderId'],
            'supplyId': item['supplyId'],
            'boxTypeId': item['boxTypeId'],
            'boxTypeName': item['boxTypeName'],
            'createDate': datetime.fromisoformat(item['createDate']).date(),
            'createTime': datetime.fromisoformat(item['createDate']).time(),
            'changeDate': datetime.fromisoformat(item['changeDate']).date(),
            'changeTime': datetime.fromisoformat(item['changeDate']).time(),
            'detailsQuantity': item['detailsQuantity'],
            'warehouseId': item['warehouseId'],
            'warehouseName': item['warehouseName'],
            'transitWarehouseId': item['transitWarehouseId'],
            'transitWarehouseName': item['transitWarehouseName'],
            'supplyDate': datetime.fromisoformat(item['supplyDate']).date() if item['supplyDate'] != None else None,
            'supplyTime': datetime.fromisoformat(item['supplyDate']).time() if item['supplyDate'] != None else None,
            'factDate': datetime.fromisoformat(item['factDate']).date() if item['factDate'] != None else None,
            'factTime': datetime.fromisoformat(item['factDate']).time() if item['factDate'] != None else None,
            'incomeQuantity': item['incomeQuantity'],
            'statusId': item['statusId'],
            'statusName': item['statusName'],
            'rejectReason': item['rejectReason'],
            'virtualType': item['virtualType'],
            'userUid': item['userUid']
        }
        df_data.append(row)

    df = pd.DataFrame(df_data)
    df['key'] = df.apply(lambda row: row['preorderId'] if pd.notna(row['preorderId']) else row['supplyId'], axis=1)
    return df

print('запрос со статусом "ВСЕ"')
a = resp_funct(-2)
print('запрос со статусом "ПРИНЯТО"')
b = resp_funct(7)
df1 = format_resp(a)
df2 = format_resp(b)

df_combined = pd.concat([df1, df2], ignore_index=True)
df_combined['key'] = df_combined['key'].astype(str)
df_combined['key'] = df_combined['key'].str.replace(r'\.0$', '', regex=True)
df_combined = df_combined.drop_duplicates(subset='key').reset_index(drop=True)
df_combined = df_combined.drop(columns=['key'])

df_combined.to_excel('Статус поставки.xlsx')

print('fin')
