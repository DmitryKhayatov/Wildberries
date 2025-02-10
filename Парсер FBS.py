import base64
import base64
import zipfile
from io import BytesIO
import os
import requests
import json
import pandas as pd
import time
from datetime import datetime, timedelta

df_arch_otchets = pd.read_excel(r"c:\Users\User\Парсеры_161024\Архивные еженедельные отчёты.xlsx")
df_arch_otchets['Дата начала'] = pd.to_datetime(df_arch_otchets['Дата начала'])
min_year = df_arch_otchets['Дата начала'].dt.year.min()

#########################################
year_start, month_start = min_year, 1  ##
#########################################

df_seller_id = pd.read_excel(r"c:\Users\User\Парсеры_161024\seller_id.xlsx")
seller_id = df_seller_id['seller_id'][0]


def decoder_funct(filename, content):
    file_name = filename
    base64_content = content
    decoded_content = base64.b64decode(base64_content)

    output_directory = os.path.join(os.getcwd(), "FBS")
    os.makedirs(output_directory, exist_ok=True)

    new_file_name = os.path.splitext(file_name)[0] + ".xlsx"
    new_file_path = os.path.join(output_directory, new_file_name)

    with open(new_file_path, "wb") as new_file:
        new_file.write(decoded_content)

    print(f"Файл '{new_file_name}' успешно сохранен в папку '{output_directory}'\n")


def download_funct(date_start, date_end, seller_id, name):
    url = f"https://marketplace.wildberries.ru/ns/marketplace-app/marketplace-remote-wh/api/v3/portal/orders/archive/excel?dateFrom={date_start}&dateTo={date_end}&status=all&timezoneOffset=-7&type=all"

    header_value = f"BasketUID=d9b7509528994afd9e9f7416bd6ab9de; ___wbu=9d701a62-0fbe-4172-8a18-82d55c5d0962.1703672299; wb-pid=gYGeqrvXWLVKUbcNANNO2yONAAABkC9uSRUgkVneyuZwfD0yHkDbGKjUhjCJBNN6UFnh3J4w7SFNXw; external-locale=ru; _wbauid=5259289681721046210; wb-id=gYF3igJwl0JMFZTk9QxQHl_LAAABkUcrIhGXjbEY1n4NqYN01a7n4_H6bj-5342MYfUW80RsRaQIkmYzZmMzYWUzLTU1YTEtNDM3ZS1hNmExLWMwYmVlNjE0ZGZkOA; x-supplier-id-external={seller_id}; wbx-validation-key=34921d56-55ae-4bd5-958c-acb5f8d694df; WBTokenV3=eyJhbGciOiJSUzI1NiIsInR5cCI6IkpXVCJ9.eyJpYXQiOjE3MjkwNzU1NzQsInZlcnNpb24iOjIsInVzZXIiOiIxNDU4NTk5OSIsInNoYXJkX2tleSI6IjIiLCJjbGllbnRfaWQiOiJzZWxsZXItcG9ydGFsIiwic2Vzc2lvbl9pZCI6IjFhMmU1NTg1MDg5ZjRmMzFiNjkyNGJkYzI5NGFjNjFkIiwidXNlcl9yZWdpc3RyYXRpb25fZHQiOjE2NzE3MzIyNjMsInZhbGlkYXRpb25fa2V5IjoiNzA4ZDExNDQ4MWZhMGI2N2FlYmU0NmI0OTJjMThiZTIyODc3OWM1YmQxMTM3MjUwODFiNGFkN2Y2MzBkZGQ5MyJ9.AYSGY132eBxALObDxq8Ns1Vhi0SQKecyyKp6_wAYeRA0xA_6kf7hPoBEjmWHhL-GOqF3NKK7GDj8ho07QrYVOvgCj1nhTWFgex3Gi9GPmFxr0-qFc3XPA8bMGdZXN35wIQ1yP9JObQFP-Kb7xrsuq878svKnlIeMvTgMEyWCeQ1iXDOvSjESl6NONTRWBswZ7InhW7AQRbXs2WRNtJlxBLHy1RXc2PWEiBeugxWeYTG9JeUAiZnpCFQX1iCuSAGEmNAtvwAoT4U4L7PkWigYG0Kz1JytLkZx1Qak_zhTMdaP233t5IAf50WtnrzrnyLlXGkpffboEXDFEfstvg2J_Q; __zzatw-wb=MDA0dC0cTHtmcDhhDHEWTT17CT4VHThHKHIzd2UqP2olZFBiKDVRP0FaW1Q4NmdBEXUmCQg3LGBwVxlRExpceEdXeiwbGndyKFQQD2RCSWllbQwtUlFRS19/Dg4/aU5ZQ11wS3E6EmBWGB5CWgtMeFtLKRZHGzJhXkZpdRVQOA0YQkZ1eClDblNjfVwgdVtWeylLRTJtLFM4PmE+dV9vG3siXyoIJGM1Xz9EaVhTMCpYQXt1J3Z+KmUzPGwfaEdgJEheUQoqIQ1pN2wXPHVlLwkxLGJ5MVIvE0tsP0caRFpbQDsyVghDQE1HFF9BWncyUlFRS2EQR0lrZU5TQixmG3EVTQgNND1aciIPWzklWAgSPwsmIBMIaypUDBRfRkZ2bxt/Nl0cOWMRCxl+OmNdRkc3FSR7dSYKCTU3YnAvTCB7SykWRxsyYV5GaXUVTzg/W292JjEmPmlPaERdUENcSgpWTRd0JCpWOD5iQkEnMS1DVxlRDxZhDhYYRRcje0I3Yhk4QhgvPV8/YngiD2lIYCFMVVR+KiIWeG8oS3FPLH12X30beylOIA0lVBMhP05yqwtA/Q==; cfidsw-wb=u+ixTpo2ajfeRavcj985+bNEGpeiDgscD7SUa5d46DRE6DAZQkve+4uvTS66yVieOwKmnlJRk4fvlBxpCzvBB+gQh6bLrLpmpY3hDtEfisIY6Y2nBJk93Jehn1db6fGzhmfXFzpCbn1mP2adEHu+L0/kpUvoR1D1ozu4MBkI3w=="
    headers = {
        "Content-Type": "application/json",
        "User-Agent": "Chrome/115.0.0.0",
        "Accept": "*/*",
        "Accept-Encoding": "gzip, deflate, br",
        "Connection": "keep-alive",
        "cookie": f"{header_value}"
    }

    response = requests.get(url, headers=headers)
    response_result = response.json()
    #     print(response_result)

    if response_result['error'] != False:
        print(response_result['errorText'])
    else:

        file = response_result['data']['file']
        decoder_funct(name, file)


def get_month_start_end(year, month):
    start_date = datetime(year, month, 1)
    if month == 12:
        end_date = datetime(year + 1, 1, 1) - timedelta(days=1)
    else:
        end_date = datetime(year, month + 1, 1) - timedelta(days=1)
    return start_date, end_date

print('init')

current_date = datetime.now()
year, month = year_start, month_start

while year < current_date.year or (year == current_date.year and month <= current_date.month):
    if year == current_date.year and month == current_date.month:
        start_date = datetime(year, month, 1)
        end_date = current_date
    else:
        start_date, end_date = get_month_start_end(year, month)

    print(f"Начало: {start_date.strftime('%d.%m.%Y')}, Конец: {end_date.strftime('%d.%m.%Y')}")

    download_funct(start_date.strftime('%d.%m.%Y'), end_date.strftime('%d.%m.%Y'), seller_id,
                   start_date.strftime('%Y_%m'))

    if month == 12:
        year += 1
        month = 1
    else:
        month += 1

print('fin')

