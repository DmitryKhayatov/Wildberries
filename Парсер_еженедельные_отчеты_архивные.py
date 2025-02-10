import base64
import zipfile
from io import BytesIO
import os
import requests
import json
import pandas as pd
import time

#df_otchets = pd.read_excel("Еженедельные отчёты.xlsx")
df_arch_otchets = pd.read_excel(r"c:\Users\User\Парсеры_161024\Архивные еженедельные отчёты.xlsx")
df_seller_id = pd.read_excel(r"c:\Users\User\Парсеры_161024\seller_id.xlsx")
seller_id = df_seller_id['seller_id'][0]

def decoder_funct(filename, content):

    file_name = filename
    base64_content = content
    decoded_content = base64.b64decode(base64_content)
    zip_file = BytesIO(decoded_content)

    # Распаковка ZIP
    with zipfile.ZipFile(zip_file, 'r') as zip_ref:
        extracted_file_name = zip_ref.namelist()[0]
        extracted_file_content = zip_ref.read(extracted_file_name)

    output_directory = os.path.join(os.getcwd(), "Детали архивные")
    os.makedirs(output_directory, exist_ok=True)
    new_file_name = os.path.splitext(file_name)[0] + ".xlsx"
    new_file_path = os.path.join(output_directory, new_file_name)

    with open(new_file_path, "wb") as new_file:
        new_file.write(extracted_file_content)
    print(f"Распакованный файл успешно сохранен как '{new_file_path}'.")

### Запрос данных вб

for otchet_arch_num in df_arch_otchets['№ отчета']:
    print(otchet_arch_num)

    url = f"https://seller-weekly-report.wildberries.ru/ns/realization-reports/suppliers-portal-analytics/api/v1/reports/{otchet_arch_num}/details/archived-excel"
    #header_value = f"BasketUID=d9b7509528994afd9e9f7416bd6ab9de; ___wbu=9d701a62-0fbe-4172-8a18-82d55c5d0962.1703672299; wb-pid=gYGeqrvXWLVKUbcNANNO2yONAAABkC9uSRUgkVneyuZwfD0yHkDbGKjUhjCJBNN6UFnh3J4w7SFNXw; wbx-validation-key=97547a22-46a4-497a-8fc7-32df72cd2c6d; external-locale=ru; _wbauid=5259289681721046210; WBTokenV3=eyJhbGciOiJSUzI1NiIsInR5cCI6IkpXVCJ9.eyJpYXQiOjE3MjEyMDE0NzMsInZlcnNpb24iOjIsInVzZXIiOiIxNDU4NTk5OSIsInNoYXJkX2tleSI6IjIiLCJjbGllbnRfaWQiOiJzZWxsZXItcG9ydGFsIiwic2Vzc2lvbl9pZCI6IjFhMmU1NTg1MDg5ZjRmMzFiNjkyNGJkYzI5NGFjNjFkIiwidXNlcl9yZWdpc3RyYXRpb25fZHQiOjE2NzE3MzIyNjMsInZhbGlkYXRpb25fa2V5IjoiODkxNzU3YTFmYWIwMDkyODU3NDc2NzEwNDZjY2E4MmJhMzM1ZWIxMjE5ZjBlZGE1ODgxZDk2YjM3MzUzMzM2NSIsInBob25lIjoiIn0.OWj-o7XwrOLxvqUxHlBb2KOwByH-w7TnxrKCWZyDjFqvHiSSOt5Bo4XEZx4YUuraQwpnyPHU1HEUJv06coF6kVVDSp_wTLco1hWKjnxdxOpop5OZbMNb1lJ-21Nhp9gquLW-Rnaeyc35oibuRcDoB2R1rzRCW9_sw6ApTYKnUBk1eMpYjK8RRr_M94qUCle4WML2ct0ASZLqbNdwJP2AhczMn3EOVAfHn9esGrehGNPPNDUqs3rx5X-WGEzuaI7V09ws3J5onNP29fVOX98Rc94qPyCdfF0DOP6Ek11BD_MR7u6UIXU2e2SJ88LDD3EqpjI-3-Al3ZrTg_4CCv2Xbg; x-supplier-id-external={seller_id}; __zzatw-wb=MDA0dC0cTHtmcDhhDHEWTT17CT4VHThHKHIzd2UqP2olZFBiKDVRP0FaW1Q4NmdBEXUmCQg3LGBwVxlRExpceEdXeiwbEn1tLE8MDmE/QWllbQwtUlFRS19/Dg4/aU5ZQ11wS3E6EmBWGB5CWgtMeFtLKRZHGzJhXkZpdRVQOA0YQkZ1eClDblNjfVwgdVtWeylLRTJtLFM4PmE+dV9vG3siXyoIJGM1Xz9EaVhTMCpYQXt1J3Z+KmUzPGwfYE1bKENaUH8nGg1pN2wXPHVlLwkxLGJ5MVIvE0tsP0caRFpbQDsyVghDQE1HFF9BWncyUlFRS2EQR0lrZU5TQixmG3EVTQgNND1aciIPWzklWAgSPwsmIBN4cSVYfxBeQ0Nubxt/Nl0cOWMRCxl+OmNdRkc3FSR7dSYKCTU3YnAvTCB7SykWRxsyYV5GaXUVTzkOYm5CJ3wmQWwiHEReVHlcSjJYTBJ0JSMNPA1gcUIpLyo/VxlRDxZhDhYYRRcje0I3Yhk4QhgvPV8/YngiD2lIYCFEW08KJR4Uf28rS3FPLH12X30beylOIA0lVBMhP05yXg8AGw==; cfidsw-wb=6UMgP4PEH0A9YAuz3ajxstHudxna9XdMHdX56Gmt0/bAfgX1zyswyOfumtXjac22XqoFOZFVH3lWZcH6TFHch4zvmAqlTI44qnkyn6/AKdZ1205SmY/FWdDy1irFdH/PTJmpjEUw3A60kLwz5ecuKPArV8CBW+naodr140Lp"
    # header_value = f"BasketUID=d9b7509528994afd9e9f7416bd6ab9de; ___wbu=9d701a62-0fbe-4172-8a18-82d55c5d0962.1703672299; wb-pid=gYGeqrvXWLVKUbcNANNO2yONAAABkC9uSRUgkVneyuZwfD0yHkDbGKjUhjCJBNN6UFnh3J4w7SFNXw; external-locale=ru; _wbauid=5259289681721046210; wb-id=gYF3igJwl0JMFZTk9QxQHl_LAAABkUcrIhGXjbEY1n4NqYN01a7n4_H6bj-5342MYfUW80RsRaQIkmYzZmMzYWUzLTU1YTEtNDM3ZS1hNmExLWMwYmVlNjE0ZGZkOA; wbx-validation-key=98594e2b-33c0-4905-bc17-b9c47cb0cd95; WBTokenV3=eyJhbGciOiJSUzI1NiIsInR5cCI6IkpXVCJ9.eyJpYXQiOjE3MjM3OTQ3MzEsInZlcnNpb24iOjIsInVzZXIiOiIxNDU4NTk5OSIsInNoYXJkX2tleSI6IjIiLCJjbGllbnRfaWQiOiJzZWxsZXItcG9ydGFsIiwic2Vzc2lvbl9pZCI6IjFhMmU1NTg1MDg5ZjRmMzFiNjkyNGJkYzI5NGFjNjFkIiwidXNlcl9yZWdpc3RyYXRpb25fZHQiOjE2NzE3MzIyNjMsInZhbGlkYXRpb25fa2V5IjoiMjQ3ZjhiYjUyNWI3YzI4MTQ0NTRlZTY5MDkyMzQ3OGNhN2UyZGY3ODhlYzg3ZGE2OTU1NmM0ZDk5M2IxNmE2MCIsInBob25lIjoiIn0.mpy9jyLAUWgEZrFYQ0Xav3LwJ8Y9ZNjoU6CdGDAmV2fBeu_yzIPswO5k_6KjIxvk5521hcfW1RjOAtaicISWMr4koxEuli6g-Ry1tou-0TCJJQYt0Ofemmn0CYP9cai4M2zo0TEmpj5HvLYpUiH2bpa5E5tmezMD94tJx7DEePchWyFLjJ2OQJ687HrFruQo-d9hmWIwEICW1bQtifHomYw3pbOBHJshP5DNWYPdRXb50JMDE9i3LUBmYslJD-tqiP0LkRVlGQlrttHz4uXeqDEyr9E5AUCaO2eYZSsVnISEkQVc6gxrg9bRql5lZ_cnkyqmxWHoiS0quw_wm5HxmQ; x-supplier-id-external={seller_id}; __zzatw-wb=MDA0dC0cTHtmcDhhDHEWTT17CT4VHThHKHIzd2UqP2olZFBiKDVRP0FaW1Q4NmdBEXUmCQg3LGBwVxlRExpceEdXeiwbFXdwKFYLE2BESmllbQwtUlFRS19/Dg4/aU5ZQ11wS3E6EmBWGB5CWgtMeFtLKRZHGzJhXkZpdRVQOA0YQkZ1eClDblNjfVwgdVtWeylLRTJtLFM4PmE+dV9vG3siXyoIJGM1Xz9EaVhTMCpYQXt1J3Z+KmUzPGwfY0deJEpZVX4tGQ1pN2wXPHVlLwkxLGJ5MVIvE0tsP0caRFpbQDsyVghDQE1HFF9BWncyUlFRS2EQR0lrZU5TQixmG3EVTQgNND1aciIPWzklWAgSPwsmIBN7ayhUDg9jQklubxt/Nl0cOWMRCxl+OmNdRkc3FSR7dSYKCTU3YnAvTCB7SykWRxsyYV5GaXUVUz0/GERFdnYmPyNTaEReH0pcSgkoGhZ0dFUICg5bQHImLi1uVxlRDxZhDhYYRRcje0I3Yhk4QhgvPV8/YngiD2lIYCFHVVJ+LSAWe20nS3FPLH12X30beylOIA0lVBMhP05yucHJMA==; cfidsw-wb=gJVjzr9L7zYR33X+kg9h5qNMMjNBQp332CE1YKyZqIoOwgySOag8bYjPKbxPZHlMvd/Qdq/toMXqLeL/nbkmiSQw36uTlIgKCYMFwhErPtt/BUbPTNaNC5hUffDg2R6YqByq+xUiwmwngR+hdmT6Ox7JlGruJxpjkXyWuPiOEQ=="
    header_value = f"BasketUID=d9b7509528994afd9e9f7416bd6ab9de; ___wbu=9d701a62-0fbe-4172-8a18-82d55c5d0962.1703672299; wb-pid=gYGeqrvXWLVKUbcNANNO2yONAAABkC9uSRUgkVneyuZwfD0yHkDbGKjUhjCJBNN6UFnh3J4w7SFNXw; external-locale=ru; _wbauid=5259289681721046210; wb-id=gYF3igJwl0JMFZTk9QxQHl_LAAABkUcrIhGXjbEY1n4NqYN01a7n4_H6bj-5342MYfUW80RsRaQIkmYzZmMzYWUzLTU1YTEtNDM3ZS1hNmExLWMwYmVlNjE0ZGZkOA; x-supplier-id-external={seller_id}; wbx-validation-key=34921d56-55ae-4bd5-958c-acb5f8d694df; WBTokenV3=eyJhbGciOiJSUzI1NiIsInR5cCI6IkpXVCJ9.eyJpYXQiOjE3MjkwNzU1NzQsInZlcnNpb24iOjIsInVzZXIiOiIxNDU4NTk5OSIsInNoYXJkX2tleSI6IjIiLCJjbGllbnRfaWQiOiJzZWxsZXItcG9ydGFsIiwic2Vzc2lvbl9pZCI6IjFhMmU1NTg1MDg5ZjRmMzFiNjkyNGJkYzI5NGFjNjFkIiwidXNlcl9yZWdpc3RyYXRpb25fZHQiOjE2NzE3MzIyNjMsInZhbGlkYXRpb25fa2V5IjoiNzA4ZDExNDQ4MWZhMGI2N2FlYmU0NmI0OTJjMThiZTIyODc3OWM1YmQxMTM3MjUwODFiNGFkN2Y2MzBkZGQ5MyJ9.AYSGY132eBxALObDxq8Ns1Vhi0SQKecyyKp6_wAYeRA0xA_6kf7hPoBEjmWHhL-GOqF3NKK7GDj8ho07QrYVOvgCj1nhTWFgex3Gi9GPmFxr0-qFc3XPA8bMGdZXN35wIQ1yP9JObQFP-Kb7xrsuq878svKnlIeMvTgMEyWCeQ1iXDOvSjESl6NONTRWBswZ7InhW7AQRbXs2WRNtJlxBLHy1RXc2PWEiBeugxWeYTG9JeUAiZnpCFQX1iCuSAGEmNAtvwAoT4U4L7PkWigYG0Kz1JytLkZx1Qak_zhTMdaP233t5IAf50WtnrzrnyLlXGkpffboEXDFEfstvg2J_Q; __zzatw-wb=MDA0dC0cTHtmcDhhDHEWTT17CT4VHThHKHIzd2UqP2olZFBiKDVRP0FaW1Q4NmdBEXUmCQg3LGBwVxlRExpceEdXeiwbGndyKFQQD2RCSWllbQwtUlFRS19/Dg4/aU5ZQ11wS3E6EmBWGB5CWgtMeFtLKRZHGzJhXkZpdRVQOA0YQkZ1eClDblNjfVwgdVtWeylLRTJtLFM4PmE+dV9vG3siXyoIJGM1Xz9EaVhTMCpYQXt1J3Z+KmUzPGwfaEdgJEheUQoqIQ1pN2wXPHVlLwkxLGJ5MVIvE0tsP0caRFpbQDsyVghDQE1HFF9BWncyUlFRS2EQR0lrZU5TQixmG3EVTQgNND1aciIPWzklWAgSPwsmIBMIaypUDBRfRkZ2bxt/Nl0cOWMRCxl+OmNdRkc3FSR7dSYKCTU3YnAvTCB7SykWRxsyYV5GaXUVTzg/W292JjEmPmlPaERdUENcSgpWTRd0JCpWOD5iQkEnMS1DVxlRDxZhDhYYRRcje0I3Yhk4QhgvPV8/YngiD2lIYCFMVVR+KiIWeG8oS3FPLH12X30beylOIA0lVBMhP05yqwtA/Q==; cfidsw-wb=u+ixTpo2ajfeRavcj985+bNEGpeiDgscD7SUa5d46DRE6DAZQkve+4uvTS66yVieOwKmnlJRk4fvlBxpCzvBB+gQh6bLrLpmpY3hDtEfisIY6Y2nBJk93Jehn1db6fGzhmfXFzpCbn1mP2adEHu+L0/kpUvoR1D1ozu4MBkI3w=="
    headers = {
        "cookie": f"{header_value}"
    }

    response = requests.get(url=url, headers=headers)
    while response.status_code == 429:
        print(response.status_code, response.content)
        time.sleep(60)
        response = requests.get(url=url, headers=headers)
    while response.status_code == 500:
        print(response.status_code, response.content)
        time.sleep(300)
        response = requests.get(url=url, headers=headers)
    response_result = response.json()
    name = response_result['data']['name']
    file = response_result['data']['file']
    decoder_funct(name, file)

print('fin')