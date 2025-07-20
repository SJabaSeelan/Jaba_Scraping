# import json
# import requests
# import pandas as pd
#
# # url = "https://results.us.securityeducation.com/api/reporting/v0.1.0/phishing"
# url="https://results.us.securityeducation.com/api/reporting/v0.1.0/auhentication"
# # url4=  "https://results.us.securityeducation.com/api/reporting/v0.1.0/users
# my_headers = {'x-apikey-token' : 'YV5iYWWji13z25LX/ZAQOwmHKmHqdpQk8pQzIcSQ5hGl3cpog'}
# r = requests.get(url,headers=my_headers)
# d=r.json()
# print(d)
# # print(d['data'])
# # df=pd.DataFrame(d['data'])
# # print(df)

# This file will stract data from https://en.wikipedia.org/wiki/List_of_dinosaur_genera and then create a list with allthe names of dinosaurs.

import pandas as pd
pd.set_option('display.max_rows', None)
import requests
from bs4 import BeautifulSoup
import time
import numpy as np
from openpyxl.workbook import Workbook
import re

url = 'https://en.wikipedia.org/wiki/List_of_dinosaur_genera'
html = requests.get(url)
#print(html.text)
soup = BeautifulSoup(html.text,'html.parser')
#print(soup)
#print('#############################')
urls = soup.find_all('a', href=True)
#print(urls)
links_and_names = [(url['href'], url.text) for url in urls]
#print(links_and_names)
dino_data_clean = [links_and_names[link] for link in range(len(links_and_names)) if links_and_names[link][0].startswith("/wiki/")]
#print(dino_data_clean)
dino_data_clean = dino_data_clean[:2317:]
# print(dino_data_clean)
dino_df = pd.DataFrame(dino_data_clean, columns = ['url', 'dinosaur'])
#print(dino_df)
dino_df['dinosaur'] = dino_df['dinosaur'].replace('', np.nan)
#print(dino_df)
dino_df = dino_df.dropna(axis = 0, subset = ['dinosaur'])
#print(dino_df)
dino_data_clean = dino_df.set_index('url')['dinosaur'].to_dict()
#print(dino_data_clean)
dino_data = [('https://en.wikipedia.org'+ url, dinosaur) for url, dinosaur in dino_data_clean.items()]
# print(dino_data)
dino_data = dino_data[38::]
#print(dino_data)
dino_urls = [element for pair in dino_data for element in pair if element.startswith('https://en.wikipedia.org')]
#print(dino_urls)
dino_info = []
for url in range(200):
    html = requests.get(dino_urls[url])
    soup = BeautifulSoup(html.text, 'html.parser')
    paragraphs = soup.select('p')
    clean_paragraphs = [paragraph.text.strip() for paragraph in paragraphs]
    clean_paragraphs = clean_paragraphs[:4:]
    dino_info.append(''.join(clean_paragraphs))
    #progress += 1
dino_df = pd.DataFrame(dino_data, columns = ['URL','Dinosaur'])
#print(dino_df)
dino_details = pd.DataFrame(dino_info, columns = ['Info'])
# print(dino_details)
dino_df = pd.concat([dino_df, dino_details], ignore_index = True, axis = 1)
#print(dino_df)
file_name = r'C:\Users\Jaba\Desktop\DataEngineeeringDoc\Dino_data.xlsx'
dino_df.to_excel(file_name)
#print('DataFrame is written to Excel File successfully!')

dino_df= pd.read_excel(file_name)
dino_df.drop('Unnamed: 0', inplace=True, axis=1)
dino_df.columns = ['URLs', 'Dinosaur', 'Info']
dino_info = dino_df['Info'].to_dict()
# print(dino_info)
print(dino_info)
dino_info = dino_info.values()
heights = []
progress = 0

for element in dino_info:
    if re.findall('\d+\smeters', str(element)):
        height = re.findall('\d+\smeters', str(element))
        heights.append(height)

    else:
        heights.append(list('-'))
    progress += 1

heights_clean = []
for element in heights:
    for height in enumerate(element):
        if height[0] == 0:
            heights_clean.append(height[1])

weights = []
for element in dino_info:
    if re.findall('\d+\stonnes|\d+\skilograms', str(element)):
        weight = re.findall('\d+\stonnes|\d+\skilograms', str(element))
        weights.append(weight)
    else:
        weights.append(list('-'))

weights_clean = []
for element in weights:
    for weight in enumerate(element):
        if weight[0] == 0:
            weights_clean.append(weight[1])

dino_df.drop('Info', inplace=True, axis=1)
heights_df = pd.DataFrame(heights_clean, columns=['Height'])
weights_df = pd.DataFrame(weights_clean, columns=['Weight'])
dino_df = pd.concat([dino_df, heights_df, weights_df], ignore_index=True, axis=1)
print(dino_df)
dino_df.columns = ['URL', 'Dinosaur', 'Height', 'Weight']
print(dino_df)
dino_df.to_excel(file_name)
print('DataFrame successfully exported to CSV file!')
