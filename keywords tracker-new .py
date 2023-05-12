import undetected_chromedriver as uc
from selenium import webdriver
from bs4 import BeautifulSoup
from datetime import date
import time
import pandas as pd
import requests
import random
import os

# Get current working directory
# cwd = os.getcwd()

# # Define relative path to Excel file
# filename = 'Keywords SEO Program\keywords list.xlsx'
# filepath = os.path.join(cwd, filename)

# Define working directory
wd = r'Z:/業務客服部/M.業務客服部/M.行銷業務/行銷日常/Romulus/Keywords SEO Program/keywords list.xlsx'

# Read Keyword file
keywords = pd.read_excel(wd)

# No. of top N result extracting
n_result = 30

# initializing
se_results = []

# delay_choices = [2, 5, 3, 6, 7, 1]  #延遲的秒數
# delay = random.choice(delay_choices)  #隨機選取秒數
delay = random.choice(range(3,9))  #隨機選取秒數

# 表頭
headers = {
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9", 
    "Accept-Encoding": "gzip, deflate, br", 
    "Accept-Language": "zh-TW,zh;q=0.9", 
    "Host": "example.com",  #目標網站
    "Sec-Fetch-Dest": "document", 
    "Sec-Fetch-Mode": "navigate", 
    "Sec-Fetch-Site": "none", 
    "Upgrade-Insecure-Requests": "1", 
    "User-Agent": "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/85.0.4183.102 Safari/537.36", #使用者代理
    "Referer": "https://www.google.com/"  #參照位址
}

# user_agent = UserAgent()
user_agent = 'Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/110.0.0.0 Mobile Safari/537.36'
response = requests.get(url="https://example.com", headers={'User-Agent': user_agent})

# Loop over keyword list
for keyword in keywords['keywords']:
    # Start webdriver
    options = webdriver.ChromeOptions()
    options.add_argument('--ignore-certificate-errors')
    options.add_argument("--test-type")
    options.add_argument("--headless") #hidden browser
    options.add_argument('--lang=zh-TW')
    # options.add_experimental_option('excludeSwitches', ['enable-logging']) #To fix device event log error
    # options.use_chromium = True  #To fix device event log error
    # options.add_argument("--user-agent=iphone")
    # options.binary_location = r"C:/Program Files/Google/Chrome/Application/Chrome.exe"
    driver = uc.Chrome(options=options)


    time.sleep(delay)  #延遲

    n_rank = 0 #reset rank
    n_keyword = 0 #initialize keyword
    url = 'https://www.google.com.tw/search?num={}&q={}'.format(n_result,keyword)
    print('#{} --- {} --- {} ...'.format(n_keyword,keyword,url))
    driver.get(url)    
    page_source = driver.page_source
    soup = BeautifulSoup(page_source, 'lxml')
    
    #SEM class start with 'v5yQqb'  <--pls check if update needed
    #Organic class start with 'yuRUbf'   
    results_selector = soup.select('div[class*="v5yQqb"] , div[class*="yuRUbf"]')
    
    #Loop over the results
    for result_selector in results_selector:
        # Case when SEM / Organic Result
        if result_selector['class'][0].startswith('v5yQqb'):
            domain_name_class = 'x2VHCd OSrXXb qzEoUe'
            result_type = 'SEM'
            # domain_name = result_selector.select('span[class*="{}"]'.format(domain_name_class))[0].get_text()                
        else:
            domain_name_class = 'dyjrff qzEoUe'
            result_type = 'Organic'
            # domain_name = result_selector.select('cite[class*="{}"]'.format(domain_name_class))[0].get_text()    
        link = result_selector.select('a')[0]['href']
        
        n_rank += 1


        time.sleep(delay)  #延遲

        temp_dict = {
            'query_date' : date.today().strftime("%Y%m%d"),
            'keyword' : keyword,
            'rank' : n_rank,
            'result_type' : result_type,
            # 'domain_name' : domain_name,
            'link' : link
            }        
        se_results.append(temp_dict)

    time.sleep(delay)

    driver.close()
    driver.quit()

df_se_results = pd.DataFrame(se_results)

# Define relative path to CSV file
# filename_1 = 'keywords list.xlsx'
# filepath_1 = os.path.join(cwd, filename_1)

# 讀取keywords URL.csv和se_result_20230223.csv文件中的數據
# df1 = pd.read_csv(filepath_1, encoding='utf-8-sig')
df1 = pd.read_csv('Z:/業務客服部/M.業務客服部/M.行銷業務/行銷日常/Romulus/Keywords SEO Program/keywords URL.csv',encoding='utf-8-sig')
df2 = df_se_results

# 使用merge函數將兩個DataFrame對象按照第一列進行合併
df_merged = pd.merge(df1, df2, how='inner', left_on='keyword', right_on='keyword')
print(df_merged)
# 遍歷合併後的DataFrame對象，比對第二列和第六列的數據
result = []
for keyword, group in df_merged.groupby('keyword'):
    matched = False
    for index, row in group.iterrows():
        if row[df_merged.columns[1]] == row[df_merged.columns[5]]:
            result.append(row['rank'])
            matched = True
            break
    if not matched:
        result.append('>30')

# 將結果保存為一個新的DataFrame對象，並寫入到一個csv文件中
df_result = pd.DataFrame({'keywords': df_merged['keyword'].unique(), 'result': result})

# Export to csv
df_result.to_csv('Z:/業務客服部/M.業務客服部/M.行銷業務/行銷日常/Romulus/Keywords SEO Program/results/df_result_{}_{}.csv'.format(date.today().strftime("%Y_%m%d_%H%M%S")),encoding='utf-8-sig',index=False)