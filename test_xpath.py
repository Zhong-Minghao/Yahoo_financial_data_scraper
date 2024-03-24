import requests
from lxml import html
import math
import pandas as pd

stock = 'NVDA'

summary_url = f'https://finance.yahoo.com/quote/'+stock
user_agent = {'User-Agent':'Mozilla/5.0 (Windows NT 10.0; WOW64; rv:77.0) Gecko/20100101 Firefox/77.0'}

payload = {'p' : stock}

response = requests.get(summary_url, params=payload, headers=user_agent)

data = html.fromstring(response.text)



profile_url = f'https://finance.yahoo.com/quote/{stock}/profile'
payload = {'p' : stock}
response = requests.get(profile_url, params=payload, headers=user_agent)
data = html.fromstring(response.text)

profile_data = data.xpath('//*[@class="Mb(25px)"]/p[2]/span')

profile_data_str = [profile.text_content() for profile in profile_data]

df_i = pd.DataFrame([profile_data_str[0::2], profile_data_str[1::2]])

profile_data_df = df_i.transpose()



statistics_url = f'https://finance.yahoo.com/quote/{stock}/profile'
response = requests.get(statistics_url, params=payload, headers=user_agent)
data = html.fromstring(response.text)
profile_data = data.xpath('//*[@class="Mb(25px)"]/p[2]/span')
profile_data_str = [profile.text_content() for profile in profile_data]
df_i = pd.DataFrame([profile_data_str[0::2], profile_data_str[1::2]])
profile_data_df = df_i.transpose()



statistics_url = f'https://finance.yahoo.com/quote/{stock}/key-statistics'
response = requests.get(statistics_url, params=payload, headers=user_agent)
data = html.fromstring(response.text)
valuation_measures_headers = data.xpath('//*[@class="Pt(20px)"]/span')
header_row = [header.text_content() for header in valuation_measures_headers]
header_row.pop(0)
header_row.insert(0, 'Measure')
header_row.insert(1, 'Current')

financial_highlights_data = data.xpath('//*[@class="Fl(start) W(50%) smartphone_W(100%)"]/div/div/div/div/table/tbody/tr/td')
data_uf = [data.text_content() for data in financial_highlights_data]
data_array = [data_uf[i:i+2] for i in range(0, int(len(data_uf)),2)]

valuation_measures_df = pd.DataFrame(data_array, columns=['Measure', 'Current'])

# else
valuation_measures_data = data.xpath('//*[@class="Fl(start) smartphone_W(100%) W(50%)"]/div/div/div/div/table/tbody/tr/td')
data_uf = [data.text_content() for data in valuation_measures_data]
data_array = [data_uf[i:i+2] for i in range(0, int(len(data_uf)),2)]

valuation_measures_df = pd.DataFrame(data_array, columns=['Measure', 'Current'])



income_statement_url = f'https://finance.yahoo.com/quote/{stock}/financials'
payload = {'p' : stock}
response = requests.get(income_statement_url, params=payload, headers=user_agent)
data = html.fromstring(response.text)

income_statement_headers = data.xpath('//*[@class="D(tbhg)"]/div[1]/div')
income_statement_rows = data.xpath('//*[@data-test="fin-row"]/div[1]/div')

header_row = [header.text_content() for header in income_statement_headers]
data_uf = [data.text_content() for data in income_statement_rows]

j = len(data.xpath('//*[@class="D(tbr) C($primaryColor)"]/div'))
data_array = [data_uf[j*i:j*i+j] for i in range(0,math.ceil(len(data_uf)/j))]

income_statement_df = pd.DataFrame(data_array, columns=header_row)



balance_sheet_url = f'https://finance.yahoo.com/quote/{stock}/balance-sheet'
payload = {'p' : stock}
response = requests.get(balance_sheet_url, params=payload, headers=user_agent)
data = html.fromstring(response.text)

balance_sheet_headers = data.xpath('//div[@class="D(tbhg)"]/div[1]/div')
balance_sheet_data = data.xpath('//div[@data-test="fin-row"]/div[1]/div')

header_row = [header.text_content() for header in balance_sheet_headers]
data_uf = [data.text_content() for data in balance_sheet_data]

j = len(data.xpath('//*[@class="D(tbr) C($primaryColor)"]/div'))
data_array = [data_uf[j*i:j*i+j] for i in range(0,math.ceil(len(data_uf)/j))]

balance_sheet_df = pd.DataFrame(data_array, columns=header_row)


file_name = f'/{stock}_financial_data.xlsx'
writer = pd.ExcelWriter(file_name, engine='xlsxwriter')
balance_sheet_df.to_excel(writer, sheet_name='Balance_Sheet', index=False)
writer._save()
writer.close()