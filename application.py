import requests
from lxml import html
import pandas as pd
import xlsxwriter
import math
from tkinter import filedialog, Tk
from itertools import chain

class scrape_company_data():


    # --- Initialize Instance Variables --- #

    def __init__(self, stock, user_agent):
        self.stock = stock
        self.user_agent = user_agent

    # --- Scrape Company Summary Data --- #

    def scrape_summary_data(self):

        # Send a GET request and recieve the html response back
        summary_url = f'https://finance.yahoo.com/quote/{self.stock}'
        payload = {'p' : self.stock}
        response = requests.get(summary_url, params=payload, headers=self.user_agent)
        data = html.fromstring(response.text)

        # Get specific data from DOM elements using xpath
        company_name = data.xpath('//*[@id="quote-header-info"]/div[2]/div[1]/div[1]/h1')
        summary_data_headers = data.xpath('//*[@id="quote-summary"]/div/table/tbody/tr/td[1]')
        summary_data = data.xpath('//*[@id="quote-summary"]/div/table/tbody/tr/td[2]')

        # Data Processing
        company_name_list = [name.text_content() for name in company_name]
        company_name_str = company_name_list[0]
        summary_data_header_str  = [header.text_content() for header in summary_data_headers]
        summary_data_str  = [data.text_content() for data in summary_data]
        summary_data_header_str.insert(0, 'Company')
        summary_data_str.insert(0, company_name_str)

        df_i = pd.DataFrame([summary_data_str], columns=summary_data_header_str)

        summary_data_df = df_i.transpose()

        print(self.stock + ' Financial Summary Data Scraped')

        return summary_data_df


    # --- Scrape Company Profile Data --- #

    def scrape_profile_data(self):
        profile_url = f'https://finance.yahoo.com/quote/{self.stock}/profile'
        payload = {'p' : self.stock}
        response = requests.get(profile_url, params=payload, headers=self.user_agent)
        data = html.fromstring(response.text)

        profile_data = data.xpath('//*[@class="Mb(25px)"]/p[2]/span')

        profile_data_str = [profile.text_content() for profile in profile_data]

        df_i = pd.DataFrame([profile_data_str[0::2], profile_data_str[1::2]])

        profile_data_df = df_i.transpose()

        print(self.stock + ' Company Profile Data Scraped')

        return profile_data_df

    
    # --- Scrape Company Income Statement Data --- #
    
    def scrape_income_statement_data(self):
        income_statement_url = f'https://finance.yahoo.com/quote/{self.stock}/financials'
        payload = {'p' : self.stock}
        response = requests.get(income_statement_url, params=payload, headers=self.user_agent)
        data = html.fromstring(response.text)

        income_statement_headers = data.xpath('//*[@class="D(tbhg)"]/div[1]/div')
        income_statement_rows = data.xpath('//*[@data-test="fin-row"]/div[1]/div')

        header_row = [header.text_content() for header in income_statement_headers]
        data_uf = [data.text_content() for data in income_statement_rows]

        j = len(data.xpath('//*[@class="D(tbr) C($primaryColor)"]/div'))
        data_array = [data_uf[j*i:j*i+j] for i in range(0,math.ceil(len(data_uf)/j))]

        income_statement_df = pd.DataFrame(data_array, columns=header_row)

        print(self.stock + ' Income Statement Data Scraped')

        return income_statement_df

    
    # --- Scrape Company Balance Sheet Data --- #

    def scrape_balance_sheet_data(self):
        balance_sheet_url = f'https://finance.yahoo.com/quote/{self.stock}/balance-sheet'
        payload = {'p' : self.stock}
        response = requests.get(balance_sheet_url, params=payload, headers=self.user_agent)
        data = html.fromstring(response.text)

        balance_sheet_headers = data.xpath('//div[@class="D(tbhg)"]/div[1]/div')
        balance_sheet_data = data.xpath('//div[@data-test="fin-row"]/div[1]/div')

        header_row = [header.text_content() for header in balance_sheet_headers]
        data_uf = [data.text_content() for data in balance_sheet_data]

        j = len(data.xpath('//*[@class="D(tbr) C($primaryColor)"]/div'))
        data_array = [data_uf[j*i:j*i+j] for i in range(0,math.ceil(len(data_uf)/j))]

        balance_sheet_df = pd.DataFrame(data_array, columns=header_row)

        print(self.stock + ' Balance Sheet Data Scraped')

        return balance_sheet_df

    
    # --- Scrape Company Cash Flow Data --- #
    
    def scrape_cash_flow_data(self):
        balance_sheet_url = f'https://finance.yahoo.com/quote/{self.stock}/cash-flow'
        payload = {'p' : self.stock}
        response = requests.get(balance_sheet_url, params=payload, headers=self.user_agent)
        data = html.fromstring(response.text)

        cash_flow_headers = data.xpath('//div[@class="D(tbhg)"]/div[1]/div')
        cash_flow_data = data.xpath('//div[@data-test="fin-row"]/div[1]/div')

        header_row = [header.text_content() for header in cash_flow_headers]
        data_uf = [data.text_content() for data in cash_flow_data]

        j = len(data.xpath('//*[@class="D(tbr) C($primaryColor)"]/div'))
        data_array = [data_uf[j*i:j*i+j] for i in range(0,math.ceil(len(data_uf)/j))]

        cash_flow_df = pd.DataFrame(data_array, columns=header_row)

        print(self.stock + ' Cash Flow Data Scraped')

        return cash_flow_df


    # --- Scrape Company Statistics Data --- #

    def scrape_valuation_measures_data(self):

        statistics_url = f'https://finance.yahoo.com/quote/{self.stock}/key-statistics'
        payload = {'p' : self.stock}
        response = requests.get(statistics_url, params=payload, headers=self.user_agent)
        data = html.fromstring(response.text)

        # valuation_measures_headers = data.xpath('//*[@class="Bdtw(0px) C($primaryColor)"]/th/span')   #
        # valuation_measures_data = data.xpath('//*[@class="W(100%) Bdcl(c)  M(0) Whs(n) BdEnd Bdc($seperatorColor) D(itb)"]/tbody/tr/td')
        valuation_measures_data = data.xpath('//*[@class="Fl(start) smartphone_W(100%) W(50%)"]/div/div/div/div/table/tbody/tr/td')
        # header_row = [header.text_content() for header in valuation_measures_headers]
        data_uf = [data.text_content() for data in valuation_measures_data]
        data_array = [data_uf[i:i + 2] for i in range(0, int(len(data_uf)), 2)]
        # header_row.pop(0)
        # header_row.insert(0, 'Measure')
        # header_row.insert(1, 'Current')

        # j = len(valuation_measures_headers) + 1
        # data_array = [data_uf[j*i:j*i+j] for i in range(0,math.ceil(len(data_uf)/j))]

        # valuation_measures_df = pd.DataFrame(data_array, columns=header_row)
        valuation_measures_df = pd.DataFrame(data_array, columns=['Measure', 'Current'])

        print(self.stock + ' Valuation Measures Data Scraped')

        return valuation_measures_df
    
    # --- Scrape Financial Highlights and Trading Data --- #

    def scrape_highlights_and_trading_data(self):

        statistics_url = f'https://finance.yahoo.com/quote/{self.stock}/key-statistics'
        payload = {'p' : self.stock}
        response = requests.get(statistics_url, params=payload, headers=self.user_agent)
        data = html.fromstring(response.text)

        trading_information_data = data.xpath('//*[@class="Pstart(20px) smartphone_Pstart(0px)"]/div')
        financial_highlights_data = data.xpath('//*[@class="Mb(10px) Pend(20px) smartphone_Pend(0px)"]/div')
        trading_data_labels_uf = [data.xpath('./div/div/table//tr/td[1]/span/text()') for data in trading_information_data]

        trading_data_uf = [data.xpath('.//*[@class="W(100%) Bdcl(c) "]/tbody/tr/td[2]/descendant-or-self::text()') for data in trading_information_data]
        financial_data_labels_uf = [data.xpath('./div/div/table//tr/td[1]/span/text()') for data in financial_highlights_data]

        # Forgive me for this abomination of code you're about to witness
        for i in range(len(financial_data_labels_uf)):

            if i == 1:
                financial_data_labels_uf[i][1] += ' (ttm)'

            elif i == 2:
                financial_data_labels_uf[i][0] += ' (ttm)'
                financial_data_labels_uf[i][1] += ' (ttm)'

            elif i == 3:
                financial_data_labels_uf[i][0] += ' (ttm)'
                financial_data_labels_uf[i][1] += ' (ttm)'
                financial_data_labels_uf[i][2] += ' (yoy)'
                financial_data_labels_uf[i][3] += ' (ttm)'
                financial_data_labels_uf[i][5] += ' (ttm)'
                financial_data_labels_uf[i][6] += ' (ttm)'
                financial_data_labels_uf[i][7] += ' (yoy)'

            elif i == 4:
                financial_data_labels_uf[i][0] += ' (mrq)'
                financial_data_labels_uf[i][1] += ' (mrq)'
                financial_data_labels_uf[i][2] += ' (mrq)'
                financial_data_labels_uf[i][3] += ' (mrq)'
                financial_data_labels_uf[i][4] += ' (mrq)'
                financial_data_labels_uf[i][5] += ' (mrq)'

            elif i == 5:
                financial_data_labels_uf[i][0] += ' (ttm)'
                financial_data_labels_uf[i][1] += ' (ttm)'

        financial_data_uf = [data.xpath('.//*[@class="W(100%) Bdcl(c) "]/tbody/tr/td[2]/descendant-or-self::text()') for data in financial_highlights_data]
        merged_trading_data = list(zip(chain.from_iterable(trading_data_labels_uf), chain.from_iterable(trading_data_uf)))
        merged_financial_data = list(zip(chain.from_iterable(financial_data_labels_uf), chain.from_iterable(financial_data_uf)))

        trading_data_df = pd.DataFrame(merged_trading_data)
        financial_data_df = pd.DataFrame(merged_financial_data)

        print(self.stock + ' Highlights and Trading Data Scraped')

        return {'trading': trading_data_df, 'financial': financial_data_df}


class application_methods:

    @staticmethod
    def load_input_data():
        root = Tk()
        root.attributes("-topmost", True)
        root.withdraw()
        file_name = filedialog.askopenfilename(title='Select Input File')

        with open(file_name, 'r', encoding='utf8') as f:
            stocks = [line.strip() for line in f]

        root.destroy()
        
        return stocks
    
    @staticmethod
    def select_output_directory():
        root = Tk()
        root.attributes("-topmost", True)
        root.withdraw()
        directory_path = filedialog.askdirectory(title='Select Output Folder')

        return directory_path
    
    @staticmethod
    def write_xlsx_file(stock,
    output_directory,
    summary_data, 
    profile_data, 
    income_statement,
    balance_sheet, 
    cash_flow, 
    valuation_measures,
    highlight_data,
    trading_data):

        # Write each stock data into a seperate xlsx file
        file_name = output_directory + f'/{stock}_financial_data.xlsx'
        writer = pd.ExcelWriter(file_name, engine='xlsxwriter')

        summary_data.to_excel(writer, sheet_name='Summary', header=False)
        profile_data.to_excel(writer, sheet_name='Profile', index=False, header=False)
        income_statement.to_excel(writer, sheet_name='Income_Statement', index=False)
        balance_sheet.to_excel(writer, sheet_name='Balance_Sheet', index=False)
        cash_flow.to_excel(writer, sheet_name='Cash_Flow', index=False)
        valuation_measures.to_excel(writer, sheet_name='Valuation_Measures', index=False)
        highlight_data.to_excel(writer, sheet_name='Financial_Highlights', index=False, header=False)
        trading_data.to_excel(writer, sheet_name='Trading_Information', index=False, header=False)

        # --- Save Data --- #

        writer._save()


    

        