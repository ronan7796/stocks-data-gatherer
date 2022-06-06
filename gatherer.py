import time

import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

NEEDED_LABEL = ('Quý - tỷ đ', 'Doanh thu thuần', 'Lợi nhuận gộp', 'LN trước thuế', 'LN sau thuế',
                'P/E', 'EPS', 'ROE(%)', 'BBiên LN gộp(%)', 'Biên LNST(%)', 'Chỉ số tài chính', 'Năm - tỷ đ')

NEEDED_LABEL_BANK_STOCK = ('Năm - tỷ đ', 'Thu nhập lãi thuần', 'Lãi thuần từ HĐ dịch vụ', 'Lãi thuần từ HĐ đầu tư',
                           'Tổng thu nhập HĐ (TOI)',
                           'Chi phí dự phòng', 'Lợi nhuận trước dự phòng', 'LN trước thuế', 'LN sau thuế', 'P/E', 'EPS',
                           'Biên lãi thuần (NIM)(%)', 'Chỉ số tài chính', 'Quý - tỷ đ')

BANK_STOCK = ('ACB', 'BAB', 'BID', 'CTG', 'EIB', 'EVF', 'HDB', 'LPB', 'MBB',
              'MSB', 'NVB', 'OCB', 'SHB', 'SSB', 'STB', 'TCB', 'TPB', 'VCB',
              'VIB', 'VPB', 'NAB', 'ABB', 'KLB', 'BVB', 'PGB', 'SGB')

EXCESS_LABEL = ('1Y Hi/ Loinfo', 'VNIndex(%) info', 'Vốn hóa (tỷ)',
                'KLGD', 'P/Einfo', 'EV/EBITDAinfo', 'P/Binfo', 'Cổ tức info')

# Time wait for each page to load
TIME_WAIT = 2

# Excel writer init
writer = pd.ExcelWriter('report.xlsx')


def get_user_input():
    """Get user input on how many stock they want to gather data from"""
    number_of_tickets = int(input('How many stocks do you want data from: '))
    stock_symbols = []
    for i in range(number_of_tickets):
        stock = input(f'Enter symbol({i + 1}/{number_of_tickets}): ')
        stock_symbols.append(stock)
    return stock_symbols


def get_data(driver, needed_label=NEEDED_LABEL):
    """Get stock data from the website https://tcinvest.tcbs.com.vn"""
    label_elements_in_financial = driver.find_elements(By.CSS_SELECTOR, '.label')
    label_raw_text = [
        element.text for element in label_elements_in_financial if element.text]
    label = [element.replace('\ninfo\nbar_chart', '')
             for element in label_raw_text if
             element not in EXCESS_LABEL and 'Bạn nghĩ sao về' not in element and 'Giá\n' not in element]

    value_elements_in_financial = driver.find_elements(
        By.CSS_SELECTOR, '.value')
    value = [element.text.split(
        '\n') for element in value_elements_in_financial if '\n' in element.text]

    full_data = dict(zip(label, value))

    need_data = {key: full_data.get(key)
                 for key in full_data if key in needed_label}

    return pd.DataFrame.from_dict(need_data)


def main():
    stock_symbols = get_user_input()
    options = webdriver.ChromeOptions()
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    driver = webdriver.Chrome(
        r'C:\Users\Administrator\Downloads\New folder/chromedriver.exe', options=options)

    print('============START GETTING DATA============')

    # This maximize the driver window to see the result easier
    driver.maximize_window()

    for stock_symbol in stock_symbols:

        if stock_symbol not in BANK_STOCK:
            URL = f'https://tcinvest.tcbs.com.vn/tc-price/tc-analysis/financial?ticker={stock_symbol}'
            driver.get(URL)
            time.sleep(TIME_WAIT)

            df1 = get_data(driver)

            WebDriverWait(driver, 20).until(EC.element_to_be_clickable(
                (By.XPATH,
                 "//*[@id='analysis-right']/div[2]/div[2]/app-financial/app-analysis-finacial-tab/div/div[1]/div[2]"))).click()
            time.sleep(TIME_WAIT)

            df2 = get_data(driver)

            WebDriverWait(driver, 20).until(EC.element_to_be_clickable(
                (By.XPATH, '//*[@id="mat-select-0"]'))).click()
            time.sleep(0.2)
            WebDriverWait(driver, 20).until(EC.element_to_be_clickable(
                (By.XPATH, '/html/body/div[2]/div[2]/div/div/div/mat-option[2]'))).click()
            time.sleep(TIME_WAIT)

            df4 = get_data(driver)

            WebDriverWait(driver, 20).until(EC.element_to_be_clickable(
                (By.XPATH,
                 "//*[@id='analysis-right']/div[2]/div[2]/app-financial/app-analysis-finacial-tab/div/div[1]/div[1]"))).click()
            time.sleep(TIME_WAIT)

            df3 = get_data(driver)

            result = pd.concat([df1, df2, df3, df4], axis=1).transpose()
            result.to_excel(writer, sheet_name=f'{stock_symbol}')

            print(f'{stock_symbol} SHEET GENERATED')

        elif stock_symbol.upper() in BANK_STOCK:
            URL = f'https://tcinvest.tcbs.com.vn/tc-price/tc-analysis/financial?ticker={stock_symbol}'
            driver.get(URL)
            time.sleep(TIME_WAIT)

            df1 = get_data(driver, needed_label=NEEDED_LABEL_BANK_STOCK)

            WebDriverWait(driver, 20).until(EC.element_to_be_clickable(
                (By.XPATH,
                 "//*[@id='analysis-right']/div[2]/div[2]/app-financial/app-analysis-finacial-tab/div/div[1]/div[2]"))).click()
            time.sleep(TIME_WAIT)

            df2 = get_data(driver, needed_label=NEEDED_LABEL_BANK_STOCK)

            WebDriverWait(driver, 20).until(EC.element_to_be_clickable(
                (By.XPATH, '//*[@id="mat-select-0"]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 20).until(EC.element_to_be_clickable(
                (By.XPATH, '/html/body/div[2]/div[2]/div/div/div/mat-option[2]'))).click()
            time.sleep(TIME_WAIT)

            df4 = get_data(driver, needed_label=NEEDED_LABEL_BANK_STOCK)

            WebDriverWait(driver, 20).until(EC.element_to_be_clickable(
                (By.XPATH,
                 "//*[@id='analysis-right']/div[2]/div[2]/app-financial/app-analysis-finacial-tab/div/div[1]/div[1]"))).click()
            time.sleep(TIME_WAIT)

            df3 = get_data(driver, needed_label=NEEDED_LABEL_BANK_STOCK)

            result = pd.concat([df1, df2, df3, df4], axis=1).transpose()
            result.to_excel(writer, sheet_name=f'{stock_symbol}')

            print(f'{stock_symbol} SHEET GENERATED')

    writer.save()
    print('============FINISH============')


if __name__ == '__main__':
    main()
