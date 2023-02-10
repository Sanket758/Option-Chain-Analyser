from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import NoSuchElementException

import glob
import time
import re
import math
import os
import pandas as pd
import requests
from PIL import Image


class OptionChainAnalyzer:
    def __init__(self) -> None:
        self.options = Options()
        self.options.add_argument('--headless')
        user_agent = "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/109.0.0.0 Safari/537.36"
        self.options.add_argument('user-agent={0}'.format(user_agent))
        self.history = []
        self.cmp = 0

    def get_current_index_price(self):
        try:
            self.browser.implicitly_wait(30)
            current_index_price = self.browser.find_element(
                By.ID, "equity_underlyingVal")
            cmp = re.search('[0-9,.]+', current_index_price.text).group()
            self.cmp = cmp
            return cmp
        except NoSuchElementException:
            time.sleep(10)
            try:
                current_index_price = self.browser.find_element(
                    By.XPATH, '//*[@id="equity_underlyingVal"]')
                cmp = re.search('[0-9,.]+', current_index_price.text).group()
                self.cmp = cmp
                return cmp
            except NoSuchElementException:
                return self.cmp

    def refresh_chain(self):
        el = self.browser.find_element(
            By.XPATH, '//*[@id="optionchain_equity_sp"]/div/div/div[3]/div[1]/div[1]/span[3]/a'
        )
        el.click()
        print('Refreshed')

    def download_option_chain(self):
        # if chrome driver is not installed run following:
        # self.browser = webdriver.Chrome(ChromeDriverManager().install(), options=self.options)

        self.browser = webdriver.Chrome(options=self.options)
        self.browser.get("https://www.nseindia.com/option-chain")
        self.browser.maximize_window()
        # print(self.browser.current_url)
        self.browser.implicitly_wait(30)
        # To check if correct page is downloaded
        self.browser.save_screenshot('ss.png')

        # try:
        #     current_index_price = self.browser.find_element(
        #         By.ID, "equity_underlyingVal")
        # except NoSuchElementException:
        #     current_index_price = self.browser.find_element(
        #         By.ID, "equity_underlyingVal")
        # print(f"{current_index_price.text}")

        try:
            el = self.browser.find_element(By.ID, "downloadOCTable")
        except NoSuchElementException:
            time.sleep(10)
            el = self.browser.find_element(By.ID, "downloadOCTable")
        # print(el.text)
        el.click()
        time.sleep(2)
        downloaded_file = max(glob.glob('*.csv'), key=os.path.getctime)
        return downloaded_file

    @staticmethod
    def fix_commas(df):
        def replace_commas(text):
            try:
                return int(text.replace(',', ''))
            except ValueError:
                return int(0)

        df['ASK QTY'] = df['ASK QTY'].apply(lambda x: replace_commas(x))
        df['ASK QTY.1'] = df['ASK QTY.1'].apply(lambda x: replace_commas(x))

        df['BID QTY'] = df['BID QTY'].apply(lambda x: replace_commas(x))
        df['BID QTY.1'] = df['BID QTY.1'].apply(lambda x: replace_commas(x))

        df['OI'] = df['OI'].apply(lambda x: replace_commas(x))
        df['OI.1'] = df['OI.1'].apply(lambda x: replace_commas(x))

        df['VOLUME'] = df['VOLUME'].apply(lambda x: replace_commas(x))
        df['VOLUME.1'] = df['VOLUME.1'].apply(lambda x: replace_commas(x))

        def fix_underscore(text):
            if text == '-':
                return 0
            else:
                return replace_commas(text)

        df['CHNG IN OI'] = df['CHNG IN OI'].apply(lambda x: fix_underscore(x))
        df['CHNG IN OI.1'] = df['CHNG IN OI.1'].apply(lambda x: fix_underscore(x))
        return df

    def show_data(self):
        print("="*100)
        t = time.localtime()
        current_time = time.strftime("%H:%M:%S", t)

        print('='*22, current_time, '='*20)

        print('='*10, 'Call', '='*10, '='*10, 'Put', '='*10)
        print(
            f"{'Buying':<10} {self.df['BID QTY'].sum():<10} {'|':>6} {self.df['BID QTY.1'].sum():^25}")
        print(
            f"{'Selling':<10} {self.df['ASK QTY'].sum():<10} {'|':>6} {self.df['ASK QTY.1'].sum():^25}")

        print('='*9, 'OI Call', '='*9, '='*9, 'OI Put', '='*9)
        total_oi_call = self.df['OI'].sum()
        total_oi_put = self.df['OI.1'].sum()
        total_oi_call_volume = self.df['VOLUME'].sum()
        total_oi_put_volume = self.df['VOLUME.1'].sum()
        total_coi_call = self.df['CHNG IN OI'].astype(float).sum()
        total_coi_put = self.df['CHNG IN OI.1'].astype(float).sum()
        print(f"{'OI':<10} {total_oi_call:<10} {'|':>6} {total_oi_put:^25}")
        print(
            f"{'Volume':<10} {total_oi_call_volume:<10} {'|':>6} {total_oi_put_volume:^25}")
        print(f"{'ChngOI':<10} {total_coi_call:<10} {'|':>6} {total_coi_put:^25}")

        print('='*22, "PCR", '='*22)
        volume_pcr = total_oi_put_volume/total_oi_call_volume
        print(f"{'Volume PCR':<10} {volume_pcr:^40}")
        oi_pcr = total_oi_put/total_oi_call
        print(f"{'PCR':<10} {oi_pcr:^40}")

        print('='*22, "Top 5 Highest change in OI", '='*22)
        coi_call = self.df.sort_values(by="CHNG IN OI")
        coi_put = self.df.sort_values(by="CHNG IN OI.1")

        top_5_negative_call = coi_call.head()[["OI", "CHNG IN OI", "VOLUME", "STRIKE"]]
        top_5_positive_call = coi_call.tail()[["OI", "CHNG IN OI", "VOLUME", "STRIKE"]]

        top_5_negative_put = coi_put.head()[["OI.1", "CHNG IN OI.1", "VOLUME.1", "STRIKE"]]
        top_5_positive_put = coi_put.tail()[["OI.1", "CHNG IN OI.1", "VOLUME.1", "STRIKE"]]

        print('='*11, 'Call', '='*18, '='*15, 'Put', '='*15)
        for col in ["OI", "CHNG IN OI", "VOLUME", "STRIKE"] + list(reversed(["OI.1", "CHNG IN OI.1", "VOLUME.1"])):
            print(f"{col:^10}", end='')
        print()

        for row1, row2 in zip(top_5_negative_call.iterrows(), top_5_negative_put.iterrows()):
            for col in ["OI", "CHNG IN OI", "VOLUME", "STRIKE"]:
                print(f"{row1[1][col]:^10}", end='')
            for col in list(reversed(["OI.1", "CHNG IN OI.1", "VOLUME.1"])):
                print(f"{row2[1][col]:^10}", end='')
            print()
        print()
        for row1, row2 in zip(top_5_positive_call.iterrows(), top_5_positive_put.iterrows()):
            for col in ["OI", "CHNG IN OI", "VOLUME", "STRIKE"]:
                print(f"{row1[1][col]:^10}", end='')
            for col in list(reversed(["OI.1", "CHNG IN OI.1", "VOLUME.1"])):
                print(f"{row2[1][col]:^10}", end='')
            print()

        # @Todo: Add code for vwap
        # vwap = Sum of (Price∗Volume for each Trade) / Total Volume

        self.history.append({
            "time": current_time,
            "PCR": round(oi_pcr, 3),
            "Vol PCR": round(volume_pcr, 3),
            "OI Call": total_oi_call,
            "OI Put": total_oi_put,
            "OI Vol Call": total_oi_call_volume,
            "OI Vol Put": total_oi_put_volume,
            "Price": self.get_current_index_price()
        })

    def show_history(self):
        print("="*100)
        print(
            f"{'Time':^15}{'Call':^15}{'Put':^15}{'Diff':^15}{'PCR':^15}{'Signal':^15}{'Price':^15}"
        )
        for item in self.history:
            print(
                f"{item['time']:^15}{item['OI Call']:^15}{item['OI Put']:^15}{item['OI Call']-item['OI Put']:^15}{item['PCR']:^15}{'Sell' if float(item['PCR'])<0.7 else 'Buy':^15}{item['Price']:^15}"
            )

    def start(self):
        while True:
            try:
                downloaded_file = self.download_option_chain()
                # downloaded_file = "option-chain-ED-NIFTY-09-Feb-2023.csv"
                self.df = pd.read_csv(downloaded_file, skiprows=1).iloc[:, 1:-1]
                # Relevant strikes
                middle = self.df.shape[0]//2
                self.relevant_strikes = self.df.iloc[middle-15:middle+10, :]
                self.df = self.fix_commas(self.df)
                try:
                    self.show_data()
                except ZeroDivisionError:
                    # Sometime data is not scraped properly, so sleep and try again
                    time.sleep(10)
                    continue
                self.show_history()
                self.browser.quit()
                # Redo every five minutes
                time.sleep(180)
            except Exception as e:
                print(e)
                time.sleep(5)


autotrender = OptionChainAnalyzer()
autotrender.start()
