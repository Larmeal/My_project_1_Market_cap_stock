from selenium.common.exceptions import NoSuchElementException
from selenium import webdriver
import time
import pandas as pd
import random

class StockMarketcap:

    def __init__(self, name_sector):
        self.chrome_driver = "C:\Developer\chromedriver.exe"
        self.driver = webdriver.Chrome(executable_path=self.chrome_driver)
        self.driver.get(f"https://marketdata.set.or.th/mkt/sectorquotation.do?language=th&country=TH&market=SET&sector={name_sector}")
        self.driver.maximize_window()

        self.link_bank = []
        self.name_stock = []
        self.name_stock_max_to_min = []
        self.market_cap = []
        self.market_cap_test = []
        self.market_cap_max_to_min = []

    def save_link_of_stocks(self):
        time.sleep(2)
        self.row_of_table = self.driver.find_elements_by_xpath('//*[@id="maincontent"]/div/div[2]/div/div/div/div[3]/table/tbody/tr')
        for i in range(len(self.row_of_table)):
            self.symbol_sector = self.driver.find_element_by_xpath(f'//*[@id="maincontent"]/div/div[2]/div/div/div/div[3]/table/tbody/tr[{i + 1}]/td[1]/a')
            self.name_stock.append(self.symbol_sector.text)
            self.link_bank.append(self.symbol_sector.get_attribute("href"))

    def get_data_of_stock(self, name_sector):
        for i in self.name_stock:
            try:
                try:
                    self.driver.get(f"https://marketdata.set.or.th/mkt/stockquotation.do?symbol={i}&ssoPageId=1&language=th&country=TH")
                    time.sleep(2)
                    self.financial_statement = self.driver.find_element_by_xpath('//*[@id="maincontent"]/div/div[2]/div/ul/li[2]/a')
                    self.financial_statement.click()
                    time.sleep(1)
                    self.each_market_cap = self.driver.find_element_by_xpath('//*[@id="maincontent"]/div/div[4]/table/tbody[2]/tr[2]/td[6]').text
                    self.market_cap.append(float(self.each_market_cap.split()[0].replace(",", "")))
                    self.market_cap_test.append(float(self.each_market_cap.split()[0].replace(",", "")))

                except NoSuchElementException:
                    self.symbol_error = {
                        "Symbol": i,
                        "Sector": name_sector
                    }

                    self.df = pd.DataFrame(self.symbol_error)
                    self.df.to_csv("Symbol Error.csv", mode="a")

            except ValueError:
                    self.value_random = random.randint(1, 5)
                    self.market_cap.append(self.value_random)
                    self.market_cap_test.append(self.value_random)

    def switch_order(self):
        for i in range(len(self.market_cap)):
            self.max_market = max(self.market_cap_test)
            self.market_cap_max_to_min.append(self.max_market)
            self.market_cap_test.remove(self.max_market)

        for i in self.market_cap_max_to_min:
            self.stock_pick = self.name_stock[self.market_cap.index(i)]
            self.name_stock_max_to_min.append(self.stock_pick)

    def save_to_csv(self, name_sector):
        
        self.stock_dict = {
            "Stock": self.name_stock_max_to_min,
            "Market cap (*000)": self.market_cap_max_to_min
        }

        self.df = pd.DataFrame(self.stock_dict)
        self.df.to_csv(f"{name_sector}.csv", mode="w")

        self.driver.quit()

        

# ask_sectors = input("What is the sector that you want: ").upper()

all_stock = ["AGRI", "FOOD", "FASHION", "HOME", "PERSON", "BANK", "FIN", "INSUR", "AUTO", "IMM", "PAPER", "PETRO", "PKG", "STEEL", "CONMAT", "PROP", "PF%26REIT", "CONS", "ENERG", "MINE", "COMM", "HELTH", "MEDIA", "PROF", "TOURISM", "TRANS", "ETRON", "ICT"]

for i in all_stock:
    stock = StockMarketcap(i)
    stock.save_link_of_stocks()
    stock.get_data_of_stock(i)
    stock.switch_order()
    stock.save_to_csv(i)

