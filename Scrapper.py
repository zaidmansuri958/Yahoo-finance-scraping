import time
import numpy as np
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.options import Options


class StockScraper:
    def __init__(self, driver, timeout=20):
        self.driver = driver
        self.wait = WebDriverWait(self.driver, timeout=timeout)
        self.data = []

    def wait_for_page_load(self):
        page_title = self.driver.title

        try:
            self.wait.until(
                lambda d: d.execute_script(
                    "return document.readyState") == "complete"
            )
        except:
            print(
                f"The page {page_title} did not fully loaded within given duration")
        else:
            print(f"The page {page_title} is fully loaded")

    def access_url(self, url):
        self.driver.get(url)
        self.wait_for_page_load()

    def access_most_active_stocks(self):
        actions = ActionChains(self.driver)
        # hover to the market menu
        markets_menu = self.wait.until(
            EC.presence_of_element_located((
                By.XPATH, '/html[1]/body[1]/div[2]/header[1]/div[1]/div[1]/div[1]/div[4]/div[1]/div[1]/ul[1]/li[3]/a[1]/span[1]')
            ))
        actions.move_to_element(markets_menu).perform()

        # click on the trending tickers
        trending_tickers = self.wait.until(
            EC.element_to_be_clickable(
                (By.XPATH, "//div[contains(text(),'Trending Tickers')]"))
        )

        trending_tickers.click()
        self.wait_for_page_load()
        try:
            most_active = self.wait.until(
                EC.element_to_be_clickable(
                    (By.XPATH, '/html[1]/body[1]/div[2]/main[1]/section[1]/section[1]/section[1]/article[1]/section[1]/div[1]/nav[1]/ul[1]/li[1]/a[1]/span[1]'))

            )
        except:
            print("Time exceed while finding most active")
        else:
            most_active.click()
            self.wait_for_page_load()

    def extract_stock_data(self):

        while True:
            self.wait.until(
                EC.presence_of_element_located((By.TAG_NAME, "table")))

            rows = driver.find_elements(By.CSS_SELECTOR, "table tbody tr")

            for row in rows:
                values = row.find_elements(By.TAG_NAME, "td")

                stock = {
                    "name": values[1].text,
                    "symbol": values[0].text,
                    "price": values[3].text,
                    "change": values[4].text,
                    "volume": values[6].text,
                    "market_cap": values[8].text,
                    "pe_ratio": values[9].text,
                }

                self.data.append(stock)

            try:
                next_button = self.wait.until(
                    EC.element_to_be_clickable(
                        (By.XPATH, '//*[@id="nimbus-app"]/section/section/section/article/section[1]/div/div[3]/div[3]/button[3]'))

                )
            except:
                print(
                    "Next button is not clickable we have traverse throught all the pages")
                break
            else:
                next_button.click()
                time.sleep(1)

    def clean_and_save_data(self, filename="stock_data"):
        stocks_df = pd.DataFrame(self.data)

        stocks_df = stocks_df.apply(
            lambda col: col.str.strip() if col.dtype == "object" else col)

        stocks_df["price"] = pd.to_numeric(stocks_df["price"])

        stocks_df["change"] = pd.to_numeric(
            stocks_df["change"].str.replace("+", "", regex=False))

        stocks_df["volume"] = pd.to_numeric(
            stocks_df["volume"].str.replace("M", "", regex=False))

        def convert_market_cap(val):
            if isinstance(val, str):
                if "B" in val:
                    return float(val.replace("B", ""))
                elif "T" in val:
                    return float(val.replace("T", "")) * 1000
            return np.nan

        stocks_df["market_cap"] = stocks_df["market_cap"].apply(
            convert_market_cap)

        stocks_df["pe_ratio"] = pd.to_numeric(
            stocks_df["pe_ratio"]
            .replace("-", np.nan)
            .replace("nan", np.nan)
            .astype(str)
            .str.replace(",", "", regex=False),
            errors='coerce'
        )

        stocks_df = stocks_df.rename(columns={
            "price": "price_usd",
            "volume": "volume_M",
            "market_cap": "market_cap_B"
        })

        stocks_df.to_excel(f"{filename}.xlsx", index=False)


if __name__ == "__main__":
    chrome_options = Options()
    chrome_options.add_argument("--disable-http2")
    chrome_options.add_argument("--incognito")
    chrome_options.add_argument(
        "--disable-blink-features=AutomationControlled")
    chrome_options.add_argument("--ignore-certificate-errors")
    chrome_options.add_argument("--enable-features=NetworkServiceInProcess")
    chrome_options.add_argument("--disable-features=NetworkService")
    chrome_options.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/93.0.4577.63 Safari/537.36"
    )

    driver = webdriver.Chrome(options=chrome_options)
    driver.maximize_window()

    url = "https://finance.yahoo.com/"
    scraper = StockScraper(driver)
    scraper.access_url(url)
    scraper.access_most_active_stocks()
    scraper.extract_stock_data()
    scraper.clean_and_save_data("stock_data")

    driver.quit()
