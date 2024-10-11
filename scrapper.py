from selenium import webdriver
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup
import pandas as pd

# Website for Sentiment analysis

options = webdriver.ChromeOptions()
options.add_experimental_option('excludeSwitches', ['enable-logging'])
driver = webdriver.Chrome("./drivers/chromedriver", options=options)

driver.get("https://www.macrotrends.net/stocks/charts/{stock_ticker}/apple/{fundamental}")

fundamentals = ['cash-flow-statement','financial-ratios'

content = driver.page_source
soup = BeautifulSoup(content)

hScrollBar = driver.find_element(By.ID,"jqxScrollThumbhorizontalScrollBarjqxgrid")
hearders = 'jqx-widget-header jqx-grid-header'
columnTag = 'jqx-grid-column-header jqx-widget-header'
rowTags = 'jqx-grid-cell jqx-item jqx-grid-cell-pinned'
# driver.execute_script("arguments.scrollBy(0,arguments[0].scrollHeight)",hScrollBar)


dataColumn = []
for a in soup.find_all(class_= hearders)[0].find_all(class_=columnTag):
    if a.text != "":
        dataColumn.append(a.text)
        print('col found' + a.text)


rows = []
for rElement in driver.find_elements(By.XPATH, "//div/div[contains(@id,'row')]"):
    dataRow= []
    for cell in rElement.find_elements(By.XPATH, ".//div[contains(@class,'jqx-grid-cell')]/div"):
        if cell.text != "":
            dataRow.append(cell.text)
            print('row found ' + cell.text)
    print('**********************************************************************')
    rows.append(dataRow)
    

df = pd.DataFrame(rows,  columns = dataColumn)
print(df)
file_name = 'MarksData.xlsx'
  
# saving the excel
df.to_excel(file_name)
print('DataFrame is written to Excel File successfully.')
