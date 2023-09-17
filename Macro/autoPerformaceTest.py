from browsermobproxy import Server
from selenium import webdriver
import json

# Path to BrowserMob Proxy batch file
server = Server("C:\\Users\\qaxsd\\OneDrive - 서울과학기술대학교\\Project\\Automation\\browsermob-proxy-2.1.4-bin\\browsermob-proxy-2.1.4\\bin\\browsermob-proxy.bat")
server.start()
proxy = server.create_proxy()

chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument("--proxy-server={0}".format(proxy.proxy))
chrome_options.add_argument('--ignore-certificate-errors')
driver = webdriver.Chrome(options = chrome_options)

proxy.new_har("seoultechUniv")
driver.get("https://www.seoultech.ac.kr/index.jsp")

# Get the HAR data
har_data = proxy.har

# This data can be analyzed in depth or written to a file for future analysis.
with open('har_data.json', 'w') as f:
    json.dump(har_data, f)

driver.quit()
server.stop()