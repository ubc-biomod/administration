import os
import requests
from selenium import webdriver
from selenium.webdriver.chrome.options import Options

with open("urls.txt", "r", encoding="utf-8") as file:
    content = file.read()
    urls = content.split("\n")

# 1. Target PDF link
for pdf_url in urls: 
    pdf_id = pdf_url.split("F=")[1] 
    if len(pdf_id) <= 0:
        continue
    output_name = f"{pdf_id}_resume.pdf"
    print(output_name)

    # 2. Connect to your already open Chrome browser (launched with port 9222)
    chrome_options = Options()
    chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
    driver = webdriver.Chrome(options=chrome_options)

    # 3. Quickly ping the URL in Selenium just to ensure cookies for this domain are active
    driver.get(pdf_url)

    # 4. Copy the authenticated login cookies from Chrome into a Python requests session
    session = requests.Session()
    for cookie in driver.get_cookies():
        session.cookies.set(cookie['name'], cookie['value'])

    # 5. Download the file using your authenticated session
    print(f"Downloading PDF from: {pdf_url}")
    response = session.get(pdf_url)

    if response.status_code == 200:
        with open(output_name, 'wb') as f:
            f.write(response.content)
        print(f"Success! Saved as {output_name}")
    else:
        print(f"Failed to download. Status code: {response.status_code}")
