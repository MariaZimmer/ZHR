from selenium import webdriver
from selenium.webdriver.chrome.service import Service

service = Service("e:/Users/hp/Downloads/rozliczenie2025/chromedriver.exe")
driver = webdriver.Chrome(service=service)
driver.get("https://www.google.com")
input("Naciśnij Enter, aby zamknąć...")
