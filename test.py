from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time

# Chrome browser start
driver = webdriver.Chrome()

try:
    # Google open
    driver.get("https://www.google.com")

    # Browser maximize
    driver.maximize_window()

    # Google search box
    search_box = driver.find_element(By.NAME, "q")

    # Search text
    search_box.send_keys("Python kya hai")

    # Enter press
    search_box.send_keys(Keys.RETURN)

    # Result dekhne ke liye wait
    time.sleep(5)

finally:
    # Browser close
    driver.quit()