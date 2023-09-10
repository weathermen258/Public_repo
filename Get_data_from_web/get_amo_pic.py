from selenium import webdriver
import time
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service as ChromeService
import pandas as pd
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select

options = webdriver.ChromeOptions()
options.add_argument('--headless')
driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()),options=options)
driver.get("http://amo.gov.vn/radar/VIN")
driver.maximize_window()
driver.get_screenshot_as_file("./amo.png")
driver.quit()
from PIL import Image
img = Image.open('./amo.png')
box = (210, 180, 500, 500)
img2 = img.crop(box)
img2.save('amo_cropped.png')
#img2.show()
