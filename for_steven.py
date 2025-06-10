from selenium import webdriver
from selenium.common.exceptions import SessionNotCreatedException
from PIL import Image
import time

# from main import chrome_options

options = webdriver.ChromeOptions()
options.add_experimental_option("detach", True)
# options.add_argument(r"--user-data-dir=C:\Users\PC\AppData\Local\Google\Chrome\User Data\Default")
options.add_argument(r"--user-data-dir=C:\Users\PC\AppData\Local\Google\Chrome\User Data")
options.add_argument("--profile-directory=Profile 5")
# try:
driver = webdriver.Chrome(options=options)
# except SessionNotCreatedException as e:
#     print("Session failed to start")
#     exit()
driver.get("https://www.youtube.com/")


time.sleep(2)


screenshot_path = "full_screenshot.png"
driver.save_screenshot(screenshot_path)

crop_box = (50, 100, 450, 400)

img = Image.open(screenshot_path)
cropped_img = img.crop(crop_box)

cropped_img_path = "cropped_screenshot.png"
cropped_img.save(cropped_img_path)

print(f"Cropped screenshot saved to {cropped_img_path}")

# driver.quit()
