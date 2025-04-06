from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
import os
import time
import shutil
import pandas as pd
from selenium.webdriver.support.wait import WebDriverWait

opt = Options()
opt.add_experimental_option("debuggerAddress", "localhost:8989")
opt.add_argument("--start-maximized")
service = Service("chromedriver.exe")
driver = webdriver.Chrome(service=service, options=opt)
# === Load Excel ===
df = pd.read_excel("Book1.xlsx", engine="openpyxl")

downloads_dir = os.path.join(os.path.expanduser("~"), "Downloads")

if not os.path.exists("M:\\College\\Python\\FLS Image"):
    os.makedirs("M:\\College\\Python\\FLS Image")


# === Process Each Row ===
for index, row in df.iterrows():
    fls_id = str(row["FLS ID"])
    file_url = row["URL"]

    print(f"Downloading {fls_id} from {file_url}")
    driver.get(file_url)
    time.sleep(5)

    try:
        # Find the download button (top-right corner)
        download_btn = WebDriverWait(driver, 15).until(EC.visibility_of_element_located((By.XPATH, "//div[@aria-label='Download']")))

        # Optional: highlight button
        driver.execute_script("arguments[0].style.border='2px solid green'", download_btn)

        # Use ActionChains to move and click
        actions = ActionChains(driver)
        actions.move_to_element(download_btn).click().perform()

        print("✅ Download button clicked using ActionChains")
    except Exception as e:
        print(f"❌ Failed to click download button: {e}")

    time.sleep(4)

    # === Find latest .jpg file in Downloads ===
    jpg_files = [f for f in os.listdir(downloads_dir) if f.lower().endswith(".jpg")]
    jpg_files = sorted(jpg_files, key=lambda x: os.path.getmtime(os.path.join(downloads_dir, x)), reverse=True)

    if jpg_files:
        latest_file = jpg_files[0]
        source_path = os.path.join(downloads_dir, latest_file)
        target_path = os.path.join("M:\\College\\Python\\FLS Image", f"{fls_id}.jpg")

        # Rename and move
        shutil.move(source_path, target_path)
        print(f"📁 Moved and renamed: {latest_file} ➜ {fls_id}.jpg")
    else:
        print(f"❌ No JPG files found in Downloads for {fls_id}")

