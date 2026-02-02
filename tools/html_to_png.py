import os
import sys
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from pathlib import Path
from PIL import Image
import io
import time

# Directory containing HTML files
HTML_DIR = Path(os.path.dirname(__file__)).parent
OUTPUT_DIR = HTML_DIR / "to-be-slides"
OUTPUT_DIR.mkdir(exist_ok=True)

# Find all HTML files in the directory (non-recursive)
# Exclude index, projects, and alaris files
excluded_files = ["index.html", "projects.html", "alaris-project.html"]
html_files = [
    f
    for f in HTML_DIR.iterdir()
    if f.suffix == ".html" and f.name not in excluded_files
]

# Set up headless Chrome with exact dimensions
chrome_options = Options()
chrome_options.add_argument("--headless=new")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--hide-scrollbars")

# Path to chromedriver (assume it's in PATH)
driver = webdriver.Chrome(options=chrome_options)

for html_file in html_files:
    file_url = f"file://{html_file.resolve()}"
    driver.get(file_url)

    # Wait for rendering
    time.sleep(0.3)

    try:
        # Wait for slide container to load
        slide_element = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.CLASS_NAME, "slide-container"))
        )

        # Set viewport to exact slide dimensions using CDP
        driver.execute_cdp_cmd(
            "Emulation.setDeviceMetricsOverride",
            {"width": 1280, "height": 720, "deviceScaleFactor": 1, "mobile": False},
        )

        # Wait for layout recalculation
        time.sleep(0.3)

        # Take screenshot
        png_bytes = driver.get_screenshot_as_png()
        img = Image.open(io.BytesIO(png_bytes))

        # Save as PNG
        out_path = OUTPUT_DIR / (html_file.stem + ".png")
        img.save(str(out_path), "PNG")
        print(f"Saved {out_path} [{img.size[0]}x{img.size[1]}]")
    except TimeoutException:
        print(f"Timeout loading {html_file}")
    except Exception as e:
        print(f"Error taking screenshot of {html_file}: {e}")

driver.quit()
