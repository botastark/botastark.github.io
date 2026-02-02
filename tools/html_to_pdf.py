import os
import sys
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from pathlib import Path
import time
import base64

# Directory containing HTML files
HTML_DIR = Path(os.path.dirname(__file__)).parent
OUTPUT_DIR = HTML_DIR / "to-be-slides"
OUTPUT_DIR.mkdir(exist_ok=True)

# Find all HTML files in the directory (non-recursive)
# Exclude index, projects, and alaris files
excluded_files = ["index.html", "projects.html", "alaris-project.html"]
html_files = sorted(
    [
        f
        for f in HTML_DIR.iterdir()
        if f.suffix == ".html" and f.name not in excluded_files
    ]
)

# Set up headless Chrome
chrome_options = Options()
chrome_options.add_argument("--headless=new")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")

driver = webdriver.Chrome(options=chrome_options)

# Store all PDF data
pdf_pages = []

for html_file in html_files:
    file_url = f"file://{html_file.resolve()}"
    driver.get(file_url)

    # Wait for rendering
    time.sleep(0.3)

    try:
        # Wait for slide container to load
        WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.CLASS_NAME, "slide-container"))
        )

        # Use Chrome's print to PDF functionality for highest quality
        pdf_options = {
            "landscape": False,
            "displayHeaderFooter": False,
            "printBackground": True,
            "preferCSSPageSize": False,
            "paperWidth": 10.67,  # 1280px at 120 DPI
            "paperHeight": 6.0,  # 720px at 120 DPI
            "marginTop": 0,
            "marginBottom": 0,
            "marginLeft": 0,
            "marginRight": 0,
            "scale": 1,
        }

        result = driver.execute_cdp_cmd("Page.printToPDF", pdf_options)
        pdf_data = base64.b64decode(result["data"])

        # Save individual PDF
        output_pdf = OUTPUT_DIR / f"{html_file.stem}.pdf"
        with open(output_pdf, "wb") as f:
            f.write(pdf_data)

        print(f"Generated PDF: {output_pdf.name}")
        pdf_pages.append(output_pdf)

    except Exception as e:
        print(f"Error creating PDF for {html_file}: {e}")

driver.quit()

print(f"\n✓ Created {len(pdf_pages)} individual PDF files in to-be-slides/")
print(
    f"To merge them, run: gs -dBATCH -dNOPAUSE -q -sDEVICE=pdfwrite -sOutputFile=to-be-slides/presentation_merged.pdf to-be-slides/*.pdf"
)
