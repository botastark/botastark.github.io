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
import subprocess

# Directory containing HTML files
BASE_DIR = Path(os.path.dirname(__file__)).parent
TUM_DIR = BASE_DIR / "TUM"
OUTPUT_DIR = TUM_DIR / "output"
OUTPUT_DIR.mkdir(exist_ok=True)

# Find all page HTML files in TUM directory and sort them numerically
html_files = sorted(
    [f for f in TUM_DIR.iterdir() if f.suffix == ".html" and f.name.startswith("page")],
    key=lambda x: int(x.stem.replace("page", "")),
)

print(f"Found {len(html_files)} HTML pages in TUM folder")

# Set up headless Chrome
chrome_options = Options()
chrome_options.add_argument("--headless=new")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--window-size=1280,720")
chrome_options.add_argument("--force-device-scale-factor=1")

driver = webdriver.Chrome(options=chrome_options)
driver.set_window_size(1280, 720)

# Store all PDF file paths
pdf_pages = []

for html_file in html_files:
    file_url = f"file://{html_file.resolve()}"
    driver.get(file_url)

    # Wait for rendering and KaTeX if present
    time.sleep(1.5)  # Increased wait time for KaTeX rendering

    try:
        # Wait for slide container to load
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CLASS_NAME, "slide-container"))
        )

        # Use Chrome's print to PDF functionality
        # Set exact dimensions to match 1280x720 canvas (in inches)
        # 1280px ÷ 96 DPI = 13.333 inches width
        # 720px ÷ 96 DPI = 7.5 inches height
        pdf_options = {
            "landscape": True,
            "displayHeaderFooter": False,
            "printBackground": True,
            "preferCSSPageSize": True,
            "marginTop": 0,
            "marginBottom": 0,
            "marginLeft": 0,
            "marginRight": 0,
            "scale": 1.0,
        }

        result = driver.execute_cdp_cmd("Page.printToPDF", pdf_options)
        pdf_data = base64.b64decode(result["data"])

        # Save individual PDF
        output_pdf = OUTPUT_DIR / f"{html_file.stem}.pdf"
        with open(output_pdf, "wb") as f:
            f.write(pdf_data)

        print(f"✓ Generated PDF: {html_file.stem}.pdf")
        pdf_pages.append(output_pdf)

    except Exception as e:
        print(f"✗ Error creating PDF for {html_file.name}: {e}")

driver.quit()

print(f"\n✓ Created {len(pdf_pages)} individual PDF files in TUM/output/")

# Merge PDFs using ghostscript
if pdf_pages:
    merged_pdf = OUTPUT_DIR / "TUM_presentation.pdf"
    pdf_files = [str(p) for p in pdf_pages]

    gs_command = [
        "gs",
        "-dBATCH",
        "-dNOPAUSE",
        "-q",
        "-sDEVICE=pdfwrite",
        f"-sOutputFile={merged_pdf}",
    ] + pdf_files

    try:
        print(f"\nMerging PDFs into {merged_pdf.name}...")
        subprocess.run(gs_command, check=True)
        print(f"✓ Successfully created merged PDF: {merged_pdf}")
        print(f"  Location: {merged_pdf}")
    except subprocess.CalledProcessError as e:
        print(f"✗ Error merging PDFs: {e}")
        print(f"  You can manually merge using:")
        print(
            f"  gs -dBATCH -dNOPAUSE -q -sDEVICE=pdfwrite -sOutputFile=TUM/output/TUM_presentation.pdf TUM/output/page*.pdf"
        )
    except FileNotFoundError:
        print("✗ Ghostscript (gs) not found. Please install it to merge PDFs:")
        print("  sudo apt-get install ghostscript")
        print(
            f"  Then run: gs -dBATCH -dNOPAUSE -q -sDEVICE=pdfwrite -sOutputFile=TUM/output/TUM_presentation.pdf TUM/output/page*.pdf"
        )
else:
    print("No PDFs were created.")
