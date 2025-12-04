📌 Requirements (Skip if already have)

Python 3.8+ (Recommended)
https://www.python.org/downloads/

Required Python packages:
pandas
openpyxl
selenium (automated downloads)

▶ How to Install - Install Python Packages
Run this in terminal:
pip install pandas openpyxl selenium

📥 Preparing Raw Files
Before running the script:

Visit the link below to clone the project into your working directory:
https://github.com/wilsoncheah55-dotcom/MrDIY_Automation_Task#

Ensure you have all 4 files below in the same filepath or folder:
Python_data_processing_module.py
excel_sample_data_qae.xlsx
msedgedriver.exe
README.md

If your Microsoft Edge version does not match the included WebDriver (msedgedriver.exe), download the WebDriver version that corresponds to your current Edge browser:
https://developer.microsoft.com/en-us/microsoft-edge/tools/webdriver/

⚠️ How to Run the Script:
1. Open Cmd, navigate to your working directory (eg: cd D:\Automation Anywhere Files\Automation Anywhere\My Docs\Report Downloader):
2. Type "python Python_data_processing_module.py" and press Enter

Once the scraping completes, you will see three files downloaded to the following location:
C:\Users\Your_Username\Downloads\exchange-rates.csv (Buying)
C:\Users\Your_UsernameDownloads\exchange-rates (1).csv (Middle Rate)
C:\Users\Your_Username\Downloads\exchange-rates (2).csv (Selling)

Three CSV dataset files will be generated on your working directory (same Path or Folder as your python script):
product.csv
sales.csv
store.csv
compiled_rates.csv

Then, you will be prompted to:
3. Choose a Region (or type "ALL")

4. Choose a Product Category (or type "ALL")

The system is case-insensitive, so:
north
NORTH
North
NoRtH

all work.

📊 Output - All sales_amount, sales_cost & profit already converted to MYR
After running successfully, the script will generate a final excel file (aggregated data) on your working directory:
sales_report.xlsx

This Excel file includes 2 sheets:
1️⃣ By Region
2️⃣ By Product Category

Each sheet includes:
Formatted headers, auto-sized columns, borders around data, and values rounded to at most 2 decimal places
