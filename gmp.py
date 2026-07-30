import os
from datetime import datetime, timedelta
from pathlib import Path
import time
import openpyxl
import pandas as pd
from playwright.sync_api import sync_playwright
import requests


def ntfy_message(message_text):
    """
    Sirf text message phone par bhejne ke liye.
    """
    url = "https://ntfy.sh/hello_anjal_bhi"
    requests.post(url, data=message_text.encode("utf-8"))


current_dir = Path(__file__).resolve().parent
path = f"{current_dir}/files/"

# Ensure 'files' folder exists automatically
os.makedirs(path, exist_ok=True)


def file_delete(filename):
    file_path = f"{current_dir}/files/{filename}"
    if os.path.exists(file_path):
        os.remove(file_path)
        return True
    return False


# 1. DATA READ KARNE KE LIYE
def xlsx_value(position_row_and_column, file_name):
    position_row_and_column = position_row_and_column.upper()
    wb = openpyxl.load_workbook(f"{current_dir}/files/{file_name}")
    kkkk = wb.active[position_row_and_column].value
    return kkkk


# 2. DATA UPDATE KARNE KE LIYE
def xlsx_update(position_row_and_column, new_value, file_name):
    wb = openpyxl.load_workbook(f"{current_dir}/files/{file_name}")
    wb.active[position_row_and_column.upper()] = new_value
    wb.save(f"{current_dir}/files/{file_name}")


def date():
    return datetime.now().strftime("%Y-%m-%d")


def wait_time(seconds):
    time.sleep(seconds)


def time_add(start_time_str, hours_to_add):
    dt = datetime.strptime(start_time_str, "%Y-%m-%d %H:%M:%S")
    new_dt = dt + timedelta(hours=int(hours_to_add))
    return new_dt.strftime("%Y-%m-%d %H:%M:%S")


def time_time():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


url = "https://www.investorgain.com/report/ipo-gmp-live/331/"

with sync_playwright() as p:
    # GitHub Actions ke liye headless=True hona zaroori hai
    browser = p.chromium.launch(headless=True)
    page = browser.new_page()

    page.goto(url)
    page.wait_for_timeout(5000)

    for i in range(5):
        page.keyboard.press("PageDown")
        time.sleep(1)

    page.wait_for_timeout(3000)
    poora_html = page.content()
    browser.close()

# HTML Save
html_file_name = f"{path}gmp.html"
with open(html_file_name, "w", encoding="utf-8") as file:
    file.write(poora_html)

# Excel Conversion
html_file = f"{path}gmp.html"
try:
    tables = pd.read_html(html_file)
    main_table = max(tables, key=len)
    excel_file = f"{path}gmp.xlsx"
    main_table.to_excel(excel_file, index=False, sheet_name="Live IPO GMP")
except Exception as e:
    print(f"Excel convert karne mein error aayi: {e}")

# Filtering & Processing
input_file = f"{path}gmp.xlsx"
try:
    xls = pd.ExcelFile(input_file)
    raw_df = pd.read_excel(input_file, sheet_name=xls.sheet_names[0])

    headers = raw_df.columns
    col_name = headers[0]  # Name
    col_rating = headers[2]  # Rating

    mainboard_file = f"{path}Mainboard_IPOs.xlsx"
    sme_file = f"{path}SME_IPOs.xlsx"

    df_empty = pd.DataFrame(columns=headers)
    df_empty.to_excel(mainboard_file, index=False, sheet_name="Top Mainboard")
    df_empty.to_excel(sme_file, index=False, sheet_name="Top SME")

    mainboard_rows = []
    sme_rows = []

    for index, row in raw_df.iterrows():
        val_a = row[col_name]
        if pd.isna(val_a) or str(val_a).strip() == "":
            continue

        val_str = str(val_a).replace(" ", "")
        val_upper = val_str.upper()

        is_mainboard = (
            "IPOU" in val_upper or "IPOO" in val_upper or "IPOCT" in val_upper
        )
        is_sme = "SMEU" in val_upper or "SMEO" in val_upper or "SMECT" in val_upper

        if is_mainboard or is_sme:
            rating_val = str(row[col_rating]).strip()
            if rating_val in ["🔥🔥🔥🔥", "🔥🔥🔥🔥🔥"]:
                if is_mainboard:
                    mainboard_rows.append(row)
                elif is_sme:
                    sme_rows.append(row)

    if mainboard_rows:
        df_mainboard = pd.DataFrame(mainboard_rows, columns=headers)
        df_mainboard.to_excel(
            mainboard_file, index=False, sheet_name="Top Mainboard"
        )

    if sme_rows:
        df_sme = pd.DataFrame(sme_rows, columns=headers)
        df_sme.to_excel(sme_file, index=False, sheet_name="Top SME")

    xlsx_update("c13", len(mainboard_rows), "main.xlsx")
    xlsx_update("c14", len(sme_rows), "main.xlsx")
except Exception as e:
    print(f"Error: {e}")

xlsx_update("d12", time_time(), "main.xlsx")
xlsx_update("e12", time_add(time_time(), 15), "main.xlsx")

# Send NTFY Alerts
mainboard_row = 1
while mainboard_row < int(xlsx_value("c13", "main.xlsx")) + 1:
    mainboard_row += 1
    ntfy_message(f"""MAIN_IPO:- {xlsx_value(f'a{mainboard_row}', 'Mainboard_IPOs.xlsx')} 
PRICE:-  {xlsx_value(f"b{mainboard_row}", "Mainboard_IPOs.xlsx")}
RATING:-  {xlsx_value(f"c{mainboard_row}", "Mainboard_IPOs.xlsx")}
CLOSING DATE:-  {xlsx_value(f"i{mainboard_row}", "Mainboard_IPOs.xlsx")}
LISTING DATE:-   {xlsx_value(f"k{mainboard_row}", "Mainboard_IPOs.xlsx")}""")
    wait_time(1)

sme_row = 1
while sme_row < int(xlsx_value("c14", "main.xlsx")) + 1:
    sme_row += 1
    ntfy_message(f"""SME_IPO:- {xlsx_value(f'a{sme_row}', 'SME_IPOs.xlsx')} 
PRICE:-  {xlsx_value(f"b{sme_row}", "SME_IPOs.xlsx")}
RATING:-  {xlsx_value(f"c{sme_row}", "SME_IPOs.xlsx")}
CLOSING DATE:-  {xlsx_value(f"i{sme_row}", "SME_IPOs.xlsx")}
LISTING DATE:-   {xlsx_value(f"k{sme_row}", "SME_IPOs.xlsx")}""")
    wait_time(1)

xlsx_update("c15", date(), "main.xlsx")
