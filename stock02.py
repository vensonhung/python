#玩股網產出ETF淨值
import sys
from pathlib import Path

# 直接執行時使用腳本所在目錄；貼到 Notebook 時使用專案目錄。
PROJECT_DIR = (
    Path(__file__).resolve().parent
    if "__file__" in globals()
    else Path(r"D:\SData")
)
if not (PROJECT_DIR / "selenium_chrome.py").is_file():
    raise FileNotFoundError(
        "找不到 {}，請確認 selenium_chrome.py 與 stock02.py 放在同一目錄。".format(
            PROJECT_DIR / "selenium_chrome.py"
        )
    )
if str(PROJECT_DIR) not in sys.path:
    sys.path.insert(0, str(PROJECT_DIR))

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium_chrome import create_chrome_driver
import pandas as pd
import time
from datetime import datetime

url = "https://www.wantgoo.com/stock/etf/net-value"

driver = create_chrome_driver()

try:
    driver.execute_cdp_cmd(
        "Page.addScriptToEvaluateOnNewDocument",
        {
            "source": """
                Object.defineProperty(navigator, 'webdriver', {
                    get: () => undefined
                })
            """
        }
    )

    driver.get(url)

    wait = WebDriverWait(driver, 60)

    wait.until(
        lambda d: len(d.find_elements(By.CSS_SELECTOR, "table.table-sequence")) > 0
    )

    wait.until(
        lambda d: len(d.find_elements(By.CSS_SELECTOR, "table.table-sequence tbody tr")) > 0
    )

    time.sleep(3)

    rows = driver.find_elements(By.CSS_SELECTOR, "table.table-sequence tbody tr")

    data = []

    for row in rows:
        cells = row.find_elements(By.TAG_NAME, "td")
        values = [cell.text.strip().replace("\n", " ") for cell in cells]

        if len(values) < 10:
            continue

        data.append({
            "ETF代號": values[0],
            "ETF名稱": values[1],
            "淨值": values[2],
            "淨值漲跌%": values[3],
            "市價": values[4],
            "市價漲跌%": values[5],
            "折溢價": values[6],
            "折溢價%": values[7],
            "成交量": values[8],
            "追蹤標的": values[9],
        })

    df = pd.DataFrame(data)

finally:
    driver.quit()


# =========================
# 匯出 Excel
# =========================

df["ETF代號"] = df["ETF代號"].astype(str).str.strip()

df["ETF代號"] = df["ETF代號"].apply(
    lambda x: x.zfill(4) if x.isdigit() else x
)

run_date = datetime.now().strftime("%Y%m%d")
output_path = fr"D:\SData\wantgoo_etf_net_value_{run_date}.xlsx"

with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
    df.to_excel(writer, index=False, sheet_name="ETF淨值來源")

    ws = writer.book["ETF淨值來源"]

    for cell in ws["A"]:
        cell.number_format = "@"

    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 28
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 12
    ws.column_dimensions["E"].width = 12
    ws.column_dimensions["F"].width = 12
    ws.column_dimensions["G"].width = 12
    ws.column_dimensions["H"].width = 12
    ws.column_dimensions["I"].width = 12
    ws.column_dimensions["J"].width = 45

print("已匯出：", output_path)

df
