#取得ETP配息
from selenium.webdriver.support.ui import WebDriverWait
from selenium_chrome import create_chrome_driver
import pandas as pd
import os
import re
import time
from datetime import datetime

url = "https://www.wantgoo.com/stock/etf/dividend"

# 匯出到 D 槽根目錄
output_dir = r"D:\SData"
run_date = datetime.now().strftime("%Y%m%d")
output_file = os.path.join(output_dir, f"wantgoo_etf_dividend_{run_date}.xlsx")

if not os.path.exists(output_dir):
    os.makedirs(output_dir)



driver = create_chrome_driver()

try:
    driver.get(url)

    wait = WebDriverWait(driver, 30)
    wait.until(lambda d: d.execute_script("return document.readyState") == "complete")

    time.sleep(5)

    driver.execute_script("window.scrollTo(0, 700);")
    time.sleep(2)

    # 只抓目前顯示的「近一年」表格。
    # 頁面另有隱藏的「近五年平均」表格，若使用 table tbody tr 會一起抓入。
    rows = driver.execute_script("""
        const targetTable = Array.from(document.querySelectorAll("table")).find(table => {
            const headerText = table.querySelector("thead")?.textContent || "";
            const style = window.getComputedStyle(table);
            const isVisible =
                style.display !== "none" &&
                style.visibility !== "hidden" &&
                table.getClientRects().length > 0;

            return isVisible &&
                   headerText.includes("1月") &&
                   headerText.includes("12月") &&
                   !headerText.includes("近五年平均");
        });

        if (!targetTable) {
            return [];
        }

        const allRows = Array.from(targetTable.querySelectorAll("tbody tr"));

        return allRows.map(tr => {
            return Array.from(tr.querySelectorAll("td")).map(td => {
                return td.textContent.trim();
            });
        });
    """)

    columns = [
        "代碼", "名稱", "配息", "殖利率",
        "1月", "2月", "3月", "4月", "5月", "6月",
        "7月", "8月", "9月", "10月", "11月", "12月"
    ]

    etf_code_pattern = re.compile(r"^\d{4,6}[A-Z]?$")

    clean_data = []

    for row in rows:
        if not row:
            continue

        code = str(row[0]).strip()

        if not etf_code_pattern.match(code):
            continue

        # 建立固定 16 欄，預設全部空白
        fixed_row = [""] * len(columns)

        # 依照 td 的位置放入資料
        # 空白 td 也會保留，不會讓後面的值往前移
        for i in range(min(len(row), len(columns))):
            fixed_row[i] = row[i].strip()

        clean_data.append(fixed_row)

    if len(clean_data) == 0:
        raise Exception("沒有抓到 ETF 資料列，請確認頁面是否正常載入。")

    df = pd.DataFrame(clean_data, columns=columns)

    # 清理換行
    for col in df.columns:
        df[col] = df[col].astype(str).str.replace("\n", "", regex=False).str.strip()

    # 空字串維持空白，不要補前值
    df = df.replace({"nan": "", "None": ""})

    # 移除完全相同的重複列，並依 ETF 代碼排序
    df = (
        df.drop_duplicates()
        .sort_values("代碼", key=lambda col: col.str.upper())
        .reset_index(drop=True)
    )

    # 數字欄位轉數值，空白保持空白
    num_cols = [
        "配息", "殖利率",
        "1月", "2月", "3月", "4月", "5月", "6月",
        "7月", "8月", "9月", "10月", "11月", "12月"
    ]

    for col in num_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    # 匯出 Excel，空值會是空白格
    df.to_excel(output_file, index=False, engine="openpyxl", na_rep="")

    print("取得資料列數：", len(df))
    print("完成匯出：", output_file)
    print(df.head(20))

finally:
    driver.quit()
