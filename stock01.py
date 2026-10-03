#取得台股單日交易資料
import requests
import pandas as pd
from datetime import datetime
import os


def clean_number(value):
    """
    將 1,234、--、空白 轉成數字
    """
    if pd.isna(value):
        return None

    value = str(value).replace(",", "").replace("--", "").strip()

    if value == "":
        return None

    try:
        return float(value)
    except:
        return None


def filter_security_code(df, code_col):
    """
    篩選證券代號，支援：
    4碼股票：2330
    5碼ETF：00878
    6碼ETF：006208
    含英文ETF：00980A
    7碼特殊代號
    """

    df = df.copy()

    df[code_col] = (
        df[code_col]
        .astype(str)
        .str.strip()
        .str.upper()
    )

    # 支援 4~7 碼，允許數字與英文字母
    # 例如：2330、0050、00878、006208、00980A、7碼代號
    mask = df[code_col].str.match(r"^[0-9A-Z]{4,7}$")

    return df[mask].copy()


def get_twse_all_stocks(target_date):
    """
    取得上市股票 TWSE 指定日期全部收盤行情
    target_date: 西元日期，例如 '2025-05-15'
    """

    dt = datetime.strptime(target_date, "%Y-%m-%d")
    query_date = dt.strftime("%Y%m%d")

    url = "https://www.twse.com.tw/exchangeReport/MI_INDEX"

    params = {
        "response": "json",
        "date": query_date,
        "type": "ALLBUT0999"
    }

    headers = {
        "User-Agent": "Mozilla/5.0"
    }

    res = requests.get(url, params=params, headers=headers, timeout=20)
    data = res.json()

    if data.get("stat") != "OK":
        print("TWSE 查詢失敗：", data.get("stat"))
        return pd.DataFrame()

    target_table = None

    for table in data.get("tables", []):
        title = table.get("title", "")
        if "每日收盤行情" in title:
            target_table = table
            break

    if target_table is None:
        print("TWSE 找不到每日收盤行情資料")
        return pd.DataFrame()

    df = pd.DataFrame(
        target_table["data"],
        columns=target_table["fields"]
    )

    # 修改重點：支援 4~7 碼與含英文字母代號
    df = filter_security_code(df, "證券代號")

    result = pd.DataFrame()
    result["交易日期"] = target_date
    result["市場別"] = "上市"
    result["股票代號"] = df["證券代號"]
    result["股票名稱"] = df["證券名稱"]
    result["成交股數"] = df["成交股數"].apply(clean_number)
    result["成交金額"] = df["成交金額"].apply(clean_number)
    result["開盤價"] = df["開盤價"].apply(clean_number)
    result["最高價"] = df["最高價"].apply(clean_number)
    result["最低價"] = df["最低價"].apply(clean_number)
    result["收盤價"] = df["收盤價"].apply(clean_number)
    result["漲跌價差"] = df["漲跌價差"].apply(clean_number)
    result["成交筆數"] = df["成交筆數"].apply(clean_number)

    return result


def get_tpex_all_stocks(target_date):
    """
    取得上櫃股票 TPEx 指定日期全部收盤行情
    target_date: 西元日期，例如 '2025-05-15'
    """

    dt = datetime.strptime(target_date, "%Y-%m-%d")

    # TPEx 使用民國日期格式：114/05/15
    roc_date = f"{dt.year - 1911}/{dt.month:02d}/{dt.day:02d}"

    url = "https://www.tpex.org.tw/www/zh-tw/afterTrading/dailyQuotes"

    params = {
        "date": roc_date,
        "response": "json"
    }

    headers = {
        "User-Agent": "Mozilla/5.0"
    }

    res = requests.get(url, params=params, headers=headers, timeout=20)
    data = res.json()

    tables = data.get("tables", [])

    if not tables:
        print("TPEx 查無資料")
        return pd.DataFrame()

    table = tables[0]

    df = pd.DataFrame(
        table["data"],
        columns=table["fields"]
    )

    code_col = "代號"
    name_col = "名稱"

    if code_col not in df.columns:
        print("TPEx 找不到代號欄位，實際欄位如下：")
        print(df.columns)
        return pd.DataFrame()

    # 修改重點：TPEx 也要支援 4~7 碼與含英文字母代號
    df = filter_security_code(df, code_col)

    result = pd.DataFrame()
    result["交易日期"] = target_date
    result["市場別"] = "上櫃"
    result["股票代號"] = df[code_col]
    result["股票名稱"] = df[name_col]

    result["成交股數"] = df["成交股數"].apply(clean_number) if "成交股數" in df.columns else None
    result["成交金額"] = df["成交金額(元)"].apply(clean_number) if "成交金額(元)" in df.columns else None
    result["開盤價"] = df["開盤"].apply(clean_number) if "開盤" in df.columns else None
    result["最高價"] = df["最高"].apply(clean_number) if "最高" in df.columns else None
    result["最低價"] = df["最低"].apply(clean_number) if "最低" in df.columns else None
    result["收盤價"] = df["收盤"].apply(clean_number) if "收盤" in df.columns else None
    result["漲跌價差"] = df["漲跌"].apply(clean_number) if "漲跌" in df.columns else None
    result["成交筆數"] = df["成交筆數"].apply(clean_number) if "成交筆數" in df.columns else None

    return result


def export_twse_tpex_all_stocks(target_date):
    """
    合併上市 + 上櫃，並匯出 Excel 到 D:\SData
    """

    df_twse = get_twse_all_stocks(target_date)
    df_tpex = get_tpex_all_stocks(target_date)

    df_all = pd.concat([df_twse, df_tpex], ignore_index=True)

    if df_all.empty:
        print("上市與上櫃都查無資料，請確認日期是否為交易日。")
        return None

    # 股票代號保持文字，避免 0050、00878 變成 50、878
    df_all["股票代號"] = df_all["股票代號"].astype(str).str.strip().str.upper()

    if not df_twse.empty:
        df_twse["股票代號"] = df_twse["股票代號"].astype(str).str.strip().str.upper()

    if not df_tpex.empty:
        df_tpex["股票代號"] = df_tpex["股票代號"].astype(str).str.strip().str.upper()

    # 排序：上市在前、上櫃在後，再依股票代號排序
    market_order = {
        "上市": 1,
        "上櫃": 2
    }

    df_all["市場排序"] = df_all["市場別"].map(market_order)
    df_all = df_all.sort_values(["市場排序", "股票代號"]).drop(columns=["市場排序"])

    # 修改重點：固定輸出到 D:\SData
    output_dir = r"D:\SData"
    os.makedirs(output_dir, exist_ok=True)

    file_name = os.path.join(
        output_dir,
        f"台股上市上櫃每日股價_{target_date}.xlsx"
    )

    with pd.ExcelWriter(file_name, engine="openpyxl") as writer:
        df_all.to_excel(writer, sheet_name="上市上櫃合併", index=False)
        df_twse.to_excel(writer, sheet_name="上市TWSE", index=False)
        df_tpex.to_excel(writer, sheet_name="上櫃TPEx", index=False)

        # 設定股票代號欄位為文字格式
        for sheet_name in ["上市上櫃合併", "上市TWSE", "上櫃TPEx"]:
            ws = writer.book[sheet_name]

            # 欄位順序：
            # A 交易日期
            # B 市場別
            # C 股票代號
            # D 股票名稱
            for cell in ws["C"]:
                cell.number_format = "@"

            ws.column_dimensions["A"].width = 14
            ws.column_dimensions["B"].width = 10
            ws.column_dimensions["C"].width = 14
            ws.column_dimensions["D"].width = 24
            ws.column_dimensions["E"].width = 14
            ws.column_dimensions["F"].width = 16
            ws.column_dimensions["G"].width = 12
            ws.column_dimensions["H"].width = 12
            ws.column_dimensions["I"].width = 12
            ws.column_dimensions["J"].width = 12
            ws.column_dimensions["K"].width = 12
            ws.column_dimensions["L"].width = 12

    print("已匯出：", file_name)
    print("上市筆數：", len(df_twse))
    print("上櫃筆數：", len(df_tpex))
    print("合併筆數：", len(df_all))

    return df_all


# =========================
# 使用範例
# =========================

target_date = datetime.now().strftime("%Y-%m-%d")

df_all = export_twse_tpex_all_stocks(target_date)

if df_all is not None:
    print(df_all.head())