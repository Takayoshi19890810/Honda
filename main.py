import os
import re
import json
import time
import pandas as pd
from bs4 import BeautifulSoup
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# ✅ Google Sheets 設定
SPREADSHEET_ID = "1AwwMGKMHfduwPkrtsik40lkO1z1T8IU_yd41ku-yPi8"
SHEET_NAME = "MSN"

# ✅ JST時刻取得
now = datetime.utcnow() + timedelta(hours=9)

# ✅ スプレッドシート接続
def authorize_gspread():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    if os.path.exists("credentials.json"):
        creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
    else:
        creds_dict = json.loads(os.environ.get("GCP_SERVICE_ACCOUNT_KEY"))
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    return gspread.authorize(creds)

# ✅ 日付パース
def parse_pub_label(pub_label):
    pub_time_obj = None
    if "分前" in pub_label:
        m = re.search(r"(\d+)", pub_label)
        if m: pub_time_obj = now - timedelta(minutes=int(m.group(1)))
    elif "時間前" in pub_label:
        h = re.search(r"(\d+)", pub_label)
        if h: pub_time_obj = now - timedelta(hours=int(h.group(1)))
    elif "日前" in pub_label:
        d = re.search(r"(\d+)", pub_label)
        if d: pub_time_obj = now - timedelta(days=int(d.group(1)))
    elif re.match(r'\d+月\d+日', pub_label):
        ymd = f"{now.year}年{pub_label}"
        pub_time_obj = datetime.strptime(ymd, "%Y年%m月%d日")
    elif re.match(r'\d{4}/\d{1,2}/\d{1,2}', pub_label):
        pub_time_obj = datetime.strptime(pub_label, "%Y/%m/%d")
    elif re.match(r'\d{1,2}:\d{2}', pub_label):
        t = datetime.strptime(pub_label, "%H:%M").time()
        pub_time_obj = datetime.combine(now.date(), t)
    return pub_time_obj.strftime("%Y/%m/%d %H:%M") if pub_time_obj else pub_label

# ✅ MSNニュース取得
def get_msn_news(keyword):
    options = Options()
    options.add_argument("--headless")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

    url = f'https://www.bing.com/news/search?q={keyword}&qft=sortbydate%3d"1"&form=YFNR'
    driver.get(url)
    time.sleep(5)

    soup = BeautifulSoup(driver.page_source, "html.parser")
    driver.quit()

    titles, urls, pub_dates, sources = [], [], [], []

    for card in soup.select('div.news-card'):
        title = card.get("data-title", "").strip()
        url = card.get("data-url", "").strip()
        source = card.get("data-author", "").strip()

        pub_label = ""
        pub_tag = card.find("span", attrs={"aria-label": True})
        if pub_tag:
            pub_label = pub_tag["aria-label"].strip()

        pub_date = parse_pub_label(pub_label)

        if title and url:
            titles.append(title)
            urls.append(url)
            pub_dates.append(pub_date)
            sources.append(source if source else "MSN")

    return pd.DataFrame({
        "タイトル": titles,
        "URL": urls,
        "投稿日": pub_dates,
        "引用元": sources
    })

# ✅ スプレッドシート書き込み（重複チェックあり）
def write_to_spreadsheet(df, sheet_id, sheet_name):
    gc = authorize_gspread()
    sh = gc.open_by_key(sheet_id)

    try:
        ws = sh.worksheet(sheet_name)
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(title=sheet_name, rows="100", cols="4")
        ws.append_row(["タイトル", "URL", "投稿日", "引用元"])

    existing = ws.get_all_values()
    existing_urls = {row[1] for row in existing[1:] if len(row) > 1}

    new_rows = df[~df["URL"].isin(existing_urls)]

    if not new_rows.empty:
        ws.append_rows(new_rows.values.tolist(), value_input_option="USER_ENTERED")
        print(f"✅ {len(new_rows)}件を追加しました。")
    else:
        print("⚠️ 新しいデータはありません。")

# ✅ メイン処理
if __name__ == "__main__":
    KEYWORDS = ["ホンダ", "Honda", "HONDA"]

    for kw in KEYWORDS:
        print(f"\n--- MSNニュース取得中: {kw} ---")
        df = get_msn_news(kw)
        print(f"🔹 件数: {len(df)}")
        if not df.empty:
            write_to_spreadsheet(df, SPREADSHEET_ID, SHEET_NAME)
