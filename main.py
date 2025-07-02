import pandas as pd
import time
import os
import re
import json
from datetime import datetime, timedelta
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import gspread

# ✅ 定数
KEYWORDS = ["Honda", "HONDA", "ホンダ"]
SPREADSHEET_ID = "1AwwMGKMHfduwPkrtsik40lkO1z1T8IU_yd41ku-yPi8"
SHEET_NAME = "MSN"
JST = datetime.utcnow() + timedelta(hours=9)

# ✅ Seleniumブラウザ起動
def start_driver():
    options = Options()
    options.add_argument("--headless")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    return webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

# ✅ 相対時間 → 絶対日時へ変換
def parse_pub_date(label: str) -> str:
    try:
        label = label.strip()
        if "分前" in label:
            return (JST - timedelta(minutes=int(re.search(r"\d+", label)[0]))).strftime("%Y/%m/%d %H:%M")
        elif "時間前" in label:
            return (JST - timedelta(hours=int(re.search(r"\d+", label)[0]))).strftime("%Y/%m/%d %H:%M")
        elif "日前" in label:
            return (JST - timedelta(days=int(re.search(r"\d+", label)[0]))).strftime("%Y/%m/%d %H:%M")
        elif re.match(r'\d+月\d+日', label):
            dt = datetime.strptime(f"{JST.year}年{label}", "%Y年%m月%d日")
            return dt.strftime("%Y/%m/%d %H:%M")
        elif re.match(r'\d{4}/\d{1,2}/\d{1,2}', label):
            dt = datetime.strptime(label, "%Y/%m/%d")
            return dt.strftime("%Y/%m/%d %H:%M")
        elif re.match(r'\d{1,2}:\d{2}', label):
            t = datetime.strptime(label, "%H:%M").time()
            dt = datetime.combine(JST.date(), t)
            return dt.strftime("%Y/%m/%d %H:%M")
    except:
        pass
    return label

# ✅ MSNニュース取得
def get_msn_news(keyword: str, driver) -> list:
    print(f"🔍 検索中: {keyword}")
    url = f'https://www.bing.com/news/search?q={keyword}&qft=sortbydate%3d"1"&form=YFNR'
    driver.get(url)
    time.sleep(5)

    soup = BeautifulSoup(driver.page_source, "html.parser")
    data = []
    for card in soup.select("div.news-card"):
        title = card.get("data-title", "").strip()
        url = card.get("data-url", "").strip()
        source = card.get("data-author", "").strip()

        pub_label = ""
        tag = card.find("span", attrs={"aria-label": True})
        if tag and tag.has_attr("aria-label"):
            pub_label = tag["aria-label"].strip()

        pub_date = parse_pub_date(pub_label)

        if title and url:
            data.append({
                "キーワード": keyword,
                "タイトル": title,
                "URL": url,
                "投稿日": pub_date,
                "引用元": source if source else "MSN"
            })
    return data

# ✅ スプレッドシート書き込み
def write_to_spreadsheet(articles: list):
    print("📥 スプレッドシートに書き込み中...")
    
    if not articles:
        print("⚠️ 書き込む記事がありません。")
        return
    
    # 認証（credentials.json or 環境変数）
    if os.path.exists("credentials.json"):
        gc = gspread.service_account(filename="credentials.json")
    else:
        credentials = json.loads(os.environ["GCP_SERVICE_ACCOUNT_KEY"])
        gc = gspread.service_account_from_dict(credentials)

    sh = gc.open_by_key(SPREADSHEET_ID)

    try:
        ws = sh.worksheet(SHEET_NAME)
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(title=SHEET_NAME, rows="1", cols="10")
        ws.append_row(["キーワード", "タイトル", "URL", "投稿日", "引用元"])

    existing = ws.get_all_values()
    existing_urls = set(row[2] for row in existing[1:] if len(row) > 2)

    new_rows = [
        [a["キーワード"], a["タイトル"], a["URL"], a["投稿日"], a["引用元"]]
        for a in articles if a["URL"] not in existing_urls
    ]

    if new_rows:
        ws.append_rows(new_rows, value_input_option="USER_ENTERED")
        print(f"✅ {len(new_rows)}件をスプレッドシートに追記しました。")
    else:
        print("⚠️ すべて既存データのため、追記なし。")

# ✅ メイン処理
def main():
    driver = start_driver()
    all_articles = []
    for kw in KEYWORDS:
        all_articles.extend(get_msn_news(kw, driver))
    driver.quit()

    df = pd.DataFrame(all_articles)
    df.drop_duplicates(subset="URL", inplace=True)
    write_to_spreadsheet(df.to_dict("records"))
    print(f"✅ 全処理完了（合計: {len(df)} 件）")

if __name__ == "__main__":
    main()
