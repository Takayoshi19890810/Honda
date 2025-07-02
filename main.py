import os
import json
import time
import re
import random
import requests
from datetime import datetime, timedelta
from email.utils import parsedate_to_datetime

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
import gspread

KEYWORD = "ホンダ"
SPREADSHEET_ID = "1AwwMGKMHfduwPkrtsik40lkO1z1T8IU_yd41ku-yPi8"

def format_datetime(dt_obj):
    return dt_obj.strftime("%Y/%m/%d %H:%M")

def parse_relative_time(pub_label: str, base_time: datetime) -> str:
    pub_label = pub_label.strip().lower()
    try:
        if "分前" in pub_label or "minute" in pub_label:
            m = re.search(r"(\d+)", pub_label)
            if m:
                dt = base_time - timedelta(minutes=int(m.group(1)))
                return format_datetime(dt)
        elif "時間前" in pub_label or "hour" in pub_label:
            h = re.search(r"(\d+)", pub_label)
            if h:
                dt = base_time - timedelta(hours=int(h.group(1)))
                return format_datetime(dt)
        elif "日前" in pub_label or "day" in pub_label:
            d = re.search(r"(\d+)", pub_label)
            if d:
                dt = base_time - timedelta(days=int(d.group(1)))
                return format_datetime(dt)
        elif re.match(r'\d+月\d+日', pub_label):
            dt = datetime.strptime(f"{base_time.year}年{pub_label}", "%Y年%m月%d日")
            return format_datetime(dt)
        elif re.match(r'\d{4}/\d{1,2}/\d{1,2}', pub_label):
            dt = datetime.strptime(pub_label, "%Y/%m/%d")
            return format_datetime(dt)
        elif re.match(r'\d{1,2}:\d{2}', pub_label):
            t = datetime.strptime(pub_label, "%H:%M").time()
            dt = datetime.combine(base_time.date(), t)
            if dt > base_time:
                dt -= timedelta(days=1)
            return format_datetime(dt)
    except:
        pass
    return "取得不可"

def get_last_modified_datetime(url):
    try:
        response = requests.head(url, timeout=5)
        if 'Last-Modified' in response.headers:
            dt = parsedate_to_datetime(response.headers['Last-Modified'])
            jst = dt.astimezone(tz=timedelta(hours=9))
            return format_datetime(jst)
    except:
        pass
    return "取得不可"

def get_msn_news_with_selenium(keyword: str) -> list[dict]:
    now = datetime.utcnow() + timedelta(hours=9)
    options = Options()
    options.add_argument("--headless")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    url = f"https://www.bing.com/news/search?q={keyword}&qft=sortbydate%3d'1'&form=YFNR"
    driver.get(url)
    time.sleep(5)

    soup = BeautifulSoup(driver.page_source, "html.parser")
    driver.quit()

    data = []
    articles = soup.select("a.t_s")
    for article in articles:
        try:
            title = article.text.strip()
            url = article.get("href", "").strip()
            source_tag = article.find_next("div", class_="source")
            source = source_tag.text.strip() if source_tag else "MSN"
            time_tag = article.find_next("span")
            pub_label = time_tag.text.strip() if time_tag else ""

            pub_date = parse_relative_time(pub_label, now)
            if pub_date == "取得不可" and url:
                pub_date = get_last_modified_datetime(url)

            if title and url:
                data.append({
                    "タイトル": title,
                    "URL": url,
                    "投稿日": pub_date,
                    "引用元": source
                })
        except Exception as e:
            print(f"⚠️ MSN記事処理エラー: {e}")
            continue

    print(f"✅ MSNニュース件数: {len(data)} 件")
    return data

def write_to_spreadsheet(articles: list[dict], spreadsheet_id: str, worksheet_name: str):
    credentials_json_str = os.environ.get('GCP_SERVICE_ACCOUNT_KEY')
    credentials = json.loads(credentials_json_str) if credentials_json_str else json.load(open('credentials.json'))
    gc = gspread.service_account_from_dict(credentials)

    for attempt in range(5):
        try:
            sh = gc.open_by_key(spreadsheet_id)
            try:
                worksheet = sh.worksheet(worksheet_name)
            except gspread.exceptions.WorksheetNotFound:
                worksheet = sh.add_worksheet(title=worksheet_name, rows="1", cols="4")
                worksheet.append_row(['タイトル', 'URL', '投稿日', '引用元'])

            existing_data = worksheet.get_all_values()
            existing_urls = set(row[1] for row in existing_data[1:] if len(row) > 1)

            new_data = [[a['タイトル'], a['URL'], a['投稿日'], a['引用元']] for a in articles if a['URL'] not in existing_urls]
            if new_data:
                worksheet.append_rows(new_data, value_input_option='USER_ENTERED')
                print(f"✅ {len(new_data)}件をスプレッドシートに追記しました。")
            else:
                print("⚠️ 追記すべき新しいデータはありません。")
            return
        except gspread.exceptions.APIError as e:
            print(f"⚠️ Google API Error (attempt {attempt + 1}/5): {e}")
            time.sleep(5 + random.random() * 5)

    raise RuntimeError("❌ Googleスプレッドシートへの書き込みに失敗しました（5回試行しても成功せず）")

if __name__ == "__main__":
    print("\n--- MSN News ---")
    msn_news_articles = get_msn_news_with_selenium(KEYWORD)
    if msn_news_articles:
        write_to_spreadsheet(msn_news_articles, SPREADSHEET_ID, "MSN")
