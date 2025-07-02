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

# ✅ キーワードとスプレッドシート設定
KEYWORD = "ホンダ"
SPREADSHEET_ID = "1AwwMGKMHfduwPkrtsik40lkO1z1T8IU_yd41ku-yPi8"

def format_datetime(dt_obj):
    return dt_obj.strftime("%Y/%m/%d %H:%M")

def get_google_news_with_selenium(keyword: str) -> list[dict]:
    options = Options()
    options.add_argument("--headless")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    url = f"https://news.google.com/search?q={keyword}&hl=ja&gl=JP&ceid=JP:ja"
    driver.get(url)
    time.sleep(5)
    for _ in range(3):
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(2)

    soup = BeautifulSoup(driver.page_source, "html.parser")
    driver.quit()

    articles = soup.find_all("article")
    data = []
    for article in articles:
        try:
            a_tag = article.select_one("a.JtKRv")
            time_tag = article.select_one("time.hvbAAd")
            source_tag = article.select_one("div.vr1PYe")
            title = a_tag.text.strip()
            href = a_tag.get("href")
            url = "https://news.google.com" + href[1:] if href.startswith("./") else href
            dt = datetime.strptime(time_tag.get("datetime"), "%Y-%m-%dT%H:%M:%SZ") + timedelta(hours=9)
            pub_date = format_datetime(dt)
            source = source_tag.text.strip() if source_tag else "N/A"
            data.append({"タイトル": title, "URL": url, "投稿日": pub_date, "引用元": source})
        except Exception as e:
            # print(f"⚠️ Google記事処理エラー: {e}") # デバッグ用
            continue
    print(f"✅ Googleニュース件数: {len(data)} 件")
    return data

def get_yahoo_news_with_selenium(keyword: str) -> list[dict]:
    options = Options()
    options.add_argument("--headless")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    search_url = f"https://news.yahoo.co.jp/search?p={keyword}&ei=utf-8"
    driver.get(search_url)
    time.sleep(5)

    soup = BeautifulSoup(driver.page_source, "html.parser")
    driver.quit()
    articles = soup.find_all("li", class_=re.compile("sc-1u4589e-0"))
    data = []

    for article in articles:
        try:
            title_tag = article.find("div", class_=re.compile("sc-3ls169-0"))
            title = title_tag.text.strip() if title_tag else ""
            link_tag = article.find("a", href=True)
            url = link_tag["href"] if link_tag else ""
            time_tag = article.find("time")
            date_str = time_tag.text.strip() if time_tag else ""
            formatted_date = ""
            if date_str:
                date_str = re.sub(r'\([月火水木金土日]\)', '', date_str).strip()
                try:
                    dt_obj = datetime.strptime(date_str, "%Y/%m/%d %H:%M")
                    formatted_date = format_datetime(dt_obj)
                except ValueError:
                    formatted_date = date_str # フォーマットできない場合はそのまま
            source_text = "Yahoo!"
            if title and url:
                data.append({
                    "タイトル": title,
                    "URL": url,
                    "投稿日": formatted_date if formatted_date else "取得不可",
                    "引用元": source_text
                })
        except Exception as e:
            # print(f"⚠️ Yahoo!記事処理エラー: {e}") # デバッグ用
            continue

    print(f"✅ Yahoo!ニュース件数: {len(data)} 件")
    return data

def get_msn_news_with_selenium(keyword: str) -> list[dict]:
    options = Options()
    options.add_argument("--headless")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    url = f"https://www.bing.com/news/search?q={keyword}&qft=sortbydate%3d'1'&form=YFNR"
    driver.get(url)
    time.sleep(5) # ページが完全にロードされるように少し長めに待つ

    soup = BeautifulSoup(driver.page_source, "html.parser")
    driver.quit()
    data = []

    # 記事のコンテナ要素を特定し、その中のタイトルとURLを抽出
    # .b_algoと.b_ansがニュース記事の個別のブロックになっていると仮定
    # その中のh2タグ内のaタグがタイトルとURLを持つと仮定
    for a_tag in soup.select(".b_algo h2 a, .b_ans h2 a"):
        try:
            title = a_tag.get_text(strip=True)
            url = a_tag.get("href", "").strip()

            if not url.startswith("http"):
                continue # 有効なURLでない場合はスキップ

            # 投稿日は現在のソースからは直接特定が難しいため、「取得不可」とする
            # 必要であれば、記事の親要素から日付情報を探すセレクタを追加
            data.append({
                "タイトル": title,
                "URL": url,
                "投稿日": "取得不可", # 要調査：正確な投稿日を取得するためのセレクタが必要
                "引用元": "MSN"
            })
        except Exception as e:
            print(f"⚠️ MSN記事処理エラー: {e}") # エラーの詳細を出力してデバッグしやすくする
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
    print("\n--- Google News ---")
    google_news_articles = get_google_news_with_selenium(KEYWORD)
    if google_news_articles:
        write_to_spreadsheet(google_news_articles, SPREADSHEET_ID, "Google")

    print("\n--- Yahoo! News ---")
    yahoo_news_articles = get_yahoo_news_with_selenium(KEYWORD)
    if yahoo_news_articles:
        write_to_spreadsheet(yahoo_news_articles, SPREADSHEET_ID, "Yahoo")

    print("\n--- MSN News ---")
    msn_news_articles = get_msn_news_with_selenium(KEYWORD)
    if msn_news_articles:
        write_to_spreadsheet(msn_news_articles, SPREADSHEET_ID, "MSN")
