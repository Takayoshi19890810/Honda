# -*- coding: utf-8 -*-
import os
import json
import time
import re
import random
import argparse
import requests
from datetime import datetime, timedelta
from email.utils import parsedate_to_datetime

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
import gspread

# =========================
# 既定（引数/環境変数で上書き可）
# =========================
DEFAULT_KEYWORD = "ホンダ"  # 例: "ホンダ", "マツダ", "日産" など
DEFAULT_SPREADSHEET_ID = "1AwwMGKMHfduwPkrtsik40lkO1z1T8IU_yd41ku-yPi8"  # ホンダ用

# =========================
# 共通ユーティリティ
# =========================
def format_datetime(dt_obj: datetime) -> str:
    return dt_obj.strftime("%Y/%m/%d %H:%M")

TIME_RE = re.compile(r"(\d+)\s*(分|時間|日)\s*前")  # 例: "7 時間前", "15分前"

def parse_relative_time(pub_label: str, base_time: datetime) -> str:
    """
    "2時間前" や "15分前" のような相対表現をJST日時文字列に変換
    """
    if not pub_label:
        return "取得不可"
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
            dt = datetime.strptime(pub_label, "%m月%d日")
            dt = dt.replace(year=base_time.year)
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

def get_last_modified_datetime(url: str) -> str:
    """
    HEADの Last-Modified をJSTにして返す
    """
    try:
        response = requests.head(url, timeout=5)
        if 'Last-Modified' in response.headers:
            dt = parsedate_to_datetime(response.headers['Last-Modified'])
            jst = dt + timedelta(hours=9)
            return format_datetime(jst)
    except:
        pass
    return "取得不可"

def make_driver() -> webdriver.Chrome:
    options = Options()
    options.add_argument("--headless=new")  # 新ヘッドレス
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--window-size=1280,2000")
    options.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    )
    return webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

def clean_source_text(text: str) -> str:
    """
    'Merkmal（メルクマール） on MSN 1 時間' → 'Merkmal（メルクマール）'
    """
    if not text:
        return ""
    t = re.sub(r"\bon\s+MSN\b", "", text, flags=re.IGNORECASE)   # "on MSN" 除去
    t = TIME_RE.sub("", t)                                       # "◯時間前" 等 除去
    t = t.replace("・", " ").replace("•", " ").replace("·", " ")
    return re.sub(r"\s{2,}", " ", t).strip()

def find_relative_label(container) -> str:
    """
    コンテナ周辺から '◯分/時間/日前' 文字列または ISO datetime を探す
    """
    # aria-label
    for el in container.select("[aria-label]"):
        lab = el.get("aria-label", "").strip()
        if TIME_RE.search(lab):
            return lab
    # time要素
    for el in container.select("time"):
        t = (el.get_text(strip=True) or "").strip()
        if TIME_RE.search(t):
            return t
        if el.get("datetime"):  # ISO datetime (UTC Z)
            return el.get("datetime")
    # 汎用テキスト
    texts = [
        container.get_text(" ", strip=True),
        (container.parent.get_text(" ", strip=True) if container.parent else "")
    ]
    for txt in texts:
        m = TIME_RE.search(txt)
        if m:
            return m.group(0)
    return ""

# =========================
# 各サイトのスクレイパ
# =========================
def get_google_news_with_selenium(keyword: str) -> list[dict]:
    driver = make_driver()
    url = f"https://news.google.com/search?q={keyword}&hl=ja&gl=JP&ceid=JP:ja"
    driver.get(url)
    time.sleep(5)
    for _ in range(3):
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(1.2)

    soup = BeautifulSoup(driver.page_source, "html.parser")
    driver.quit()

    articles = soup.find_all("article")
    data: list[dict] = []
    for article in articles:
        try:
            a_tag = article.select_one("a.JtKRv")
            time_tag = article.select_one("time.hvbAAd")
            source_tag = article.select_one("div.vr1PYe")
            if not a_tag or not time_tag:
                continue

            title = a_tag.text.strip()
            href = a_tag.get("href")
            url = "https://news.google.com" + href[1:] if href and href.startswith("./") else href

            dt = datetime.strptime(time_tag.get("datetime"), "%Y-%m-%dT%H:%M:%SZ") + timedelta(hours=9)
            pub_date = format_datetime(dt)
            source = source_tag.text.strip() if source_tag else "N/A"

            if title and url:
                data.append({"タイトル": title, "URL": url, "投稿日": pub_date, "引用元": source})
        except:
            continue
    print(f"✅ Googleニュース件数: {len(data)} 件")
    return data

def get_yahoo_news_with_selenium(keyword: str) -> list[dict]:
    driver = make_driver()
    search_url = (
        f"https://news.yahoo.co.jp/search?p={keyword}"
        f"&ei=utf-8&categories=domestic,world,business,it,science,life,local"
    )
    driver.get(search_url)
    time.sleep(5)

    soup = BeautifulSoup(driver.page_source, "html.parser")
    driver.quit()

    articles = soup.find_all("li", class_=re.compile("sc-1u4589e-0"))
    articles_data: list[dict] = []

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
                except:
                    formatted_date = date_str

            # 引用元（媒体名）
            source_text = ""
            source_tag = article.find("div", class_="sc-n3vj8g-0 yoLqH")
            if source_tag:
                inner = source_tag.find("div", class_="sc-110wjhy-8 bsEjY")
                if inner and inner.span:
                    candidate = inner.span.text.strip()
                    if not candidate.isdigit():
                        source_text = candidate
            if not source_text or source_text.isdigit():
                alt_spans = article.find_all(["span", "div"], string=True)
                for s in alt_spans:
                    text = s.text.strip()
                    if 2 <= len(text) <= 20 and not text.isdigit() and re.search(r'[ぁ-んァ-ン一-龥A-Za-z]', text):
                        source_text = text
                        break

            if title and url:
                articles_data.append({
                    "タイトル": title,
                    "URL": url,
                    "投稿日": formatted_date if formatted_date else "取得不可",
                    "引用元": source_text or "Yahoo"
                })
        except:
            continue

    print(f"✅ Yahoo!ニュース件数: {len(articles_data)} 件")
    return articles_data

def get_msn_news_with_selenium(keyword: str) -> list[dict]:
    """
    MSN(Bingニュース) 強化版：
    - Cookie同意対応
    - a.title / a[data-title] 両対応
    - 周辺テキストから媒体名と相対時刻を分離抽出
    - 'on MSN' や '◯時間前' を引用元から除去
    - 取れない日時は Last-Modified で補完
    """
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.common.exceptions import TimeoutException

    now = datetime.utcnow() + timedelta(hours=9)
    driver = make_driver()
    url = (
        f"https://www.bing.com/news/search?q={keyword}"
        "&qft=sortbydate%3D%271%27"
        "&setlang=ja&cc=JP&FORM=HDRSC6"
    )
    driver.get(url)

    # Cookie同意
    try:
        WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.ID, "bnp_btn_accept"))
        ).click()
    except TimeoutException:
        pass

    # 記事読み込み待ち
    try:
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "a.title, a[data-title]"))
        )
    except TimeoutException:
        time.sleep(2)

    # Lazy Load対策スクロール
    for _ in range(4):
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(1.0)

    soup = BeautifulSoup(driver.page_source, "html.parser")
    driver.quit()

    data: list[dict] = []
    anchors = soup.select("a.title, a[data-title]")

    for a in anchors:
        try:
            title = (a.get("data-title") or a.get_text(strip=True) or "").strip()
            href = a.get("href") or ""
            if not (title and href):
                continue

            parent = a.find_parent(["div", "li"]) or a.parent
            raw_source = ""
            if parent:
                s_el = parent.select_one("div.source, span.source")
                if s_el:
                    raw_source = s_el.get_text(" ", strip=True)

            # 相対時刻 or ISO datetime
            label = find_relative_label(parent or a)

            # 投稿日を決定
            pub_date = "取得不可"
            if label:
                if "T" in label and ":" in label and label.endswith("Z"):
                    # ISO → JST
                    try:
                        dt = datetime.strptime(label, "%Y-%m-%dT%H:%M:%SZ") + timedelta(hours=9)
                        pub_date = format_datetime(dt)
                    except:
                        pass
                else:
                    # 相対 → JST
                    pub_date = parse_relative_time(label, now)

            if pub_date == "取得不可":
                pub_date = get_last_modified_datetime(href)

            source = clean_source_text(raw_source) or "MSN"

            data.append({
                "タイトル": title,
                "URL": href,
                "投稿日": pub_date,
                "引用元": source
            })
        except:
            continue

    print(f"✅ MSNニュース件数: {len(data)} 件")
    return data

# =========================
# スプレッドシート書き込み
# =========================
def write_to_spreadsheet(articles: list[dict], spreadsheet_id: str, worksheet_name: str):
    """
    既存URLと重複しないものだけ追記。シートが無ければ作成。
    認証: GCP_SERVICE_ACCOUNT_KEY（環境） or credentials.json（ローカル）
    """
    credentials_json_str = os.environ.get('GCP_SERVICE_ACCOUNT_KEY')
    credentials = json.loads(credentials_json_str) if credentials_json_str else json.load(open('credentials.json'))
    gc = gspread.service_account_from_dict(credentials)

    for attempt in range(5):
        try:
            sh = gc.open_by_key(spreadsheet_id)
            try:
                ws = sh.worksheet(worksheet_name)
            except gspread.exceptions.WorksheetNotFound:
                ws = sh.add_worksheet(title=worksheet_name, rows="1", cols="4")
                ws.append_row(['タイトル', 'URL', '投稿日', '引用元'])

            existing = ws.get_all_values()
            existing_urls = set(row[1] for row in existing[1:] if len(row) > 1)

            new_rows = [[a['タイトル'], a['URL'], a['投稿日'], a['引用元']]
                        for a in articles if a['URL'] not in existing_urls]

            if new_rows:
                ws.append_rows(new_rows, value_input_option='USER_ENTERED')
                print(f"✅ {len(new_rows)}件をスプレッドシート「{worksheet_name}」に追記しました。")
            else:
                print(f"⚠️ 追記すべき新しいデータはありません。（{worksheet_name}）")
            return
        except gspread.exceptions.APIError as e:
            print(f"⚠️ Google API Error (attempt {attempt + 1}/5): {e}")
            time.sleep(5 + random.random() * 5)

    raise RuntimeError("❌ Googleスプレッドシートへの書き込みに失敗しました（5回試行しても成功せず）")

# =========================
# 設定の解決（引数/環境/既定）
# =========================
def resolve_config() -> tuple[str, str]:
    parser = argparse.ArgumentParser()
    parser.add_argument("--keyword", type=str, default=None, help="検索キーワード（例: ホンダ）")
    parser.add_argument("--sheet", type=str, default=None, help="スプレッドシートID")
    args = parser.parse_args()

    keyword = args.keyword or os.getenv("NEWS_KEYWORD") or DEFAULT_KEYWORD
    spreadsheet_id = args.sheet or os.getenv("SPREADSHEET_ID") or DEFAULT_SPREADSHEET_ID
    print(f"🔎 キーワード: {keyword}")
    print(f"📄 SPREADSHEET_ID: {spreadsheet_id}")
    return keyword, spreadsheet_id

# =========================
# エントリポイント
# =========================
if __name__ == "__main__":
    keyword, spreadsheet_id = resolve_config()

    print("\n--- Google News ---")
    google_news_articles = get_google_news_with_selenium(keyword)
    if google_news_articles:
        write_to_spreadsheet(google_news_articles, spreadsheet_id, "Google")

    print("\n--- Yahoo! News ---")
    yahoo_news_articles = get_yahoo_news_with_selenium(keyword)
    if yahoo_news_articles:
        write_to_spreadsheet(yahoo_news_articles, spreadsheet_id, "Yahoo")

    print("\n--- MSN News ---")
    msn_news_articles = get_msn_news_with_selenium(keyword)
    if msn_news_articles:
        write_to_spreadsheet(msn_news_articles, spreadsheet_id, "MSN")
