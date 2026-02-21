# -*- coding: utf-8 -*-
"""
從 github 資料夾內所有 Yahoo 序號查詢結果中，
提取「📱 發送到 Telegram 的獎品 📱」區塊的獎項與網址，
整理成 HTML 並可選擇上傳到 GitHub。
"""

import os
import re
import sys
import glob
import subprocess
import time
from pathlib import Path
from datetime import datetime
from urllib.request import urlopen, Request
from urllib.error import URLError, HTTPError
from collections import defaultdict

# 腳本所在目錄 = github 資料夾
BASE_DIR = Path(__file__).resolve().parent
TXT_PATTERN = str(BASE_DIR / "Yahoo序號連結查詢結果_*.txt")
OUTPUT_HTML = BASE_DIR / "Telegram獎品網址整理.html"
OUTPUT_TXT = BASE_DIR / "Telegram獎品網址清單.txt"
OUTPUT_COUPON = BASE_DIR / "allmysteven.html"  # 電子券清單（與 Telegram 獎品同步）
EXPIRY_CACHE = BASE_DIR / "expiry_cache.txt"   # 兌換期間至快取（url -> 日期）


def fetch_expiry_date(url):
    """從 txp.rs 兌換券頁面爬取「兌換期間至」日期。"""
    if "txp.rs" not in url:
        return None
    try:
        req = Request(url, headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"})
        with urlopen(req, timeout=15) as resp:
            html = resp.read().decode("utf-8", errors="ignore")
        m = re.search(r"兌換期間至[\s\S]*?([\d]{4}\.[\d]{2}\.[\d]{2})", html)
        if m:
            return m.group(1).strip()
    except (URLError, HTTPError, Exception) as e:
        pass
    return None


def load_expiry_cache():
    """載入兌換期間至快取。"""
    cache = {}
    if EXPIRY_CACHE.exists():
        try:
            for line in EXPIRY_CACHE.read_text(encoding="utf-8").strip().split("\n"):
                if "\t" in line:
                    url, date = line.split("\t", 1)
                    cache[url.strip()] = date.strip()
        except Exception:
            pass
    return cache


def save_expiry_cache(cache):
    """儲存兌換期間至快取。"""
    lines = [f"{url}\t{date}" for url, date in sorted(cache.items())]
    EXPIRY_CACHE.write_text("\n".join(lines), encoding="utf-8")


def enrich_prizes_with_expiry(entries, verbose=True):
    """為每個獎品補充兌換期間至，使用快取避免重複請求。"""
    cache = load_expiry_cache()
    updated = False
    for rec in entries:
        for p in rec["prizes"]:
            url = p["link"].strip()
            if url in cache:
                p["expiry"] = cache[url]
            else:
                expiry = fetch_expiry_date(url)
                p["expiry"] = expiry if expiry else ""
                if expiry:
                    cache[url] = expiry
                    updated = True
                if verbose and "txp.rs" in url:
                    print(f"  取得兌換期間: {p['title'][:20]}... -> {p['expiry']}")
                time.sleep(0.5)  # 避免請求過快
    if updated:
        save_expiry_cache(cache)
    return entries


def parse_telegram_section(content):
    """從單一檔案內容中解析「發送到 Telegram 的獎品」區塊。"""
    marker = "📱 發送到 Telegram 的獎品 📱"
    if marker not in content:
        return None, []

    start = content.find(marker)
    block = content[start:]
    end_marker = "--- 批次流程結束 ---"
    if end_marker in block:
        block = block[: block.find(end_marker)]

    date_m = re.search(r"發送日期:\s*(.+?)(?:\n|$)", block)
    send_date = date_m.group(1).strip() if date_m else ""

    # 獎品項目： [N] Profile X \n 標題: ... \n 時間: ... \n 連結: ...
    pattern = re.compile(
        r"\[\s*(\d+)\s*\]\s*Profile\s*(\d+)\s*\n\s*標題:\s*(.+?)\s*\n\s*時間:\s*(.+?)\s*\n\s*連結:\s*(.+?)(?=\n\s*\[\s*\d|\n\n|\n===|\Z)",
        re.DOTALL,
    )
    prizes = []
    for m in pattern.finditer(block):
        num, profile, title, time_str, link = m.groups()
        prizes.append({
            "num": num.strip(),
            "profile": profile.strip(),
            "title": title.strip(),
            "time": time_str.strip(),
            "link": link.strip(),
        })
    return send_date, prizes


def collect_all_prizes():
    """掃描 github 資料夾內所有 txt，收集 Telegram 區塊的獎項。"""
    files = sorted(glob.glob(TXT_PATTERN))
    all_entries = []

    for fpath in files:
        fname = os.path.basename(fpath)
        try:
            with open(fpath, "r", encoding="utf-8") as f:
                content = f.read()
        except Exception as e:
            print(f"讀取失敗 {fpath}: {e}")
            continue
        send_date, prizes = parse_telegram_section(content)
        if send_date is None:
            continue
        all_entries.append({"file": fname, "send_date": send_date, "prizes": prizes})

    return all_entries


def build_html(entries):
    """產生整理後的 HTML 頁面（所有網址彙總）。"""
    flat = []
    seen_urls = set()
    for rec in entries:
        send_date = rec["send_date"]
        for p in rec["prizes"]:
            if p["link"] in seen_urls:
                continue
            seen_urls.add(p["link"])
            flat.append({**p, "send_date": send_date})
    flat = _sort_and_group_prizes(flat)
    has_any_expiry = any(p.get("expiry") for p in flat)
    total_count = len(flat)
    rows = []
    for i, p in enumerate(flat, 1):
        send_date = p.get("send_date", "")
        if has_any_expiry:
            rows.append(
                f"<tr><td>{i}</td><td>{p['title']}</td><td>{p.get('expiry') or ''}</td>"
                f"<td>{send_date}</td><td>{p['time']}</td><td>Profile {p['profile']}</td>"
                f'<td><a href="{p["link"]}" target="_blank" rel="noopener">開啟</a></td></tr>'
            )
        else:
            rows.append(
                f"<tr><td>{i}</td><td>{p['title']}</td>"
                f"<td>{send_date}</td><td>{p['time']}</td><td>Profile {p['profile']}</td>"
                f'<td><a href="{p["link"]}" target="_blank" rel="noopener">開啟</a></td></tr>'
            )

    html = f"""<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>📱 發送到 Telegram 的獎品 - 網址整理</title>
    <style>
        * {{ box-sizing: border-box; }}
        body {{ font-family: "Microsoft JhengHei", "Segoe UI", sans-serif; margin: 20px; background: #1a1a2e; color: #eee; }}
        h1 {{ color: #00d9ff; text-align: center; }}
        .summary {{ text-align: center; margin-bottom: 24px; font-size: 1.1em; color: #aaa; }}
        .table-wrap {{ overflow-x: auto; background: #16213e; padding: 16px; border-radius: 12px; }}
        table {{ width: 100%; border-collapse: collapse; }}
        th, td {{ border: 1px solid #0f3460; padding: 10px; text-align: left; }}
        th {{ background: #0f3460; color: #00d9ff; }}
        tr:nth-child(even) {{ background: #1a1a2e; }}
        a {{ color: #00d9ff; }}
    </style>
</head>
<body>
    <h1>📱 發送到 Telegram 的獎品 📱</h1>
    <p class="summary">最後更新：{datetime.now().strftime('%Y-%m-%d %H:%M')}｜共 {len(entries)} 個查詢日、{total_count} 筆獎項網址（同類型、依到期日排序）</p>
    <div class="table-wrap">
        <table>
            <thead><tr><th>#</th><th>標題</th>{'<th>兌換期間至</th>' if has_any_expiry else ''}<th>發送日期</th><th>時間</th><th>Profile</th><th>連結</th></tr></thead>
            <tbody>
                {''.join(rows)}
            </tbody>
        </table>
    </div>
</body>
</html>
"""
    return html


def _sort_and_group_prizes(flat):
    """同類型放一起，依到期日由近到遠排序。"""
    groups = defaultdict(list)
    for p in flat:
        groups[p["title"]].append(p)
    # 每組內依到期日排序（無到期日的放最後）
    FAR = "9999.99.99"
    for title in groups:
        groups[title].sort(key=lambda p: p.get("expiry") or FAR)
    # 各組依「該組最早到期日」排序，到期日近的組排前面
    def group_min_expiry(items):
        expiries = [p.get("expiry") for p in items if p.get("expiry")]
        return min(expiries) if expiries else FAR
    sorted_pairs = sorted(groups.items(), key=lambda x: group_min_expiry(x[1]))
    result = []
    for title, items in sorted_pairs:
        result.extend(items)
    return result


def build_allmysteven_html(entries):
    """產生電子券清單 allmysteven.html（品項、兌換期間至、使用連結）。"""
    flat = []
    seen_urls = set()
    for rec in entries:
        for p in rec["prizes"]:
            if p["link"] in seen_urls:
                continue
            seen_urls.add(p["link"])
            flat.append(p)
    flat = _sort_and_group_prizes(flat)
    has_any_expiry = any(p.get("expiry") for p in flat)
    rows = []
    for i, p in enumerate(flat, 1):
        expiry = p.get("expiry") or ""
        if has_any_expiry:
            rows.append(
                f'<tr><td>{i}</td><td>{p["title"]}</td><td>{expiry}</td>'
                f'<td><a href="{p["link"]}" target="_blank" rel="noopener" class="btn">使用</a></td></tr>'
            )
        else:
            rows.append(
                f'<tr><td>{i}</td><td>{p["title"]}</td>'
                f'<td><a href="{p["link"]}" target="_blank" rel="noopener" class="btn">使用</a></td></tr>'
            )
    th_expiry = '<th>兌換期間至</th>' if has_any_expiry else ''
    thead = f'<tr><th>#</th><th>品項名稱</th>{th_expiry}<th>操作</th></tr>'
    html = f"""<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>我的電子券清單</title>
    <style>
        body {{ font-family: "Microsoft JhengHei", sans-serif; max-width: 900px; margin: 20px auto; padding: 20px; background: #1a1a2e; color: #eee; }}
        h1 {{ color: #00d9ff; }}
        a {{ color: #00d9ff; }}
        .summary {{ color: #888; margin-bottom: 20px; }}
        table {{ width: 100%; border-collapse: collapse; }}
        th, td {{ border: 1px solid #0f3460; padding: 12px; text-align: left; }}
        th {{ background: #0f3460; color: #00d9ff; }}
        .btn {{ display: inline-block; padding: 6px 16px; background: #007bff; color: white !important; text-decoration: none; border-radius: 6px; }}
        .btn:hover {{ background: #0056b3; }}
    </style>
</head>
<body>
    <h1>🎟️ 我的電子商品券</h1>
    <p class="summary">最後更新：{datetime.now().strftime('%Y-%m-%d %H:%M')}</p>
    <table>
        <thead>{thead}</thead>
        <tbody>{''.join(rows)}</tbody>
    </table>
    <p class="summary">目前共有商品券：{len(flat)} 張</p>
</body>
</html>
"""
    return html


def build_txt_url_list(entries):
    """產生純網址清單（每行一個 URL）。"""
    seen = set()
    lines = []
    for rec in entries:
        for p in rec["prizes"]:
            link = p["link"].strip()
            if link and link not in seen:
                seen.add(link)
                lines.append(link)
    return "\n".join(lines)


def git_upload():
    """執行 git add、commit、push。"""
    try:
        subprocess.run(["git", "add", "Telegram獎品網址整理.html", "Telegram獎品網址清單.txt", "allmysteven.html", "index.html"], 
                       cwd=BASE_DIR, check=True, capture_output=True, text=True)
        subprocess.run(["git", "commit", "-m", f"更新 Telegram 獎品網址整理 {datetime.now().strftime('%Y-%m-%d %H:%M')}"], 
                       cwd=BASE_DIR, check=True, capture_output=True, text=True)
        subprocess.run(["git", "push", "origin", "main"], 
                       cwd=BASE_DIR, check=True, capture_output=True, text=True)
        return True
    except subprocess.CalledProcessError as e:
        print(f"Git 執行失敗: {e}")
        if e.stderr:
            print(e.stderr)
        return False


def main():
    do_upload = "--upload" in sys.argv or "-u" in sys.argv

    print("正在掃描 github 資料夾內的 Yahoo 序號查詢結果...")
    entries = collect_all_prizes()
    if not entries:
        print("未找到任何「📱 發送到 Telegram 的獎品 📱」區塊。")
        return

    print(f"共 {len(entries)} 個日期的 Telegram 獎項區塊。")
    total_prizes = sum(len(e["prizes"]) for e in entries)
    print(f"獎項總筆數: {total_prizes}")

    skip_fetch = "--no-fetch" in sys.argv
    if skip_fetch:
        print("略過爬取兌換期間至（僅用快取）。")
        cache = load_expiry_cache()
        for rec in entries:
            for p in rec["prizes"]:
                p["expiry"] = cache.get(p["link"].strip(), "")
    else:
        print("正在爬取兌換期間至（首次較慢，之後會用快取）...")
        entries = enrich_prizes_with_expiry(entries, verbose=False)

    html = build_html(entries)
    OUTPUT_HTML.write_text(html, encoding="utf-8")
    print(f"HTML 已寫入: {OUTPUT_HTML}")

    coupon_html = build_allmysteven_html(entries)
    OUTPUT_COUPON.write_text(coupon_html, encoding="utf-8")
    print(f"電子券清單已寫入: {OUTPUT_COUPON}")

    txt_content = build_txt_url_list(entries)
    OUTPUT_TXT.write_text(txt_content, encoding="utf-8")
    print(f"網址清單已寫入: {OUTPUT_TXT}")

    if do_upload:
        print("正在上傳到 GitHub...")
        if git_upload():
            print("✅ 上傳完成！")
        else:
            print("❌ 上傳失敗，請手動執行 git push。")
    else:
        print("\n若要上傳到 GitHub，請執行：")
        print("  python extract_and_upload.py --upload")
        print("或：")
        print("  python extract_and_upload.py -u")

    print("完成。")


if __name__ == "__main__":
    main()
