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
from concurrent.futures import ThreadPoolExecutor

# 腳本所在目錄 = github 資料夾
BASE_DIR = Path(__file__).resolve().parent
TXT_PATTERN = str(BASE_DIR / "Yahoo序號連結查詢結果_*.txt")
OUTPUT_HTML = BASE_DIR / "Telegram獎品網址整理.html"
OUTPUT_TXT = BASE_DIR / "Telegram獎品網址清單.txt"
OUTPUT_COUPON = BASE_DIR / "allmysteven.html"  # 電子券清單（與 Telegram 獎品同步）
EXPIRY_CACHE = BASE_DIR / "expiry_cache.txt"   # 兌換期間至快取（url -> 日期）
MAX_THREADS = 50  # 多執行緒檢查數量


def fetch_voucher_info(url):
    """從 txp.rs 兌換券頁面爬取：兌換期間至、是否已使用。"""
    if "txp.rs" not in url:
        return None, ""
    try:
        req = Request(url, headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"})
        with urlopen(req, timeout=15) as resp:
            html = resp.read().decode("utf-8", errors="ignore")
        
        expiry = None
        # 爬取日期格式: 2026.05.23 或 2026/05/23
        m = re.search(r"([\d]{4}[\./][\d]{2}[\./][\d]{2})", html)
        if m:
            expiry = m.group(1).strip().replace("/", ".")
            
        # 檢測「已兌換/已使用」狀態的精準標記 (排除注意事項中的文字)
        used_markers = [
            "usededn",           # Edenred CSS marker
            "StatusOverlayBg",   # Edenred Status overlay
            "stamp_used.png",    # 'Used' stamp image
            "StampStatus",       # Stamp status element
            'class="stamp used"', # Specific CSS class
            'id="balance">剩餘：0' # Balance is zero (specifically in the balance element)
        ]
        is_used = any(x in html for x in used_markers)
        
        used = "已兌換" if is_used else ""
        return expiry, used
    except (URLError, HTTPError, Exception):
        pass
    return None, ""


def load_expiry_cache():
    """載入兌換期間至快取（格式：url\texpiry 或 url\texpiry\tstatus）。"""
    cache = {}  # url -> (expiry, status)
    if EXPIRY_CACHE.exists():
        try:
            for line in EXPIRY_CACHE.read_text(encoding="utf-8").strip().split("\n"):
                if "\t" in line:
                    parts = line.split("\t", 2)
                    url = parts[0].strip()
                    expiry = parts[1].strip() if len(parts) > 1 else ""
                    status = parts[2].strip() if len(parts) > 2 else ""
                    cache[url] = (expiry, status)
        except Exception:
            pass
    return cache


def save_expiry_cache(cache):
    """儲存兌換期間至快取。"""
    lines = []
    for url in sorted(cache.keys()):
        expiry, status = cache[url]
        if status:
            lines.append(f"{url}\t{expiry}\t{status}")
        else:
            lines.append(f"{url}\t{expiry}")
    EXPIRY_CACHE.write_text("\n".join(lines), encoding="utf-8")


def enrich_prizes_with_expiry(entries, verbose=True, force_refresh=False):
    """為每個獎品補充兌換期間至、是否已使用，使用多執行緒加速。"""
    cache = load_expiry_cache() if not force_refresh else {}
    
    # 收集所有需要檢查的 prize 物件
    to_check = []
    for rec in entries:
        for p in rec["prizes"]:
            url = p["link"].strip()
            if "txp.rs" not in url:
                continue
            if url in cache and not force_refresh:
                expiry, status = cache[url]
                p["expiry"] = expiry
                p["used"] = status
            else:
                to_check.append(p)

    if not to_check:
        return entries

    print(f"  -> 共有 {len(to_check)} 筆需要連網檢查狀態 (使用 {MAX_THREADS} 執行緒)...")
    
    updated = False
    
    def process_p(p):
        url = p["link"].strip()
        expiry, used = fetch_voucher_info(url)
        p["expiry"] = expiry if expiry else ""
        p["used"] = used
        if verbose:
            tag = f"[{used}]" if used else "[可兌換]"
            print(f"  取得: {p['title'][:15]}... -> {tag}")
        return url, (p["expiry"], p["used"])

    with ThreadPoolExecutor(max_workers=MAX_THREADS) as executor:
        results = list(executor.map(process_p, to_check))
    
    for url, result_val in results:
        cache[url] = result_val
        updated = True

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
        used = p.get("used", "") or ""
        row_class = ' class="used-row"' if used else ""
        link_td = '<span class="used-badge">已兌換</span>' if used else f'<a href="{p["link"]}" target="_blank" rel="noopener">開啟</a>'
        if has_any_expiry:
            rows.append(
                f"<tr{row_class}><td>{i}</td><td>{p['title']}</td><td>{p.get('expiry') or ''}</td>"
                f"<td>{send_date}</td><td>{p['time']}</td><td>Profile {p['profile']}</td>"
                f"<td>{link_td}</td></tr>"
            )
        else:
            rows.append(
                f"<tr{row_class}><td>{i}</td><td>{p['title']}</td>"
                f"<td>{send_date}</td><td>{p['time']}</td><td>Profile {p['profile']}</td>"
                f"<td>{link_td}</td></tr>"
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
        .used-row {{ opacity: 0.6; background: #1a2a1a !important; }}
        .used-badge {{ color: #888; }}
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
    """產生電子券清單 allmysteven.html（未兌換在上、已兌換區塊在最下）。"""
    flat = []
    seen_urls = set()
    for rec in entries:
        for p in rec["prizes"]:
            if p["link"] in seen_urls:
                continue
            seen_urls.add(p["link"])
            flat.append(p)
    # 分為未兌換、已兌換
    available = [p for p in flat if not p.get("used")]
    used_list = [p for p in flat if p.get("used")]
    available = _sort_and_group_prizes(available)
    used_list = _sort_and_group_prizes(used_list)
    has_any_expiry = any(p.get("expiry") for p in flat)
    th_expiry = '<th>兌換期間至</th>' if has_any_expiry else ''
    thead = f'<tr><th>#</th><th>品項名稱</th>{th_expiry}<th>操作</th></tr>'
    # 未兌換區塊
    rows_available = []
    for i, p in enumerate(available, 1):
        expiry = p.get("expiry") or ""
        btn = f'<a href="{p["link"]}" target="_blank" rel="noopener" class="btn">使用</a>'
        if has_any_expiry:
            rows_available.append(f'<tr><td>{i}</td><td>{p["title"]}</td><td>{expiry}</td><td>{btn}</td></tr>')
        else:
            rows_available.append(f'<tr><td>{i}</td><td>{p["title"]}</td><td>{btn}</td></tr>')
    # 已兌換區塊（最下面）
    rows_used = []
    for i, p in enumerate(used_list, 1):
        expiry = p.get("expiry") or ""
        link_td = f'<a href="{p["link"]}" target="_blank" rel="noopener" class="link-used">查看</a>'
        if has_any_expiry:
            rows_used.append(f'<tr class="used-row"><td>{i}</td><td>{p["title"]}</td><td>{expiry}</td><td>{link_td}</td></tr>')
        else:
            rows_used.append(f'<tr class="used-row"><td>{i}</td><td>{p["title"]}</td><td>{link_td}</td></tr>')
    section_available = f"""
    <h2>可兌換（{len(available)} 張）</h2>
    <table>
        <thead>{thead}</thead>
        <tbody>{''.join(rows_available)}</tbody>
    </table>
    """ if available else ""
    section_used = f"""
    <h2 class="section-used">已兌換（{len(used_list)} 張）</h2>
    <div class="used-block">
        <table>
            <thead>{thead}</thead>
            <tbody>{''.join(rows_used)}</tbody>
        </table>
    </div>
    """ if used_list else ""
    html = f"""<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>我的電子商品券</title>
    <style>
        body {{ font-family: "Microsoft JhengHei", sans-serif; max-width: 900px; margin: 20px auto; padding: 20px; background: #1a1a2e; color: #eee; }}
        h1 {{ color: #00d9ff; }}
        h2 {{ color: #00d9ff; font-size: 1.1em; margin-top: 24px; margin-bottom: 12px; }}
        h2.section-used {{ color: #888; }}
        .used-block {{ opacity: 0.85; margin-top: 8px; }}
        a {{ color: #00d9ff; }}
        .link-used {{ color: #888; font-size: 0.9em; }}
        .summary {{ color: #888; margin-bottom: 20px; }}
        table {{ width: 100%; border-collapse: collapse; margin-bottom: 20px; }}
        th, td {{ border: 1px solid #0f3460; padding: 12px; text-align: left; }}
        th {{ background: #0f3460; color: #00d9ff; }}
        .btn {{ display: inline-block; padding: 6px 16px; background: #007bff; color: white !important; text-decoration: none; border-radius: 6px; }}
        .btn:hover {{ background: #0056b3; }}
        .used-row {{ background: #1a2a1a !important; }}
    </style>
</head>
<body>
    <h1>🎟️ 我的電子商品券</h1>
    <p class="summary">最後更新：{datetime.now().strftime('%Y-%m-%d %H:%M')}｜可兌換 {len(available)} 張、已兌換 {len(used_list)} 張</p>
    {section_available}
    {section_used}
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
        subprocess.run(["git", "add", "."],  # 上傳所有檔案（含 TXT、HTML、快取等） 
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
    # 預設為執行後自動上傳 GitHub，可以直接從 IDE 點擊執行 (PY) 
    do_upload = True

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
                expiry, status = cache.get(p["link"].strip(), ("", ""))
                p["expiry"] = expiry
                p["used"] = status
    else:
        force_refresh = "--refresh" in sys.argv
        if force_refresh:
            print("正在重新爬取所有兌換券（含已使用狀態）...")
        else:
            print("正在爬取兌換期間至（使用多執行緒加速，之後會用快取）...")
        # 傳入 verbose=True 以便看到進度
        entries = enrich_prizes_with_expiry(entries, verbose=True, force_refresh=force_refresh)

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
