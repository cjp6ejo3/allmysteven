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
from pathlib import Path
from datetime import datetime

# 腳本所在目錄 = github 資料夾
BASE_DIR = Path(__file__).resolve().parent
TXT_PATTERN = str(BASE_DIR / "Yahoo序號連結查詢結果_*.txt")
OUTPUT_HTML = BASE_DIR / "Telegram獎品網址整理.html"
OUTPUT_TXT = BASE_DIR / "Telegram獎品網址清單.txt"


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
    rows = []
    total_count = 0
    seen_urls = set()  # 去重（同一網址可能在不同日期出現）

    for rec in entries:
        send_date = rec["send_date"]
        prizes = rec["prizes"]
        if not prizes:
            continue
        for p in prizes:
            if p["link"] in seen_urls:
                continue
            seen_urls.add(p["link"])
            total_count += 1
            rows.append(
                f"""
                <tr>
                    <td>{total_count}</td>
                    <td>{p['title']}</td>
                    <td>{send_date}</td>
                    <td>{p['time']}</td>
                    <td>Profile {p['profile']}</td>
                    <td><a href="{p['link']}" target="_blank" rel="noopener">開啟</a></td>
                </tr>
                """
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
    <p class="summary">最後更新：{datetime.now().strftime('%Y-%m-%d %H:%M')}｜共 {len(entries)} 個查詢日、{total_count} 筆獎項網址</p>
    <div class="table-wrap">
        <table>
            <thead><tr><th>#</th><th>標題</th><th>發送日期</th><th>時間</th><th>Profile</th><th>連結</th></tr></thead>
            <tbody>
                {''.join(rows)}
            </tbody>
        </table>
    </div>
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
        subprocess.run(["git", "add", "Telegram獎品網址整理.html", "Telegram獎品網址清單.txt"], 
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

    html = build_html(entries)
    OUTPUT_HTML.write_text(html, encoding="utf-8")
    print(f"HTML 已寫入: {OUTPUT_HTML}")

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
