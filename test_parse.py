
import re
from pathlib import Path

def parse_telegram_section(content):
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

content = Path(r"c:\Users\myhome\Desktop\github\Yahoo序號連結查詢結果_20260319.txt").read_text(encoding="utf-8")
date, prizes = parse_telegram_section(content)
print(f"Found {len(prizes)} prizes in 20260319.txt")
for p in prizes[:3]:
    print(p)
