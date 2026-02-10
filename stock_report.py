#!/usr/bin/env python3
import csv
import datetime as dt
import html
import json
import os
import re
import sys
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET

ROOT = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(ROOT, "data")
REPORTS_DIR = os.path.join(ROOT, "reports")
HOLDINGS_CSV = os.path.join(DATA_DIR, "holdings.csv")

UA = "Mozilla/5.0 (Macintosh; Intel Mac OS X) StockReport/1.0"


def ensure_dirs():
    os.makedirs(DATA_DIR, exist_ok=True)
    os.makedirs(REPORTS_DIR, exist_ok=True)


def read_holdings():
    if not os.path.exists(HOLDINGS_CSV):
        return []
    rows = []
    with open(HOLDINGS_CSV, newline="", encoding="utf-8") as f:
        for r in csv.DictReader(f):
            rows.append({
                "code": r["code"].strip(),
                "lots": float(r["lots"]),
                "avg_cost": float(r["avg_cost"]),
            })
    return rows


def write_holdings(rows):
    ensure_dirs()
    with open(HOLDINGS_CSV, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=["code", "lots", "avg_cost"])
        w.writeheader()
        for r in rows:
            w.writerow({
                "code": r["code"],
                "lots": f"{r['lots']:.4f}".rstrip("0").rstrip("."),
                "avg_cost": f"{r['avg_cost']:.4f}".rstrip("0").rstrip("."),
            })


def add_holding(code, lots, avg_cost):
    rows = read_holdings()
    code = code.strip()
    for r in rows:
        if r["code"] == code:
            total_lots = r["lots"] + lots
            if total_lots <= 0:
                r["lots"] = total_lots
                r["avg_cost"] = avg_cost
            else:
                r["avg_cost"] = (r["lots"] * r["avg_cost"] + lots * avg_cost) / total_lots
                r["lots"] = total_lots
            write_holdings(rows)
            return
    rows.append({"code": code, "lots": lots, "avg_cost": avg_cost})
    write_holdings(rows)


def http_get(url):
    req = urllib.request.Request(url, headers={"User-Agent": UA})
    with urllib.request.urlopen(req, timeout=20) as f:
        return f.read()


def parse_roc_date(s):
    # ROC date like 113/02/01
    m = re.match(r"(\d+)/(\d+)/(\d+)", s)
    if not m:
        return None
    y, mo, d = map(int, m.groups())
    return dt.date(y + 1911, mo, d)


def parse_float(s):
    s = s.strip().replace(",", "")
    if s in {"--", "-", ""}:
        return None
    try:
        return float(s)
    except ValueError:
        return None


def fetch_twse_month(stock_no, target_date):
    date_str = target_date.strftime("%Y%m") + "01"
    url = f"https://www.twse.com.tw/exchangeReport/STOCK_DAY?response=json&date={date_str}&stockNo={stock_no}"
    raw = http_get(url)
    data = json.loads(raw)
    if data.get("stat") != "OK":
        return None
    title = data.get("title", "")
    name = None
    m = re.search(rf"\b{re.escape(stock_no)}\b\s+(.+?)\s+各日成交資訊", title)
    if m:
        name = m.group(1).strip()
    rows = []
    for r in data.get("data", []):
        date = parse_roc_date(r[0])
        close = parse_float(r[6])
        change = parse_float(r[7])
        rows.append({"date": date, "close": close, "change": change})
    return {"name": name, "rows": rows}


def fetch_twse_latest(stock_no, target_date):
    month = fetch_twse_month(stock_no, target_date)
    if not month or not month["rows"]:
        return None
    rows = [r for r in month["rows"] if r["date"] and r["date"] <= target_date]
    if not rows:
        prev_month = (target_date.replace(day=1) - dt.timedelta(days=1)).replace(day=1)
        month = fetch_twse_month(stock_no, prev_month)
        if not month or not month["rows"]:
            return None
        rows = [r for r in month["rows"] if r["date"] and r["date"] <= target_date]
    if not rows:
        return None
    rows.sort(key=lambda x: x["date"])
    latest = rows[-1]
    prev = rows[-2] if len(rows) >= 2 else None
    return {"source": "TWSE", "name": month.get("name"), "latest": latest, "prev": prev}


def fetch_tpex_latest_all():
    # Latest daily close for OTC
    url = "https://www.tpex.org.tw/web/stock/aftertrading/DAILY_CLOSE_quotes/stk_quote_result.php?l=zh-tw&o=data"
    raw = http_get(url)
    text = raw.decode("utf-8", errors="ignore")
    # CSV with header
    rows = []
    reader = csv.reader(text.splitlines())
    header = None
    for r in reader:
        if not r:
            continue
        if header is None:
            header = r
            continue
        rows.append(r)
    if not header:
        return None
    hmap = {name: i for i, name in enumerate(header)}
    return {"header": header, "hmap": hmap, "rows": rows}


def get_tpex_row(data, code):
    if not data:
        return None
    h = data["hmap"]
    for r in data["rows"]:
        if r[h.get("代號", -1)] == code:
            return r
    return None


def parse_tpex_price(row, hmap):
    close = parse_float(row[hmap.get("收盤", -1)])
    change = parse_float(row[hmap.get("漲跌", -1)])
    name = row[hmap.get("名稱", -1)] if hmap.get("名稱", -1) >= 0 else None
    return close, change, name


def fetch_google_news(code, name=None, limit=5):
    q = code
    if name:
        q = f"{code} {name}"
    params = {
        "q": q,
        "hl": "zh-TW",
        "gl": "TW",
        "ceid": "TW:zh-Hant",
    }
    url = "https://news.google.com/rss/search?" + urllib.parse.urlencode(params)
    raw = http_get(url)
    root = ET.fromstring(raw)
    items = []
    for item in root.findall("./channel/item"):
        title = item.findtext("title") or ""
        link = item.findtext("link") or ""
        pub = item.findtext("pubDate") or ""
        source = item.findtext("source") or ""
        items.append({"title": title, "link": link, "pub": pub, "source": source})
        if len(items) >= limit:
            break
    return items


def fmt_money(x):
    if x is None:
        return "-"
    return f"{x:,.2f}"


def fmt_pct(x):
    if x is None:
        return "-"
    return f"{x:.2f}%"


def generate_report():
    ensure_dirs()
    holdings = read_holdings()
    if not holdings:
        print("No holdings found. Use: python3 stock_report.py add <code> <lots> <avg_cost>")
        return 1

    target_date = dt.date.today() - dt.timedelta(days=1)

    tpex_all = fetch_tpex_latest_all()

    items = []
    totals = {
        "market_value": 0.0,
        "cost": 0.0,
        "pnl": 0.0,
        "day_pnl": 0.0,
    }

    for h in holdings:
        code = h["code"]
        lots = h["lots"]
        avg_cost = h["avg_cost"]
        shares = lots * 1000

        twse = fetch_twse_latest(code, target_date)
        source = None
        name = None
        close = None
        change = None
        price_date = None
        if twse and twse.get("latest") and twse["latest"].get("close") is not None:
            source = "TWSE"
            name = twse.get("name")
            close = twse["latest"]["close"]
            change = twse["latest"]["change"]
            price_date = twse["latest"]["date"]
        else:
            row = get_tpex_row(tpex_all, code)
            if row:
                source = "TPEx"
                close, change, name = parse_tpex_price(row, tpex_all["hmap"])

        if close is None:
            items.append({
                "code": code,
                "name": name,
                "source": source or "-",
                "error": "No price data",
            })
            continue

        day_pct = None
        if change is not None and close is not None:
            prev = close - change
            if prev > 0:
                day_pct = (change / prev) * 100

        market_value = close * shares
        cost = avg_cost * shares
        pnl = (close - avg_cost) * shares
        day_pnl = (change or 0.0) * shares

        totals["market_value"] += market_value
        totals["cost"] += cost
        totals["pnl"] += pnl
        totals["day_pnl"] += day_pnl

        news = fetch_google_news(code, name=name)

        items.append({
            "code": code,
            "name": name,
            "source": source,
            "date": price_date,
            "close": close,
            "change": change,
            "day_pct": day_pct,
            "lots": lots,
            "avg_cost": avg_cost,
            "market_value": market_value,
            "cost": cost,
            "pnl": pnl,
            "day_pnl": day_pnl,
            "news": news,
        })

    # Determine report date from max available price date
    dates = [i.get("date") for i in items if i.get("date")]
    report_date = max(dates) if dates else target_date
    generated_at = dt.datetime.now()

    md_lines = []
    md_lines.append(f"# 台股每日投資報告 ({report_date})")
    md_lines.append("")
    md_lines.append(f"產生時間: {generated_at:%Y-%m-%d %H:%M}")
    md_lines.append("")
    md_lines.append("## 總覽")
    md_lines.append("")
    md_lines.append(f"- 市值: {fmt_money(totals['market_value'])}")
    md_lines.append(f"- 成本: {fmt_money(totals['cost'])}")
    md_lines.append(f"- 未實現損益: {fmt_money(totals['pnl'])}")
    md_lines.append(f"- 單日損益: {fmt_money(totals['day_pnl'])}")
    md_lines.append("")
    md_lines.append("## 個股明細")
    md_lines.append("")
    md_lines.append("| 代碼 | 名稱 | 市場 | 收盤 | 漲跌 | 漲跌幅 | 張數 | 均價 | 市值 | 單日損益 | 未實現損益 |")
    md_lines.append("|---|---|---|---|---|---|---|---|---|---|---|")
    for i in items:
        if i.get("error"):
            md_lines.append(f"| {i['code']} | - | - | - | - | - | - | - | - | - | - |")
            continue
        md_lines.append("| {code} | {name} | {source} | {close} | {change} | {day_pct} | {lots} | {avg_cost} | {mv} | {day_pnl} | {pnl} |".format(
            code=i["code"],
            name=i.get("name") or "-",
            source=i.get("source") or "-",
            close=fmt_money(i["close"]),
            change=fmt_money(i["change"]),
            day_pct=fmt_pct(i["day_pct"]),
            lots=i["lots"],
            avg_cost=fmt_money(i["avg_cost"]),
            mv=fmt_money(i["market_value"]),
            day_pnl=fmt_money(i["day_pnl"]),
            pnl=fmt_money(i["pnl"]),
        ))

    md_lines.append("")
    md_lines.append("## 相關新聞")
    md_lines.append("")
    for i in items:
        md_lines.append(f"### {i['code']} {i.get('name') or ''}")
        if not i.get("news"):
            md_lines.append("- 無最新新聞")
            continue
        for n in i["news"]:
            md_lines.append(f"- {n['title']} ({n['source']})\n  {n['pub']}\n  {n['link']}")

    md = "\n".join(md_lines)

    html_rows = []
    for i in items:
        if i.get("error"):
            html_rows.append(f"<tr><td>{html.escape(i['code'])}</td><td>-</td><td>-</td><td colspan=8>無法取得價格</td></tr>")
            continue
        html_rows.append("""
        <tr>
          <td>{code}</td>
          <td>{name}</td>
          <td>{source}</td>
          <td>{close}</td>
          <td>{change}</td>
          <td>{day_pct}</td>
          <td>{lots}</td>
          <td>{avg_cost}</td>
          <td>{mv}</td>
          <td>{day_pnl}</td>
          <td>{pnl}</td>
        </tr>
        """.format(
            code=html.escape(i["code"]),
            name=html.escape(i.get("name") or "-"),
            source=html.escape(i.get("source") or "-"),
            close=fmt_money(i["close"]),
            change=fmt_money(i["change"]),
            day_pct=fmt_pct(i["day_pct"]),
            lots=i["lots"],
            avg_cost=fmt_money(i["avg_cost"]),
            mv=fmt_money(i["market_value"]),
            day_pnl=fmt_money(i["day_pnl"]),
            pnl=fmt_money(i["pnl"]),
        ))

    news_html = []
    for i in items:
        news_html.append(f"<h3>{html.escape(i['code'])} {html.escape(i.get('name') or '')}</h3>")
        if not i.get("news"):
            news_html.append("<p>無最新新聞</p>")
            continue
        news_html.append("<ul>")
        for n in i["news"]:
            news_html.append(
                f"<li><a href=\"{html.escape(n['link'])}\" target=\"_blank\">{html.escape(n['title'])}</a> "
                f"<span class=\"meta\">({html.escape(n['source'])}) {html.escape(n['pub'])}</span></li>"
            )
        news_html.append("</ul>")

    html_doc = f"""
<!doctype html>
<html lang=\"zh-Hant\">
<head>
  <meta charset=\"utf-8\" />
  <meta name=\"viewport\" content=\"width=device-width, initial-scale=1\" />
  <title>台股每日投資報告</title>
  <style>
    :root {{
      --bg: #f6f1ea;
      --card: #fffaf3;
      --ink: #1d1a17;
      --accent: #2c5e4a;
      --muted: #6c5f55;
    }}
    body {{
      margin: 0;
      font-family: "Noto Serif TC", "Source Han Serif", "PMingLiU", serif;
      color: var(--ink);
      background: radial-gradient(circle at top right, #f9e9d5 0%, #f6f1ea 40%, #efe6db 100%);
    }}
    header {{
      padding: 24px 18px 12px;
      background: linear-gradient(135deg, #f1dfc7 0%, #fffaf3 55%, #f0e1cf 100%);
      border-bottom: 1px solid #e0d4c5;
    }}
    h1 {{
      margin: 0 0 6px;
      font-size: 22px;
      letter-spacing: 1px;
    }}
    .meta {{
      color: var(--muted);
      font-size: 12px;
    }}
    main {{
      padding: 14px 18px 28px;
    }}
    .card {{
      background: var(--card);
      border: 1px solid #e6d9c9;
      border-radius: 12px;
      padding: 14px;
      margin-bottom: 14px;
      box-shadow: 0 6px 20px rgba(0,0,0,0.06);
    }}
    table {{
      width: 100%;
      border-collapse: collapse;
      font-size: 12px;
    }}
    th, td {{
      padding: 8px 6px;
      border-bottom: 1px solid #eadfd0;
      text-align: right;
      white-space: nowrap;
    }}
    th:first-child, td:first-child,
    th:nth-child(2), td:nth-child(2),
    th:nth-child(3), td:nth-child(3) {{
      text-align: left;
    }}
    h2 {{
      margin: 6px 0 10px;
      font-size: 16px;
    }}
    h3 {{
      margin: 12px 0 6px;
      font-size: 14px;
      color: var(--accent);
    }}
    ul {{
      margin: 0;
      padding-left: 18px;
    }}
    li {{
      margin-bottom: 6px;
      line-height: 1.4;
    }}
    a {{
      color: var(--accent);
      text-decoration: none;
    }}
    a:hover {{
      text-decoration: underline;
    }}
  </style>
</head>
<body>
  <header>
    <h1>台股每日投資報告 ({report_date})</h1>
    <div class=\"meta\">產生時間: {generated_at:%Y-%m-%d %H:%M}</div>
  </header>
  <main>
    <section class=\"card\">
      <h2>總覽</h2>
      <div>市值: {fmt_money(totals['market_value'])}</div>
      <div>成本: {fmt_money(totals['cost'])}</div>
      <div>未實現損益: {fmt_money(totals['pnl'])}</div>
      <div>單日損益: {fmt_money(totals['day_pnl'])}</div>
    </section>

    <section class=\"card\">
      <h2>個股明細</h2>
      <table>
        <thead>
          <tr>
            <th>代碼</th>
            <th>名稱</th>
            <th>市場</th>
            <th>收盤</th>
            <th>漲跌</th>
            <th>漲跌幅</th>
            <th>張數</th>
            <th>均價</th>
            <th>市值</th>
            <th>單日損益</th>
            <th>未實現損益</th>
          </tr>
        </thead>
        <tbody>
          {''.join(html_rows)}
        </tbody>
      </table>
    </section>

    <section class=\"card\">
      <h2>相關新聞</h2>
      {''.join(news_html)}
    </section>
  </main>
</body>
</html>
"""

    date_tag = report_date.strftime("%Y-%m-%d")
    md_path = os.path.join(REPORTS_DIR, f"{date_tag}.md")
    html_path = os.path.join(REPORTS_DIR, f"{date_tag}.html")
    latest_md = os.path.join(REPORTS_DIR, "latest.md")
    latest_html = os.path.join(REPORTS_DIR, "latest.html")

    with open(md_path, "w", encoding="utf-8") as f:
        f.write(md)
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(html_doc)
    with open(latest_md, "w", encoding="utf-8") as f:
        f.write(md)
    with open(latest_html, "w", encoding="utf-8") as f:
        f.write(html_doc)

    print(f"Report generated: {html_path}")
    return 0


def usage():
    print("Usage:")
    print("  python3 stock_report.py add <code> <lots> <avg_cost>")
    print("  python3 stock_report.py input")
    print("  python3 stock_report.py list")
    print("  python3 stock_report.py report")


def input_holdings():
    print("Enter holdings, one per line: <code> <lots> <avg_cost>")
    print("Example: 2330 2 580.5")
    print("Press Enter on an empty line to finish.")
    count = 0
    while True:
        try:
            line = input("> ").strip()
        except EOFError:
            break
        if not line:
            break
        # Normalize separators/spaces and strip non-printable chars
        line = line.replace("，", ",").replace("　", " ")
        line = "".join(ch for ch in line if ch.isprintable())
        # Extract numbers to be robust against odd separators
        parts = re.findall(r"[0-9]+(?:\\.[0-9]+)?", line)
        if len(parts) < 3:
            print("Invalid format. Use: <code> <lots> <avg_cost>")
            continue
        code, lots_s, avg_s = parts[0], parts[1], parts[2]
        try:
            lots = float(lots_s)
            avg = float(avg_s)
        except ValueError:
            print("Invalid number for lots/avg_cost.")
            continue
        add_holding(code, lots, avg)
        count += 1
    print(f"Saved {count} holding(s).")
    return 0


def main(argv):
    if len(argv) < 2:
        usage()
        return 1
    cmd = argv[1]
    if cmd == "add" and len(argv) == 5:
        code = argv[2]
        lots = float(argv[3])
        avg = float(argv[4])
        add_holding(code, lots, avg)
        print("OK")
        return 0
    if cmd == "input":
        return input_holdings()
    if cmd == "list":
        for r in read_holdings():
            print(f"{r['code']}\t{r['lots']}\t{r['avg_cost']}")
        return 0
    if cmd == "report":
        return generate_report()
    usage()
    return 1


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
