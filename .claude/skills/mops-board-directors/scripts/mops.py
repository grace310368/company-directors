#!/usr/bin/env python3
"""公開資訊觀測站（MOPS）查詢工具：代號查詢、重訊清單、重訊全文、董監事持股明細。

只用 Python 標準函式庫。子指令：

  lookup   <關鍵字...>                         查公司代號（上市／公開發行／上櫃）
  news     <代號> --from 112/01 --to 115/10    逐月抓重訊清單 → JSON
  filter   <news.json> [--subsidiary 名稱]     篩出董事相關重訊
  detail   <代號> <發言日期YYY/MM/DD> <序號>    讀重訊全文
  holdings <代號> <民國年> <月>                 董監事持股餘額明細（stapap1）

觀測站有 WAF：每次請求之間固定間隔 --sleep 秒（預設 3）。若回傳
「因為安全性考量」頁面，程式會停止並提示，不會自動狂打重試。
"""
import argparse
import html
import json
import os
import re
import ssl
import sys
import time
import urllib.parse
import urllib.request

LEGACY = "https://mopsov.twse.com.tw/mops/web/"
NEW_API = "https://mops.twse.com.tw/mops/api/"
OPENDATA = {
    "上市": "https://openapi.twse.com.tw/v1/opendata/t187ap03_L",
    "公開發行": "https://openapi.twse.com.tw/v1/opendata/t187ap03_P",
    "上櫃": "https://www.tpex.org.tw/openapi/v1/mopsfin_t187ap03_O",
    "興櫃": "https://www.tpex.org.tw/openapi/v1/mopsfin_t187ap03_R",
}
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/128.0 Safari/537.36")
BLOCKED_MARK = "因為安全性考量"
SLEEP = 3.0
_last = [0.0]


class Blocked(Exception):
    pass


BLOCKED_HELP = ("觀測站 WAF 擋下請求（因為安全性考量）。請停 1～2 分鐘再試，"
                "或改在自己的電腦執行；不要連續重試。")

# 主旨關鍵字：命中任一個即視為董事相關
BOARD_KEYWORDS = [
    "董事", "獨立董事", "董事長", "副董事長", "常務董事", "法人代表",
    "改派", "改選", "補選", "當選", "選任", "解任", "辭任", "逝世", "延任",
    "重新指派", "指派", "審計委員", "監察人",
]
# 命中這些就排除（董事會決議事項，與名單無關）
EXCLUDE_KEYWORDS = [
    "董事會決議", "決議日期", "股利", "財報", "財務報告", "增資", "公司債",
    "取得", "處分", "背書", "資金貸與", "庫藏股", "薪酬委員", "經理人", "總經理",
    "發言人", "稽核主管", "會計主管", "財務主管", "法說會", "重要子公司",
]
# 一定要保留的（就算碰到排除字）
FORCE_KEYWORDS = ["董事長", "改派", "法人董事", "獨立董事", "當選", "重新指派"]


def _ctx():
    ca = "/root/.ccr/ca-bundle.crt"
    return ssl.create_default_context(cafile=ca) if os.path.exists(ca) else None


def http(url, data=None, json_body=None, referer=None, ua=True):
    wait = SLEEP - (time.time() - _last[0])
    if wait > 0:
        time.sleep(wait)
    headers = {}
    if ua:
        headers["User-Agent"] = UA
    if referer:
        headers["Referer"] = referer
    body = None
    if json_body is not None:
        body = json.dumps(json_body).encode()
        headers["Content-Type"] = "application/json"
        headers["Origin"] = "https://mops.twse.com.tw"
    elif data is not None:
        body = urllib.parse.urlencode(data).encode()
        headers["Content-Type"] = "application/x-www-form-urlencoded"
    req = urllib.request.Request(url, data=body, headers=headers)
    try:
        with urllib.request.urlopen(req, timeout=60, context=_ctx()) as r:
            raw = r.read()
    finally:
        _last[0] = time.time()
    text = raw.decode("utf-8", errors="replace")
    if BLOCKED_MARK in text:
        raise Blocked(url)
    return text


# ---------- HTML 工具 ----------

def strip_tags(s):
    s = re.sub(r"<br\s*/?>", "\n", s, flags=re.I)
    s = re.sub(r"<[^>]+>", "", s)
    return html.unescape(s).replace("\xa0", " ").strip()


def table_rows(page):
    """回傳頁面中所有 <tr> 的 (cells, raw_tr)。"""
    out = []
    for tr in re.findall(r"<tr[^>]*>(.*?)</tr>", page, flags=re.S | re.I):
        cells = [strip_tags(c) for c in
                 re.findall(r"<t[dh][^>]*>(.*?)</t[dh]>", tr, flags=re.S | re.I)]
        if cells:
            out.append((cells, tr))
    return out


def roc(s):
    """各種日期寫法 → 民國 YYY/MM/DD。"""
    s = s.strip()
    m = re.match(r"^(\d{4})(\d{2})(\d{2})$", s)
    if m:
        return f"{int(m[1]) - 1911:03d}/{m[2]}/{m[3]}"
    m = re.match(r"^(\d{2,3})/(\d{1,2})/(\d{1,2})$", s)
    if m:
        return f"{int(m[1]):03d}/{int(m[2]):02d}/{int(m[3]):02d}"
    return s


def ym_range(a, b):
    y, m = map(int, a.split("/"))
    y2, m2 = map(int, b.split("/"))
    while (y, m) <= (y2, m2):
        yield y, m
        m += 1
        if m == 13:
            y, m = y + 1, 1


# ---------- 子指令 ----------

def cmd_lookup(args):
    hits = []
    for market, url in OPENDATA.items():
        try:
            try:
                rows = json.loads(http(url, ua=False))
            except Blocked:  # openapi 偶爾也會被 WAF 擋，等一下再試一次
                time.sleep(10)
                rows = json.loads(http(url, ua=False))
        except Exception as e:  # 該來源連不到就略過
            print(f"[略過] {market} 清單抓不到：{e!r}", file=sys.stderr)
            continue
        for r in rows:
            code = r.get("公司代號") or r.get("SecuritiesCompanyCode") or ""
            short = r.get("公司簡稱") or r.get("CompanyAbbreviation") or ""
            name = r.get("公司名稱") or r.get("CompanyName") or ""
            chair = r.get("董事長") or r.get("Chairman") or ""
            for kw in args.keywords:
                if kw in short or kw in name or kw == code:
                    hits.append({"查詢": kw, "代號": code, "簡稱": short, "名稱": name,
                                 "市場別": market, "董事長(基本資料)": chair})
    for kw in args.keywords:
        mine = [h for h in hits if h["查詢"] == kw]
        if not mine:
            print(f"{kw}\t（查無；試試全名或不同寫法，如「彰化商業銀行」。仍查無則可能未公開發行，改查母公司代子公司公告）")
        for h in mine:
            print(f"{kw}\t{h['代號']}\t{h['簡稱']}\t{h['名稱']}\t{h['市場別']}\t董事長:{h['董事長(基本資料)']}")


def _news_month_legacy(co_id, y, m):
    page = http(LEGACY + "ajax_t05st01", data={
        "encodeURIComponent": 1, "step": 1, "firstin": 1, "off": 1,
        "keyword4": "", "code1": "", "TYPEK2": "", "checkbtn": "",
        "queryName": "co_id", "inpuType": "co_id", "TYPEK": "all",
        "co_id": co_id, "year": y, "month": m, "b_date": "", "e_date": "",
    }, referer=LEGACY + "t05st01")
    items = []
    for cells, tr in table_rows(page):
        seq = re.search(r"SEQ_NO\.value='(\d+)'", tr)
        sdate = re.search(r"SPOKE_DATE\.value='(\d+)'", tr)
        stime = re.search(r"SPOKE_TIME\.value='(\d+)'", tr)
        if not (seq and sdate):
            continue
        # 欄位：公司代號 公司名稱 發言日期 發言時間 主旨 [按鈕]
        subject = max(cells, key=len)
        items.append({"代號": co_id, "發言日期": roc(sdate[1]),
                      "發言時間": stime[1] if stime else "",
                      "序號": seq[1], "主旨": subject})
    return items, page


def cmd_news(args):
    all_items = []
    for y, m in ym_range(args.date_from, args.date_to):
        items, page = _news_month_legacy(args.co_id, y, m)
        if not items and "查無" not in page and "無資料" not in page and len(page) > 2000:
            dump = f"debug_t05st01_{args.co_id}_{y}{m:02d}.html"
            open(dump, "w", encoding="utf-8").write(page)
            print(f"[警告] {y}/{m:02d} 解析不到資料但頁面非空，原始 HTML 存於 {dump}，"
                  "請檢查版面是否改版", file=sys.stderr)
        print(f"{args.co_id} {y}/{m:02d}: {len(items)} 則", file=sys.stderr)
        all_items += items
    all_items.sort(key=lambda x: (x["發言日期"], x["發言時間"]))
    _out(all_items, args.output)


def is_board_related(subject):
    if any(k in subject for k in FORCE_KEYWORDS):
        return True
    if any(k in subject for k in EXCLUDE_KEYWORDS):
        return False
    return any(k in subject for k in BOARD_KEYWORDS)


def cmd_filter(args):
    items = json.load(open(args.news_json, encoding="utf-8"))
    out = []
    for it in items:
        s = it["主旨"]
        if args.subsidiary and not any(n in s for n in args.subsidiary):
            continue
        if is_board_related(s):
            out.append(it)
    for it in out:
        print(f"{it['發言日期']} {it['發言時間']:>6} #{it['序號']:<3} {it['主旨']}", file=sys.stderr)
    print(f"共 {len(out)} / {len(items)} 則", file=sys.stderr)
    _out(out, args.output)


def cmd_detail(args):
    y, m, d = roc(args.date).split("/")
    ad = f"{int(y) + 1911}{m}{d}"
    page = http(LEGACY + "ajax_t05st01", data={
        "step": 2, "colorchg": 1, "co_id": args.co_id, "off": 1, "firstin": "true",
        "TYPEK": "all", "SEQ_NO": args.seq, "SPOKE_TIME": args.time or "",
        "SPOKE_DATE": ad, "COMPANY_ID": args.co_id,
        "skey": f"{args.co_id}{ad}{args.seq}",
    }, referer=LEGACY + "t05st01")
    text = strip_tags(re.sub(r"<(script|style)[^>]*>.*?</\1>", "", page, flags=re.S | re.I))
    text = re.sub(r"\n\s*\n+", "\n", text)
    if args.output:
        open(args.output, "w", encoding="utf-8").write(text)
    print(text)


def cmd_holdings(args):
    page = http(LEGACY + "ajax_stapap1", data={
        "encodeURIComponent": 1, "step": 1, "firstin": 1, "off": 1,
        "keyword4": "", "code1": "", "TYPEK2": "", "checkbtn": "",
        "queryName": "co_id", "inpuType": "co_id", "TYPEK": "all", "isnew": "false",
        "co_id": args.co_id, "year": args.year, "month": f"{int(args.month):02d}",
    }, referer=LEGACY + "stapap1")
    rows = []
    for cells, _ in table_rows(page):
        if len(cells) < 3:
            continue
        title = cells[0]
        if not any(k in title for k in ("董事", "監察人")):
            continue
        rows.append({"職稱": title, "姓名": cells[1], "其他欄位": cells[2:]})
    if not rows:
        dump = f"debug_stapap1_{args.co_id}_{args.year}{int(args.month):02d}.html"
        open(dump, "w", encoding="utf-8").write(page)
        print(f"[注意] 查無董事資料（可能當月未申報，請改查前一個月）；原始 HTML：{dump}",
              file=sys.stderr)
    for r in rows:
        print(f"{r['職稱']}\t{r['姓名']}", file=sys.stderr)
    _out(rows, args.output)


def _out(obj, path):
    s = json.dumps(obj, ensure_ascii=False, indent=1)
    if path:
        os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
        open(path, "w", encoding="utf-8").write(s)
        print(f"→ {path}", file=sys.stderr)
    else:
        print(s)


def main():
    global SLEEP
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--sleep", type=float, default=3.0, help="請求間隔秒數（預設 3）")
    sp = p.add_subparsers(dest="cmd", required=True)

    a = sp.add_parser("lookup"); a.add_argument("keywords", nargs="+")
    a = sp.add_parser("news"); a.add_argument("co_id")
    a.add_argument("--from", dest="date_from", required=True, help="民國年/月，例 112/01")
    a.add_argument("--to", dest="date_to", required=True)
    a.add_argument("-o", "--output")
    a = sp.add_parser("filter"); a.add_argument("news_json")
    a.add_argument("--subsidiary", nargs="*", help="只保留主旨含這些名稱的公告（代子公司公告用）")
    a.add_argument("-o", "--output")
    a = sp.add_parser("detail"); a.add_argument("co_id"); a.add_argument("date"); a.add_argument("seq")
    a.add_argument("--time", help="發言時間 HHMMSS（清單裡有）"); a.add_argument("-o", "--output")
    a = sp.add_parser("holdings"); a.add_argument("co_id"); a.add_argument("year"); a.add_argument("month")
    a.add_argument("-o", "--output")

    args = p.parse_args()
    SLEEP = args.sleep
    try:
        {"lookup": cmd_lookup, "news": cmd_news, "filter": cmd_filter,
         "detail": cmd_detail, "holdings": cmd_holdings}[args.cmd](args)
    except Blocked as e:
        sys.exit(f"{BLOCKED_HELP}\n（被擋的網址：{e}）")


if __name__ == "__main__":
    main()
