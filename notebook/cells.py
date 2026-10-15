# %% [1] 設定：要查的公司（每次只要改這格）
# 代號：自己的觀測站代號；未公開發行填 None
# 公告來源：本公司公告填 None；由母公司代公告填母公司代號
# 篩選名稱：代子公司公告時，主旨中可能出現的子公司名稱（多寫幾種）
COMPANIES = [
    # 簡稱,     代號,    公告來源,  篩選名稱
    ("彰銀",     "2801",  None,     []),
    ("第一銀",   "5844",  "2892",   ["第一銀行", "第一商業銀行"]),
    ("合庫證",   None,    "5880",   ["合作金庫證券", "合庫證券"]),
    ("國票證",   "000779", "2889",  ["國票綜合證券", "國票證券"]),
]
NEWS_FROM = (112, 1)        # 重訊查詢起（民國年, 月）：要涵蓋最近一次改選
NEWS_TO   = (115, 10)       # 重訊查詢迄
HOLD_YM   = (115, 9)        # 持股明細月份；查無資料會自動往前一個月
SLEEP     = 3               # 每次請求間隔秒數，請勿調低，以免被觀測站封鎖
OUTPUT    = "董事名單原始資料.xlsx"

# %% [2] 工具函式（不用改）
import html, re, time, json
import requests
import pandas as pd

BASE = "https://mopsov.twse.com.tw/mops/web/"
S = requests.Session()
S.headers["User-Agent"] = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                           "(KHTML, like Gecko) Chrome/128.0 Safari/537.36")
_last = [0.0]

def post(page, data):
    wait = SLEEP - (time.time() - _last[0])
    if wait > 0:
        time.sleep(wait)
    r = S.post(BASE + page, data=data, headers={"Referer": BASE + page.replace("ajax_", "")}, timeout=60)
    _last[0] = time.time()
    r.encoding = "utf-8"
    if "因為安全性考量" in r.text:
        raise RuntimeError("被觀測站封鎖（因為安全性考量）。請等 1～2 分鐘再從中斷的那格重跑。")
    return r.text

def strip_tags(s):
    s = re.sub(r"<br\s*/?>|</(p|div|tr|li)>", "\n", s, flags=re.I)
    s = re.sub(r"<[^>]+>", "", s)
    return html.unescape(s).replace("\xa0", " ").strip()

def rows_of(page):
    for tr in re.findall(r"<tr[^>]*>(.*?)</tr>", page, flags=re.S | re.I):
        cells = [strip_tags(c) for c in re.findall(r"<t[dh][^>]*>(.*?)</t[dh]>", tr, flags=re.S | re.I)]
        if cells:
            yield cells, tr

def to_roc(ad):  # 20250410 -> 114/04/10
    return f"{int(ad[:4]) - 1911:03d}/{ad[4:6]}/{ad[6:8]}"

def months(a, b):
    y, m = a
    while (y, m) <= b:
        yield y, m
        y, m = (y + 1, 1) if m == 12 else (y, m + 1)

def save_debug(name, page):
    with open(name, "w", encoding="utf-8") as f:
        f.write(page)
    print(f"  [注意] 解析不到資料，原始網頁存成 {name}，請檢查觀測站是否改版")

def news_list(co_id):
    """抓某代號在查詢期間的所有重訊清單"""
    out = []
    for y, m in months(NEWS_FROM, NEWS_TO):
        page = post("ajax_t05st01", {
            "encodeURIComponent": 1, "step": 1, "firstin": 1, "off": 1,
            "keyword4": "", "code1": "", "TYPEK2": "", "checkbtn": "",
            "queryName": "co_id", "inpuType": "co_id", "TYPEK": "all",
            "co_id": co_id, "year": y, "month": m, "b_date": "", "e_date": ""})
        n = 0
        for cells, tr in rows_of(page):
            seq = re.search(r"SEQ_NO\.value='(\d+)'", tr)
            day = re.search(r"SPOKE_DATE\.value='(\d+)'", tr)
            tm = re.search(r"SPOKE_TIME\.value='(\d+)'", tr)
            if seq and day:
                out.append({"發言代號": co_id, "發言日期": to_roc(day[1]),
                            "發言時間": tm[1] if tm else "", "序號": seq[1],
                            "主旨": max(cells, key=len), "_ad": day[1]})
                n += 1
        if n == 0 and "查無" not in page and len(page) > 2000:
            save_debug(f"debug_t05st01_{co_id}_{y}{m:02d}.html", page)
        print(f"  {co_id} {y}/{m:02d}：{n} 則")
    return out

def news_detail(co_id, item):
    page = post("ajax_t05st01", {
        "step": 2, "colorchg": 1, "co_id": co_id, "off": 1, "firstin": "true",
        "TYPEK": "all", "SEQ_NO": item["序號"], "SPOKE_TIME": item["發言時間"],
        "SPOKE_DATE": item["_ad"], "COMPANY_ID": co_id,
        "skey": f"{co_id}{item['_ad']}{item['序號']}"})
    page = re.sub(r"<(script|style)[^>]*>.*?</\1>", "", page, flags=re.S | re.I)
    return re.sub(r"\n\s*\n+", "\n", strip_tags(page))

INCLUDE = ["董事", "董事長", "法人代表", "改派", "改選", "補選", "當選", "選任",
           "解任", "辭任", "逝世", "延任", "重新指派", "指派", "監察人", "股東常會重要決議"]
EXCLUDE = ["股利", "財報", "財務報告", "增資", "公司債", "取得", "處分", "背書", "資金貸與",
           "庫藏股", "薪酬委員", "經理人", "總經理", "發言人", "稽核主管", "會計主管", "財務主管", "法說會"]
FORCE = ["董事長", "改派", "法人董事", "獨立董事", "當選", "重新指派"]

def board_related(subject):
    if any(k in subject for k in FORCE):
        return True
    if any(k in subject for k in EXCLUDE):
        return False
    return any(k in subject for k in INCLUDE)

def holdings(co_id, y, m):
    page = post("ajax_stapap1", {
        "encodeURIComponent": 1, "step": 1, "firstin": 1, "off": 1,
        "keyword4": "", "code1": "", "TYPEK2": "", "checkbtn": "",
        "queryName": "co_id", "inpuType": "co_id", "TYPEK": "all", "isnew": "false",
        "co_id": co_id, "year": y, "month": f"{m:02d}"})
    out = [{"職稱": c[0], "姓名": c[1], "其他欄位": " | ".join(c[2:])}
           for c, _ in rows_of(page) if len(c) >= 3 and ("董事" in c[0] or "監察人" in c[0])]
    if not out and "查無" not in page and len(page) > 2000:
        save_debug(f"debug_stapap1_{co_id}_{y}{m:02d}.html", page)
    return out

print("工具函式載入完成")

# %% [3] 抓重訊清單，篩出董事相關公告（會跑一段時間）
news_cache = {}   # 同一個母公司只抓一次
news_rows = []
for name, code, parent, keys in COMPANIES:
    src = parent or code
    if src not in news_cache:
        print(f"抓 {src} 的重訊清單…")
        news_cache[src] = news_list(src)
    for it in news_cache[src]:
        if parent and not any(k in it["主旨"] for k in keys):
            continue
        if board_related(it["主旨"]):
            news_rows.append({"公司": name, **it})
df_news = pd.DataFrame(news_rows)
print(f"\n董事相關重訊共 {len(df_news)} 則")
df_news[["公司", "發言代號", "發言日期", "發言時間", "主旨"]] if len(df_news) else df_news

# %% [4] 抓每則董事相關重訊的全文（名單在全文裡，主旨通常沒有）
for i, row in df_news.iterrows():
    print(f"{row['公司']} {row['發言日期']} {row['主旨'][:40]}")
    df_news.at[i, "全文"] = news_detail(row["發言代號"], row)
print("完成")

# %% [5] 抓董監事持股餘額明細（交叉核對用；也是現任名單的草稿）
hold_rows, hold_month = [], {}
for name, code, parent, keys in COMPANIES:
    if not code:
        print(f"{name}：沒有公開發行代號，無持股明細")
        continue
    y, m = HOLD_YM
    for _ in range(3):   # 當月查無就往前找，最多 3 個月
        rows = holdings(code, y, m)
        if rows:
            break
        y, m = (y - 1, 12) if m == 1 else (y, m - 1)
    hold_month[name] = f"{y}/{m:02d}" if rows else "查無"
    print(f"{name}：{hold_month[name]}，{len(rows)} 人")
    hold_rows += [{"公司": name, "持股明細月份": hold_month[name], **r} for r in rows]
df_hold = pd.DataFrame(hold_rows)
df_hold

# %% [6] 輸出 Excel
# 分頁：持股明細名單（草稿）／晚於持股明細的異動（要手動套用）／董事相關重訊全文／查詢說明
def after_hold(r):
    hm = hold_month.get(r["公司"])
    if not hm or hm == "查無":
        return True   # 沒持股明細：全部都要看
    y, m = map(int, hm.split("/"))
    ny, nm = (y + 1, 1) if m == 12 else (y, m + 1)
    return r["發言日期"] >= f"{ny:03d}/{nm:02d}/01"

df_pending = df_news[df_news.apply(after_hold, axis=1)] if len(df_news) else df_news
info = pd.DataFrame([
    ["推導方法", "以持股明細為現任名單草稿，再依時間順序套用『待套用異動』分頁中的重訊；"
             "沒有持股明細的公司，從最近一次全面改選或重新指派的重訊全文開始推導"],
    ["重訊期間", f"{NEWS_FROM[0]}/{NEWS_FROM[1]:02d}～{NEWS_TO[0]}/{NEWS_TO[1]:02d}"],
    ["持股明細月份", "；".join(f"{k}：{v}" for k, v in hold_month.items())],
    ["資料來源", "公開資訊觀測站 t05st01（重大訊息）、stapap1（董監事持股餘額明細）"],
], columns=["項目", "說明"])

cols = ["公司", "發言代號", "發言日期", "發言時間", "主旨", "全文"]
with pd.ExcelWriter(OUTPUT, engine="openpyxl") as xw:
    df_hold.to_excel(xw, sheet_name="持股明細名單", index=False)
    df_pending.reindex(columns=cols).to_excel(xw, sheet_name="待套用異動", index=False)
    df_news.reindex(columns=cols).to_excel(xw, sheet_name="董事相關重訊", index=False)
    info.to_excel(xw, sheet_name="查詢說明", index=False)
    for ws in xw.book.worksheets:
        ws.freeze_panes = "A2"
        for col in ws.columns:
            width = max(len(str(c.value or "")) for c in col[:50])
            ws.column_dimensions[col[0].column_letter].width = min(max(8, width * 1.6), 60)
print(f"已輸出 {OUTPUT}")
