# scripts/update.py
import pandas as pd
from collections import Counter
from itertools import combinations
from bs4 import BeautifulSoup
from shutil import copyfile
from datetime import datetime
from pathlib import Path

XLS_SRC = Path("data/marksix_latest_200.xlsx")
XLS_DST = Path("marksix_latest_200.xlsx")                 # 給 index.html 下載用
XLS_ENRICHED = Path("marksix_latest_200_updated.xlsx")    # 附加統計
HTML_PATH = Path("index.html")

H2_FREQ20 = "最近20期：各號碼出現次數（由高至低）"
H2_PAIRS50 = "最近50期：同期出現（號碼對）次數排名"
H2_PAIRS200 = "最近200期：同期出現（號碼對）次數排名"

NUMBER_COLS = ["N1","N2","N3","N4","N5","N6"]

def pad2(n:int) -> str:
    return str(n).zfill(2)

def load_data():
    df = pd.read_excel(XLS_SRC, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]
    if "日期" in df.columns:
        df["日期"] = pd.to_datetime(df["日期"], errors="coerce")
        df = df.sort_values("日期")
    return df

def get_row_numbers(row):
    nums = []
    for c in NUMBER_COLS:
        if c in row and pd.notna(row[c]):
            nums.append(int(row[c]))
    return sorted(set(nums))

def freq_count(df_sub):
    nums = []
    for _, row in df_sub.iterrows():
        nums.extend(get_row_numbers(row))
    cnt = Counter(nums)
    return sorted(cnt.items(), key=lambda x: (-x[1], x[0]))

def pair_counts(df_sub):
    pc = Counter()
    for _, row in df_sub.iterrows():
        ns = get_row_numbers(row)
        for a,b in combinations(ns, 2):
            pc[(a,b)] += 1
    return sorted(pc.items(), key=lambda x: (-x[1], x[0][0], x[0][1]))

def render_freq20_rows(freq20):
    rows = []
    maxcnt = freq20[0][1] if freq20 else 0
    for i,(num,cnt) in enumerate(freq20, start=1):
        width = int(round(cnt/maxcnt*100)) if maxcnt else 0
        rows.append(
            f"<tr><td>{i}</td><td>{pad2(num)}</td><td>{cnt}</td>"
            f"<td><div class=\"bar\"><span style=\"width:{width}%;\"></span></div></td></tr>"
        )
    return "\n".join(rows)

def render_pairs_rows(pairs, topn=20):
    rows = []
    for i,((a,b),cnt) in enumerate(pairs[:topn], start=1):
        rows.append(f"<tr><td>{i}</td><td>{pad2(a)}-{pad2(b)}</td><td>{cnt}</td></tr>")
    return "\n".join(rows)

def find_tbody_after_h2(soup, h2_text):
    target_h2 = None
    for h in soup.find_all("h2"):
        if h.get_text(strip=True) == h2_text:
            target_h2 = h
            break
    if not target_h2:
        raise RuntimeError(f"找不到標題：{h2_text}")
    tbl = target_h2.find_next("table")
    if not tbl:
        raise RuntimeError(f"找不到 {h2_text} 後面的 <table>")
    tbody = tbl.find("tbody")
    if not tbody:
        tbody = soup.new_tag("tbody")
        tbl.append(tbody)
    return tbody

def update_updated_line(soup, issue, date_str):
    for p in soup.find_all("p"):
        if p.get_text(strip=True).startswith("數據更新至："):
            p.string = f"數據更新至：{issue}（{date_str}）"
            return
    # 若沒有就加在 <h1> 後面
    h1 = soup.find("h1")
    newp = soup.new_tag("p")
    newp.string = f"數據更新至：{issue}（{date_str}）"
    if h1:
        h1.insert_after(newp)
    else:
        soup.body.insert(0, newp)

def write_enriched_excel(df, freq20, pairs50, pairs200):
    with pd.ExcelWriter(XLS_ENRICHED, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="raw", index=False)
        pd.DataFrame(
            [{"排名":i+1,"號碼":k,"出現次數":v} for i,(k,v) in enumerate(freq20)]
        ).to_excel(writer, sheet_name="freq20", index=False)
        pd.DataFrame(
            [{"排名":i+1,"號碼對":f"{pad2(a)}-{pad2(b)}","次數":c}
             for i,((a,b),c) in enumerate(pairs50[:20])]
        ).to_excel(writer, sheet_name="pairs50", index=False)
        pd.DataFrame(
            [{"排名":i+1,"號碼對":f"{pad2(a)}-{pad2(b)}","次數":c}
             for i,((a,b),c) in enumerate(pairs200[:20])]
        ).to_excel(writer, sheet_name="pairs200", index=False)

def main():
    assert XLS_SRC.exists(), f"找不到 Excel：{XLS_SRC}"
    assert HTML_PATH.exists(), f"找不到 HTML：{HTML_PATH}"

    df = load_data()
    last20 = df.tail(20)
    last50 = df.tail(50)
    last200 = df.tail(200)

    freq20 = freq_count(last20)
    pairs50 = pair_counts(last50)
    pairs200 = pair_counts(last200)

    latest = df.tail(1).iloc[0]
    issue = str(latest.get("期數", "")).strip()
    d = latest.get("日期")
    date_str = d.strftime("%Y-%m-%d") if pd.notna(d) and hasattr(d, "strftime") else str(d)

    soup = BeautifulSoup(HTML_PATH.read_text(encoding="utf-8"), "lxml")

    update_updated_line(soup, issue, date_str)

    tb = find_tbody_after_h2(soup, H2_FREQ20)
    tb.clear()
    tb.append(BeautifulSoup(render_freq20_rows(freq20), "lxml"))

    tb = find_tbody_after_h2(soup, H2_PAIRS50)
    tb.clear()
    tb.append(BeautifulSoup(render_pairs_rows(pairs50, 20), "lxml"))

    tb = find_tbody_after_h2(soup, H2_PAIRS200)
    tb.clear()
    tb.append(BeautifulSoup(render_pairs_rows(pairs200, 20), "lxml"))

    HTML_PATH.write_text(str(soup), encoding="utf-8")

    # 複製來源 Excel 到根目錄（供頁面下載按鈕使用）
    copyfile(XLS_SRC, XLS_DST)

    # 輸出附加統計的 Excel
    write_enriched_excel(df, freq20, pairs50, pairs200)

if __name__ == "__main__":
    main()
