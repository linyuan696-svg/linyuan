#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把本機的內文推上 HackMD。只送內容、不碰標籤、不改對照表、可中斷續跑。

用法
----
  推送內文.bat        有「推送清單.txt」就只推清單裡的，沒有就整表全推

清單檔格式：一行一個檔案路徑，跟對照表「檔案路徑」欄一字不差，例如
  ORIGINAL/Aegir Marinus/Aegir/01_Aegir 0.0.0.md

線上長度跟本機一樣的會自動跳過，所以整表全推也不會重複寫。
跑之前請確認要推的篇沒有在瀏覽器開著——編輯器會把舊狀態同步回去。

修正紀錄
2026-10-03  自 restore.py 分出；加清單過濾，進度檔獨立。
2026-10-04  併回 restore.py：清單檔不存在就整表全推，兩支合一。
"""
import os, sys, time, requests, openpyxl

TOKEN = os.environ.get("HACKMD_API_TOKEN")
if not TOKEN: print("找不到 HACKMD_API_TOKEN"); sys.exit(1)
API = "https://api.hackmd.io/v1"
H = {"Authorization": f"Bearer {TOKEN}", "Content-Type": "application/json"}
HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
SHEET = os.path.join(HERE, "HackMD同步對照表.xlsx")
DONE_FILE = os.path.join(HERE, "推送進度.txt")
LIST_FILE = os.path.join(HERE, "推送清單.txt")
FAIL_FILE = os.path.join(HERE, "推送失敗.txt")

PACE = 7.5          # 每篇最少花這麼多秒（避開 100 次／5 分鐘的限速）
SETTLE = 3.0        # 送出後等多久才讀回來驗


def api(method, path, **kw):
    """遇到限速或網路瞬斷都自動等待重試。"""
    last = None
    for attempt in range(1, 9):
        try:
            r = requests.request(method, API + path, headers=H, timeout=30, **kw)
        except requests.exceptions.RequestException as e:
            wait = min(15 * attempt, 90)
            print(f"      連不上（{type(e).__name__}），等 {wait} 秒後重試（{attempt}/8）…")
            time.sleep(wait)
            last = e
            continue
        if r.status_code != 429:
            return r
        wait = float(r.headers.get("Retry-After", 30))
        print(f"      被限速，等 {wait:.0f} 秒…")
        time.sleep(wait)
    raise RuntimeError(f"連續失敗，最後一次：{last}")


def online_len(nid):
    r = api("GET", f"/notes/{nid}")
    if not r.ok:
        return None
    return len(r.json().get("content") or "")


# ── 收集要處理的列 ────────────────────────────────────────────
rows = []
wb = openpyxl.load_workbook(SHEET)
for book in wb.sheetnames:
    if book == "說明":
        continue
    for i, r in enumerate(wb[book].iter_rows(min_row=2, values_only=True), 2):
        if r[3] and r[6] and r[4]:
            rows.append({"書": book, "列": i, "標題": str(r[3]),
                         "路徑": str(r[4]), "網址": str(r[6])})
wb.close()

if os.path.exists(LIST_FILE):
    want_paths = {l.strip() for l in open(LIST_FILE, encoding="utf-8") if l.strip()}
    rows = [r for r in rows if r["路徑"] in want_paths]
    print(f"依清單處理：清單 {len(want_paths)} 篇，對照表對得上 {len(rows)} 篇\n")
else:
    print(f"沒有清單檔，整表全推：{len(rows)} 篇\n")

done = set()
if os.path.exists(DONE_FILE):
    done = {l.strip() for l in open(DONE_FILE, encoding="utf-8") if l.strip()}
    print(f"偵測到上次進度：已完成 {len(done)} 篇，這次會跳過它們\n")

print("讀取 HackMD 筆記清單…")
notes = api("GET", "/notes").json()
by_link = {n.get("publishLink"): n for n in notes if n.get("publishLink")}
by_short = {n.get("shortId"): n for n in notes}

todo = [r for r in rows if r["網址"] not in done]
print(f"對照表 {len(rows)} 篇，這次要處理 {len(todo)} 篇")
print(f"預估時間約 {len(todo) * PACE / 60:.0f} 分鐘。中途可以直接關視窗，下次會從斷點接著跑。\n")
print("=" * 70)

ok = skip = fail = 0
fails = []
done_fh = open(DONE_FILE, "a", encoding="utf-8")

for idx, r in enumerate(todo, 1):
    t0 = time.time()
    tag = f"[{idx}/{len(todo)}] {r['書']}·{r['標題'][:22]}"

    path = os.path.join(REPO, r["路徑"].replace("/", os.sep))
    if not os.path.exists(path):
        print(f"  ✗ {tag} → 本機找不到檔案：{r['路徑']}")
        fails.append((r, "本機找不到檔案")); fail += 1
        continue

    note = by_link.get(r["網址"]) or by_short.get(r["網址"].rstrip("/").split("/")[-1])
    if not note:
        print(f"  ✗ {tag} → HackMD 上找不到對應筆記")
        fails.append((r, "找不到對應筆記")); fail += 1
        continue

    want = open(path, encoding="utf-8").read()
    n = len(want)

    try:
        cur = online_len(note["id"])
    except Exception as e:
        print(f"  ✗ {tag} → 連線持續失敗，先跳過：{e}")
        fails.append((r, "連線失敗")); fail += 1
        time.sleep(PACE)
        continue
    if cur == n:
        print(f"  · {tag} → 本來就是好的（{n} 字元），跳過")
        done_fh.write(r["網址"] + "\n"); done_fh.flush()
        skip += 1
        time.sleep(max(0, PACE / 2 - (time.time() - t0)))
        continue

    got = None
    try:
      for attempt in (1, 2, 3):
        resp = api("PATCH", f"/notes/{note['id']}", json={"content": want})
        if resp.status_code not in (200, 202, 204):
            print(f"  ✗ {tag} → 送出失敗 HTTP {resp.status_code}")
            break
        time.sleep(SETTLE * attempt)          # 沒生效就等久一點再試
        got = online_len(note["id"])
        if got == n:
            break
        print(f"      第 {attempt} 次沒生效（線上 {got} / 應為 {n}），重送…")
    except Exception as e:
        print(f"  ✗ {tag} → 連線持續失敗，先跳過：{e}")
        fails.append((r, "連線失敗")); fail += 1
        time.sleep(PACE)
        continue

    if got == n:
        print(f"  ✓ {tag} → {n} 字元")
        done_fh.write(r["網址"] + "\n"); done_fh.flush()
        ok += 1
    else:
        print(f"  ✗ {tag} → 三次都沒生效（線上 {got} / 應為 {n}）")
        fails.append((r, f"線上 {got} / 應為 {n}")); fail += 1

    if idx % 25 == 0:
        print(f"    ── 已處理 {idx}/{len(todo)}　成功 {ok}　跳過 {skip}　失敗 {fail}")

    time.sleep(max(0, PACE - (time.time() - t0)))

done_fh.close()

print("=" * 70)
print(f"  這一輪：還原 {ok} 篇　本來就好 {skip} 篇　失敗 {fail} 篇")
print("=" * 70)
if fails:
    with open(FAIL_FILE, "w", encoding="utf-8") as f:
        for r, why in fails:
            f.write(f"{r['書']}\t{r['列']}\t{r['標題']}\t{why}\t{r['網址']}\n")
            print(f"  {r['書']} 第 {r['列']} 列  {r['標題']}  →  {why}")
    print(f"\n  失敗清單已存成 還原失敗.txt，再跑一次這支會自動只處理它們。")
else:
    print("  沒有失敗的。")
print("\n完成，請把最後這段複製給我。")
