# -*- coding: utf-8 -*-
"""YNM 管理（獨立模組）：解析「UIO保有統計.xlsx」三張表（本期/去年同期/同期差異），
合併進 data.json 的 D['ynmUio']，供 app.html 新增的「YNM管理」頁籤使用。

來源表結構（三張表 UIO台數{本期}/UIO台數{去年同期}/UIO台數{期別}vs{期別}，欄位完全一致，
逐格核對過：每一列的「小計」加總＝「總UIO」列對應欄位、所有16欄加總＝「總UIO」列）：
  - 第1列：車齡分級標籤（D欄無標籤=「≤2月」，E~S欄＝0.5Y/1Y/2Y/.../14Y以上，共16欄）
  - 第5~22列：回廠頻度分級(A~Q,Z 共18列) × 車齡分級(16欄) 的台數矩陣，
    B欄=細分級代碼(如'2-1')、C欄=顧客分級代碼(A~Z)、D~S欄=各車齡台數、
    T=小計、U=小計%、V=累計、W=累計%
  - 第27~31列：綠區／黃區／紅區／總UIO／綠區保留率（同樣 D~S 對應16個車齡欄）
  - 第34~38列：綠區／黃區／紅區／小計／綠區保留率，依 1-3Y/4-6Y/7-13Y/1-13Y/14Y以上 五個車齡區間彙總
    （欄位：D=1-3Y, I=4-6Y, L=7-13Y, O=1-13Y, S=14Y以上, T=小計, U=%, V=累計, W=%）

用法：python build_ynm_uio.py "UIO Excel路徑" [本期標籤] [去年同期標籤]
不帶參數時用下面的 DEFAULT_* 常數。執行前 data.json 必須已存在（先跑過 build.py）。
"""
import sys
import json
from pathlib import Path

import openpyxl

HERE = Path(__file__).resolve().parent
DATA_JSON = HERE / "data.json"

DEFAULT_SRC = Path(r"C:\Users\ES\OneDrive\桌面\AI參照\谷戰報\2608UIO保有統計.xlsx")
DEFAULT_SHEET_CUR = "UIO台數2608"   # 本期（2026/8）
DEFAULT_SHEET_PRE = "UIO台數2508"   # 去年同期（2025/8）
DEFAULT_SHEET_DIFF = "UIO台數26vs25"  # 原表自帶的同期差異，直接採用，不自己重算


def _sheet(wb, name):
    """原表工作表名稱常帶尾端空白，寬鬆比對。"""
    for sn in wb.sheetnames:
        if sn.strip() == name.strip():
            return wb[sn]
    raise KeyError(f"找不到工作表：{name}（實際工作表：{wb.sheetnames}）")


def parse_sheet(ws):
    age_labels = ["≤2月"] + [ws.cell(row=1, column=c).value for c in range(5, 20)]

    freq_rows = []
    for r in range(5, 23):
        sub = ws.cell(row=r, column=2).value
        code = ws.cell(row=r, column=3).value
        if sub is None:
            continue
        vals = [ws.cell(row=r, column=c).value for c in range(4, 20)]
        freq_rows.append(dict(
            label=str(sub), code=code, v=vals,
            subtotal=ws.cell(row=r, column=20).value,
            pct=ws.cell(row=r, column=21).value,
            cum=ws.cell(row=r, column=22).value,
            cumPct=ws.cell(row=r, column=23).value,
        ))

    zone_rows = {}
    for r in range(27, 32):
        label = ws.cell(row=r, column=1).value
        if not label:
            continue
        vals = [ws.cell(row=r, column=c).value for c in range(4, 20)]
        zone_rows[label] = dict(v=vals, subtotal=ws.cell(row=r, column=20).value)

    band_rows = {}
    for r in range(34, 39):
        label = ws.cell(row=r, column=1).value
        if not label:
            continue
        band_rows[label] = dict(
            y1_3=ws.cell(row=r, column=4).value,
            y4_6=ws.cell(row=r, column=9).value,
            y7_13=ws.cell(row=r, column=12).value,
            y1_13=ws.cell(row=r, column=15).value,
            y14p=ws.cell(row=r, column=19).value,
            subtotal=ws.cell(row=r, column=20).value,
            pct=ws.cell(row=r, column=21).value,
            cum=ws.cell(row=r, column=22).value,
            cumPct=ws.cell(row=r, column=23).value,
        )

    return dict(ageLabels=age_labels, freqRows=freq_rows, zoneRows=zone_rows, bandRows=band_rows)


def _validate(period_label, parsed):
    """逐格核對：freqRows 小計加總＝總UIO小計；freqRows 逐欄加總＝總UIO逐欄，
    不符就直接噴錯，不要把沒驗證過的資料餵進網頁。"""
    total_subtotal = sum(r["subtotal"] for r in parsed["freqRows"] if isinstance(r["subtotal"], (int, float)))
    zone_total = parsed["zoneRows"].get("總UIO")
    if zone_total is None:
        raise ValueError(f"[{period_label}] 找不到「總UIO」列，表格結構可能跟預期不符")
    if total_subtotal != zone_total["subtotal"]:
        raise ValueError(f"[{period_label}] 小計核對失敗：freqRows加總={total_subtotal} vs 總UIO={zone_total['subtotal']}")
    col_sums = [sum((r["v"][i] or 0) for r in parsed["freqRows"]) for i in range(16)]
    if col_sums != zone_total["v"]:
        raise ValueError(f"[{period_label}] 逐欄核對失敗：freqRows欄加總={col_sums} vs 總UIO={zone_total['v']}")


def main():
    src = Path(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT_SRC
    sheet_cur = sys.argv[2] if len(sys.argv) > 2 else DEFAULT_SHEET_CUR
    sheet_pre = sys.argv[3] if len(sys.argv) > 3 else DEFAULT_SHEET_PRE
    sheet_diff = sys.argv[4] if len(sys.argv) > 4 else DEFAULT_SHEET_DIFF

    if not src.exists():
        print(f"[錯誤] 找不到來源檔案：{src}")
        sys.exit(1)
    if not DATA_JSON.exists():
        print(f"[錯誤] 找不到 {DATA_JSON}，請先執行 build.py 產生 data.json")
        sys.exit(1)

    wb = openpyxl.load_workbook(src, data_only=True)

    cur = parse_sheet(_sheet(wb, sheet_cur))
    pre = parse_sheet(_sheet(wb, sheet_pre))
    diff = parse_sheet(_sheet(wb, sheet_diff))

    _validate("本期", cur)
    _validate("去年同期", pre)
    # 差異表本身數字可能是負數，「總UIO」核對邏輯一樣成立（純加總關係，正負都適用）
    _validate("同期差異", diff)

    data = json.loads(DATA_JSON.read_text(encoding="utf-8"))
    data["ynmUio"] = dict(
        src=src.name,
        curLabel="2026/8",
        preLabel="2025/8",
        cur=cur,
        pre=pre,
        diff=diff,
    )
    DATA_JSON.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    print(f"YNM UIO 資料已合併進 {DATA_JSON}（來源：{src.name}，本期={sheet_cur}，去年同期={sheet_pre}，差異={sheet_diff}）")
    print("本期總UIO：", cur["zoneRows"]["總UIO"]["subtotal"], "｜去年同期總UIO：", pre["zoneRows"]["總UIO"]["subtotal"])


if __name__ == "__main__":
    main()
