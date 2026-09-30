# 裕信汽車 服務部損益經營戰情暨費用穿透分析系統（內部網頁版）

損益 → 月份 → 廠別 → 會計科目 → 費用類型 → 原始摘要，一路穿透的服務部經營戰情網頁。

- 網頁：`docs/index.html`（GitHub Pages 發佈）
- **資料已加密**：損益與費用明細以共用密碼加密（PBKDF2-SHA256 60 萬次 + AES-256-GCM），repo 裡看不到任何財務數字；輸入密碼後才在瀏覽器解密。
- 密碼不存放在本 repo 任何檔案中。
- 資料來源：`2026損益表.xlsx`（服務部總表、2601–2608 廠別月表、2601-08費用表）

## 資料夾

| 路徑 | 說明 |
|---|---|
| `docs/index.html` | 加密後的網頁（自動產生，不要手動編輯） |
| `src/build.py` | 讀 Excel → 建 Data Model、勾稽 → `src/data.json` + 資料模型 Excel |
| `src/app.html` | 網頁程式（不含資料） |
| `src/build_secure.mjs` | 把程式 + 資料用密碼加密 → `docs/index.html` |

`src/data.json` 與 Excel 已列入 `.gitignore`，不會上傳。

## 如何更新資料

1. 更新 Excel；如路徑不同，修改 `src/build.py` 最上方的 `SRC`。
2. 產生資料：
   ```bash
   cd src
   python build.py
   ```
3. 加密產生網頁（Windows PowerShell）：
   ```powershell
   $env:WEB_PASSWORD = "你的密碼"; node build_secure.mjs ../docs/index.html
   ```
4. 推送：
   ```bash
   git add docs/index.html
   git commit -m "更新損益資料"
   git push
   ```
5. 等 1–2 分鐘 GitHub Pages 重新部署。

## 更換密碼

以新密碼重跑第 3 步並推送即可，舊密碼立即失效。
