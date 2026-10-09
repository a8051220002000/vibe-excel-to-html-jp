# 🇯🇵 2026 福岡・九州自由行 — 行程隨身手冊

以 `itinerary.xlsx` 為唯一資料源，透過 CLI 轉換成 JSON，驅動一個**手機優先、可離線使用**的靜態旅遊網頁（GitHub Pages 發布）。

```
itinerary.xlsx ──▶ scripts/build_data.py ──▶ dist/itinerary_data.json ──▶ dist/index.html (+ sw.js 離線快取)
```

## 專案結構

| 路徑 | 說明 |
| --- | --- |
| `itinerary.xlsx` | 行程原始資料（行程表、景點資料、機票資訊、交通套票資訊、todo、遊日優惠） |
| `scripts/build_data.py` | Excel → JSON 轉換 CLI（Forward Fill、型別正規化、空白列過濾） |
| `scripts/inspect_sheets.py` | 檢視各工作表名稱、標題列與資料列數 |
| `scripts/verify_*.py` | 各功能的驗收腳本 |
| `dist/index.html` | 手機版網頁（每日行程 / 航班資訊 / 優惠券 / 實用資訊） |
| `dist/itinerary_data.json` | 轉換產出的資料檔（**請勿手動編輯**，以 Excel 為準） |
| `dist/sw.js` | Service Worker，離線快取 index.html 與資料檔 |
| `dist/.nojekyll` | 關閉 GitHub Pages 的 Jekyll 處理 |
| `.github/workflows/pages.yml` | push 到 `main` 時自動把 `dist/` 發布到 GitHub Pages |

## 環境需求

- Python 3.9+
- `openpyxl`：`python3 -m pip install openpyxl`

## 🔁 更新 Excel 後如何重新產生資料

1. 編輯並儲存 `itinerary.xlsx`（維持原工作表名稱與欄位結構）。
2. （可選）確認工作表結構沒有跑掉：

   ```bash
   python3 scripts/inspect_sheets.py
   ```

3. 重新轉換（可重複執行；除 `generated_at` 時間戳外，結果只取決於 Excel 內容）：

   ```bash
   python3 scripts/build_data.py
   # 或自訂路徑
   python3 scripts/build_data.py --input itinerary.xlsx --output dist/itinerary_data.json
   ```

4. 驗證產出：

   ```bash
   python3 -c "import json; d=json.load(open('dist/itinerary_data.json')); assert len(d['schedule']) > 0, 'Schedule empty'; print('Schema Validation PASSED')"
   ```

5. 本機預覽（Service Worker 需透過 http 存取，不能直接雙擊開檔）：

   ```bash
   python3 -m http.server 8000 --directory dist
   # 瀏覽器開啟 http://localhost:8000 ，用 DevTools 切到 375px 手機寬度檢視
   ```

6. 提交並推送，GitHub Actions 會自動重新發布：

   ```bash
   git add itinerary.xlsx dist/itinerary_data.json
   git commit -m "data: update itinerary"
   git push origin main
   ```

> 手機端採 **network-first** 快取策略：有網路時會自動抓到最新的 `itinerary_data.json`，離線時則使用上次快取的版本。
> 只有在 `dist/` 新增或更名需要預先快取的檔案時，才需要修改 `dist/sw.js` 的 `PRECACHE_URLS` 並遞增 `CACHE_VERSION`。

## 🚀 GitHub Pages 發布設定（首次）

1. 推送到 GitHub 後，進入 repo 的 **Settings → Pages**。
2. **Build and deployment → Source** 選擇 **GitHub Actions**。
3. 之後每次 push 到 `main` 且 `dist/` 有變動，`.github/workflows/pages.yml` 就會自動部署；也可在 **Actions** 頁面手動執行 *Deploy dist/ to GitHub Pages*。
4. 部署完成後網址為 `https://<帳號>.github.io/<repo 名稱>/`，用手機開啟後可「加入主畫面」，出國前先開啟一次即可離線使用。

## ✅ 驗收腳本

```bash
python3 scripts/verify_mobile_layout.py   # 375px 版面、Safe Area
python3 scripts/verify_tabs.py            # 底部導覽列 4 個分頁
python3 scripts/verify_timeline.py        # 每日行程時間軸與 Day 過濾
python3 scripts/verify_cards.py           # 航班卡片、優惠券、一鍵複製
python3 scripts/verify_offline.py         # Service Worker 離線快取
```
