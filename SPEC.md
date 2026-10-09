Spec: Excel-to-JSON Pipeline & Mobile Travel Web (SDD Workflow)
1. 系統目標
以當前目錄下的 itinerary.xlsx 為資料源，打造一個具備可重複執行性（Reproducible）與容錯性（Resilient）的 CLI 轉換管線（Pipeline），並驅動前端手機版旅遊網頁展示。

2. 資料契約與解析原則 (Data Contract & Ingestion Rules)
轉換工具須封裝為獨立 CLI 腳本（例如 scripts/build_data.py），執行規範如下：

執行命令格式： python3 scripts/build_data.py [--input itinerary.xlsx] [--output dist/itinerary_data.json]

解析原則（針對既有試算表特性）：

向下填補 (Forward Fill)： 行程表 的 Day 欄位僅出現在該天首列，後續空白列須自動沿用前值，直到遇到新的 Day 宣告。

型別安全序列化 (Type Serialization)： Excel 中的 datetime、time 物件須自動正規化為 HH:MM 或 YYYY-MM-DD 純字串；空白儲存格與全空列須自動過濾，不得噴出 None 或導致程式拋錯。

多工作表支援： 動態解析 行程表、景點資料、機票資訊、其他資訊、遊日優惠 等分頁，彙整為單一 JSON。

3. 原子化驗收清單 (Atomic Task Checklist)
請 agy 嚴格遵守：一次只執行一個子任務，通過驗收指令後立即 git commit，並在 SPEC.md 將項目打勾 [x]。

Phase 1: 資料層 CLI 建置 (Data Ingestion & Pipeline)
[x] Task 1.1: 試算表結構偵測與 Schema 探勘

目標： 撰寫探索腳本檢查 itinerary.xlsx，輸出所有工作表名稱、標題欄位名稱與資料列數統計。

驗收命令： python3 scripts/inspect_sheets.py（輸出需清晰列出各分頁與標題列）

Git 訊息： chore: add sheet inspection tool

[x] Task 1.2: 實作可重複執行的轉換 CLI (scripts/build_data.py)

目標： 實作完整轉換邏輯，具備 Forward Fill 與型別安全轉換，支援參數化輸入與輸出路徑。

驗收命令： python3 scripts/build_data.py && test -s dist/itinerary_data.json

Git 訊息： feat: implement idempotent excel-to-json build script

[x] Task 1.3: 產出資料結構自我校驗 (Schema Validation)

目標： 撰寫或透過命令驗證產出的 dist/itinerary_data.json 符合 JSON 語法，且包含 schedule（大於 0 天）、flights、coupons 鍵值。

驗收命令： python3 -c "import json; d=json.load(open('dist/itinerary_data.json')); assert len(d['schedule']) > 0, 'Schedule empty'; print('Schema Validation PASSED')"

Git 訊息： test: verify generated json integrity

Phase 2: 手機優先網頁介面 (Mobile-First UI)
[x] Task 2.1: 建立 Mobile-First HTML 基礎骨架

目標： 建立 dist/index.html，配置手機安全邊距（Safe Area Inset）、響應式 Meta 與 CSS 樣式系統。

驗收命令： 在 375px 寬度渲染下，無橫向捲軸溢出。

Git 訊息： feat: init responsive mobile layout skeleton

[x] Task 2.2: 實作底部導覽列 (Bottom Nav Bar) 與分頁切換

目標： 實作 4 個 Tab 切換：📅 每日行程、✈️ 航班資訊、🏷️ 優惠券、ℹ️ 實用資訊。

驗收命令： 點擊任一 Tab 能即時切換視圖，無跳頁與白畫面延遲。

Git 訊息： feat: add persistent bottom navigation tabs

[x] Task 2.3: 實作每日行程時間軸 (Timeline View)

目標： 讀取 schedule 資料，支援頂部 Sticky Day 橫向滑動過濾，卡片清晰展示時間、地點、交通方式與備註展開。

驗收命令： 點擊不同 Day 標籤能正確過濾當日行程。

Git 訊息： feat: render interactive itinerary timeline

[x] Task 2.4: 實作資訊卡片與工具功能

目標： 渲染機票卡片（去/回程）、優惠券網格、地址與電話的一鍵複製到剪貼簿功能。

驗收命令： 點擊複製按鈕能在畫面彈出 Toast 提示「已複製」。

Git 訊息： feat: add flight cards and copy utilities

Phase 3: 離線支援與靜態發布 (Offline & Deployment)
[x] Task 3.1: 註冊 Service Worker 達成 PWA 離線快取

目標： 建立 dist/sw.js 快取 index.html 與 itinerary_data.json，達成完全離線可用。

驗收命令： 模擬離線狀態重新載入網頁仍可正常瀏覽資料。

Git 訊息： feat: add service worker for offline accessibility

[x] Task 3.2: 建立 GitHub Pages 發布設定與說明

目標： 在 dist/ 放入 .nojekyll，並在根目錄建立 README.md 說明後續更新 Excel 時如何重新執行腳本轉換。

驗收命令： test -f dist/.nojekyll && test -f README.md

Git 訊息： docs: add build guide and github pages config
