# 新北聲音照相平台 v11

本版修正首頁「成案件數」計算邏輯。成案件數改以 `車輛案件資料` 明細為唯一基準，並支援 `2026-06-12`、`2026/6/12`、`115.06.12` 日期格式。

## 驗算結果

- 告發：64 筆
- 通檢：31 筆
- 成案件數：95 筆

詳見 `VALIDATION_V11.md`。

## Zeabur 環境變數

```env
PORT=8080
NODE_ENV=production
TZ=Asia/Taipei
GOOGLE_AUTO_SYNC=true
GOOGLE_SHEET_ID=1EfP7GoI87RRl1AUGegwPhqNvFN9xG-YXm9NoMSqm_O0
GOOGLE_SHEET_URL=https://docs.google.com/spreadsheets/d/1EfP7GoI87RRl1AUGegwPhqNvFN9xG-YXm9NoMSqm_O0/edit?gid=0#gid=0
GOOGLE_SERVICE_ACCOUNT_JSON=你的ServiceAccountJSON或base64
GOOGLE_SHEET_RAW_NAME=執行成效資料
GOOGLE_SHEET_VEHICLE_NAME=車輛案件資料
LINE_WEBHOOK_VERIFY=false
LINE_CHANNEL_ACCESS_TOKEN=你的LINE_Channel_Access_Token
LINE_CHANNEL_SECRET=你的LINE_Channel_Secret
DASHBOARD_ADMIN_TOKEN=你的管理者Token
DASHBOARD_EDITOR_TOKEN=你的編輯者Token
```


## AI 優化版 V1.0 新增功能

本版新增 AI 成效分析助理，可在前台右下角開啟，並提供主管摘要、排場建議、資料健檢、異常清單與自然語言查詢。

### 新增 API

- `GET /api/ai/status`
- `POST /api/ai/query`
- `POST /api/ai/report`
- `POST /api/ai/validate-import`
- `POST /api/ai/recommend-locations`
- `POST /api/ai/anomaly-check`

### AI 環境變數

```env
AI_ENABLED=true
AI_PROVIDER=openai
OPENAI_API_KEY=你的OpenAI_API_Key
AI_MODEL=gpt-4.1-mini
AI_REQUIRE_ROLE=editor
AI_MASK_PLATE_FOR_VIEWER=true
AI_MAX_ROWS=500
AI_REPORT_DEFAULT_STYLE=主管簡報版
AI_LOG_ENABLED=true
```

若尚未設定 OpenAI API Key，請將 `AI_PROVIDER=heuristic`，系統會使用內建規則分析，不會影響 Zeabur 啟動。

### LINE Bot AI 指令

```text
AI 幫我看本週重點
AI 建議下週排場
AI 檢查今天資料有沒有異常
AI 幫我寫主管摘要
```

AI 回覆只作決策輔助；正式告發、通檢與裁罰仍以原始資料及承辦複核為準。
