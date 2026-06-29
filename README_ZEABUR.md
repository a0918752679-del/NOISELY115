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
