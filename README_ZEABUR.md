# 新北市聲音照相平台 v10：成案件數驗算修正版


## v10 修正重點

- 修正首頁「成案件數」異常放大問題。
- 成案件數、告發件數、通檢件數改以 `車輛案件資料` 明細筆數為準。
- 月份圖表、行政區排行、點位 KPI、進度明細同步改用車輛明細驗算。
- LINE Bot 摘要與行政區排行同步改用車輛明細統計，避免 `執行成效資料` 彙總欄位重複累加。
- 保留 Google Sheet 一鍵跳轉與即時連動功能。

驗算基準：目前連動資料 `車輛案件資料` 為 80 筆，其中告發 55 筆、通檢 25 筆，因此成案件數應為 80。

## 版本重點

- 前台資料、後台匯出、LINE Bot 查詢會即時讀取 Google Sheet。
- 保留 `/admin` 後台管理模式。
- 支援一鍵匯出 Excel 總表。
- LINE Bot 支援月份、日期、行政區、機號、車牌等查詢，並回覆 KPI 與佐證資料。
- 新增 Google Sheet 連線測試 API：`/api/google-sheet/test`。

## Google Sheet 必要架構

同一份 Google Sheet 內需有兩個工作表：

1. `執行成效資料`
2. `車輛案件資料`

### 執行成效資料欄位

場次編號、機台編號、案件來源、執行時段、日期、月份、行政區、點位地址、辨識車流、超標數、告發件數、通知到檢件數、告發金額、是否完成

### 車輛案件資料欄位

案件類型、車牌、車種、日期、量測時間、行政區、點位地址、道路、量測值、標準值、超標值、金額、案件編號、官方註記、來源備註

## Zeabur 環境變數

請將 `ZEABUR_ENV_ONCE.txt` 內容貼到 Zeabur 環境變數，並改成實際值。

必要參數：

```env
GOOGLE_AUTO_SYNC=true
GOOGLE_SHEET_ID=你的GoogleSheetID
GOOGLE_SERVICE_ACCOUNT_JSON=你的ServiceAccountJSON或base64
GOOGLE_SHEET_RAW_NAME=執行成效資料
GOOGLE_SHEET_VEHICLE_NAME=車輛案件資料
```

## Service Account 權限

請將 Service Account email 加到 Google Sheet 共用名單，權限設為「編輯者」。

## 啟用後測試

健康檢查：

```text
https://你的Zeabur網域/healthz
```

Google Sheet 狀態：

```text
https://你的Zeabur網域/api/health
```

Google Sheet 讀取測試，需要帶管理或編輯 Token：

```bash
curl -H "X-Dashboard-Token: 你的管理或編輯Token" https://你的Zeabur網域/api/google-sheet/test
```

手動重新讀取 Google Sheet：

```bash
curl -X POST -H "X-Dashboard-Token: 你的管理或編輯Token" https://你的Zeabur網域/api/google-sheet/refresh
```

## 使用入口

前台：

```text
https://你的Zeabur網域/
```

後台：

```text
https://你的Zeabur網域/admin
```

LINE Webhook：

```text
https://你的Zeabur網域/line-webhook
```

## LINE Bot 測試指令

```text
執行進度
2月份執行成效
淡水區執行成效
淡水區 6月 告發率
機號 OE_ZB004
車牌 ABC-1234
查詢說明
```



## v9 管理頁新增
- 後台 `/admin` 的匯入中心已新增「開啟連動 Google Sheet」按鈕，可一鍵跳轉：
  https://docs.google.com/spreadsheets/d/1EfP7GoI87RRl1AUGegwPhqNvFN9xG-YXm9NoMSqm_O0/edit?gid=0#gid=0
- 新增「測試 Sheet 連線」按鈕，會呼叫 `/api/google-sheet/test`，確認執行成效與車輛案件資料筆數。
- `ZEABUR_ENV_ONCE.txt` 已預填 GOOGLE_SHEET_ID 與 GOOGLE_SHEET_URL。正式部署仍需填入 `GOOGLE_SERVICE_ACCOUNT_JSON`。
