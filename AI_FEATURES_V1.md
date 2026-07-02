# 新北聲音照相平台 AI 優化版 V1.0

本封包已在既有 Node/Express、Google Sheet、Excel 匯入匯出、LINE Bot 查詢功能上，新增 AI 決策輔助層。

## 已新增功能

### 1. AI 成效分析助理

前台右下角新增「AI 助理」浮動按鈕，可執行：

- 主管摘要
- 排場建議
- 資料健檢
- 異常清單
- 自然語言查詢

若環境變數設定 `AI_PROVIDER=openai` 且有提供 `OPENAI_API_KEY`，系統會使用 OpenAI 產生文字分析；若未設定 API Key，會自動使用內建規則分析，不影響平台啟動。

### 2. AI API

新增後端 API：

| API | 功能 |
|---|---|
| `GET /api/ai/status` | 檢查 AI 功能狀態 |
| `POST /api/ai/query` | 自然語言查詢 |
| `POST /api/ai/report` | 產生主管摘要、月報摘要 |
| `POST /api/ai/validate-import` | 匯入資料健檢，可上傳 Excel |
| `POST /api/ai/recommend-locations` | 行政區與排場建議 |
| `POST /api/ai/anomaly-check` | 異常資料清單 |

### 3. LINE Bot AI 指令

LINE Bot 新增 AI 指令，可輸入：

```text
AI 幫我看本週重點
AI 建議下週排場
AI 檢查今天資料有沒有異常
AI 幫我寫主管摘要
```

系統會依 Google Sheet / 平台即時資料產出摘要，不直接編造數字。

### 4. AI 資料健檢

目前會檢查：

- 場次編號重複但日期、機台、點位不同
- 完成場次但辨識車流為 0
- 有超標但告發/通檢為 0
- 機台編號是否符合 ZB001～ZB010 / OE_ZB001～OE_ZB010
- 行政區與地址疑似不一致
- 車牌缺漏
- 量測值、標準值、超標值不一致
- 案件類型非告發/通檢
- 高頻車牌提醒

### 5. AI 排場建議

系統會依下列條件產生行政區排序：

- 成案件數
- KPI 件/場
- 超標率
- 車流量
- 場次量

輸出包含：

- S/A/B/C 優先級
- 推薦原因
- 風險提醒
- 下期排場建議

## 重要限制

AI 分析僅作為決策輔助；涉及告發、通檢、裁罰或正式行政處分，仍應以 Google Sheet 原始資料及承辦複核為準。
