# AI 浮動視窗登入修正 V1

## 修正內容

1. AI 助理浮動視窗內新增「AI 功能登入」區塊。
2. 可直接在浮動視窗輸入 AI 使用密碼 / 管理密碼。
3. 登入密碼會使用既有 `/api/auth/check` 驗證，並寫入既有 localStorage key：
   - `ntpc_noise_dashboard_management_token_v1`
4. AI API 呼叫會自動帶入 `X-Dashboard-Token`。
5. 未登入或權限不足時，AI 按鈕會停用並提示需先登入。
6. 若 `AI_REQUIRE_ROLE=viewer`，前台可不登入直接使用 AI 查詢。
7. 新增登出功能，可直接清除浮動視窗登入狀態。

## 對應環境變數

AI 功能密碼沿用管理權限設定，不需新增額外密碼變數：

- `DASHBOARD_ADMIN_TOKEN`：系統管理者密碼 / Token
- `DASHBOARD_EDITOR_TOKEN`：資料編輯者密碼 / Token
- `AI_REQUIRE_ROLE`：AI 使用權限，建議維持 `editor`
- `AI_ENABLED`：需為 `true`
- `OPENAI_API_KEY`：使用 OpenAI 時需設定
- `AI_PROVIDER=heuristic`：未串 OpenAI 時可使用內建規則分析

## 驗證方式

1. 開啟平台，點右下角「AI 助理」。
2. 於浮動視窗上方輸入 `DASHBOARD_ADMIN_TOKEN` 或 `DASHBOARD_EDITOR_TOKEN`。
3. 點「登入」。
4. 顯示「已登入：資料編輯者」或「已登入：系統管理者」後，即可執行主管摘要、排場建議、資料健檢、異常清單與自然語言查詢。
