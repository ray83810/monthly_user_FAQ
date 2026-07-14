# 每月案件問題整理工具 (Monthly User FAQ Web Tool)

一個自動化的客服報表處理與爬蟲整理網頁工具。本工具能夠自動清洗每月的原始客服對話紀錄 Excel，自動排除「未提供服務」的案件，並依照電信商/合約種類將案件分類至 **SOLUTO** 與 **EWS相關** 兩個不同工作表（Worksheet），同時利用優化後的精準爬蟲，為 SOLUTO 案件搜尋高相關性的線上疑難排解補充連結。

---

## 🌐 雲端平台部署使用 (免安裝、免啟動，直接線上用)

本專案已完成雲端部署配置（包含 `requirements.txt`、`Procfile`、`vercel.json` 以及雲端 `/tmp` 暫存目錄相容支援）。您可以直接將此 GitHub 專案連結部署到雲端平台取得專屬網址。

### 1. 部署到 Render 平台 (強力推薦 ⭐⭐⭐)
Render 提供免費的容器服務，非常適合執行此類包含背景執行緒與爬蟲任務的 Flask 程式（不會有 serverless 被凍結的限制）：
1. 註冊並登入 [Render 官網](https://render.com/)（可直接使用 GitHub 帳號登入）。
2. 在儀表板點選 **New** -> **Web Service**。
3. 連結您的 GitHub 專案 `ray83810/monthly_user_FAQ`。
4. Render 會自動讀取專案中的 `Procfile` 配置，您只需選擇 **Free** 免費方案。
5. 點選 **Deploy Web Service**，部署完成後即可取得一個專屬的公網網址（例如 `https://monthly-user-faq.onrender.com`）供所有人使用！

### 2. 部署到 Vercel 平台 (備用選項)
本專案亦相容 Vercel 的無伺服器架構（Serverless）：
1. 註冊並登入 [Vercel 官網](https://vercel.com/)。
2. 點選 **Add New** -> **Project**，匯入您的 GitHub 專案 `ray83810/monthly_user_FAQ`。
3. Vercel 會自動辨識 `vercel.json` 並完成部署。
4. *注意：由於 Vercel 免費 Hobby 方案有 10 秒執行時間上限，且會在發送回應後凍結背景執行緒，若報表筆數較多，推薦優先使用 Render 平台進行部署以確保大批量查詢不中斷。*

---

## 🛠️ 本地執行步驟 (如果您想在自己電腦跑)

### 1. 安裝 Python 環境
請確保您的電腦已安裝 Python 3（建議版本 3.8 以上）。

### 2. 下載專案並安裝相依套件
在終端機（Terminal）或命令提示字元（CMD）中切換至本專案目錄，並執行以下指令安裝需要的套件：

```bash
pip install flask pandas openpyxl bs4 requests gunicorn
```

### 3. 啟動 Flask 後台伺服器
執行以下指令來啟動本地網頁伺服器：

```bash
python app.py
```
*(Windows 系統亦可使用 `py -X utf8 app.py` 啟動以確保 UTF-8 編碼正常)*

### 4. 開啟瀏覽器使用
伺服器啟動後，請在瀏覽器輸入以下網址即可開始使用：
👉 **[http://127.0.0.1:5000](http://127.0.0.1:5000)**

---

## 🌟 核心功能特點
* **多工作表劃分**：
  * **`SOLUTO` 分頁**：存放一般的 SOLUTO 案件，並自動查詢補充連結。
  * **`EWS相關` 分頁**：存放 CHT_HOME 或 SENAO 電信商的非 SOLUTO 案件（此分頁依規定不進行網頁搜尋，以加速產出速度）。
* **極高相關性搜尋**：
  * 只針對客戶提及的**「用戶問題」**與**「問題說明」**去檢索，完全剔除手機或平板型號的商品官網干擾。
  * **品牌保護過濾**：設定主流品牌（如中華電信、神腦、LINE、WeChat、DeepSeek、PDF）字眼校驗，防止搜尋結果張冠李戴。
* **裝置欄位清洗**：自動排除 `Premium K-CM9` 等合約資費代碼或 `N/A` 等預設無效值，使欄位保持乾淨。
* **樣式完美套用**：輸出文件直接套用易讀的背景配色、指定中文字體（Noto Sans TC Medium）、設定適當欄寬並自動開啟格線。
