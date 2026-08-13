# Dify Word 上傳與自動轉檔處理服務 (Dify Upload Word Service)

> ⚠️ **重點提醒：**
> 本服務為專門搭配 **Dify 知識庫** 設計的中繼處理 API。主要解決 Word 檔上傳至 Dify 時的排版跑版、縮排消失、動態編號遺失等問題，並自動完成 PDF 轉檔與預覽連結注入。
> 
> 

---

## 一、 Dify 知識庫前置設定

在上傳文件與呼叫本服務前，**必須先在 Dify 平台完成以下前置設定**：

1. **建立空白知識庫**：
* 在 Dify 的「知識庫」頁面建立一個新知識庫。


2. **新增自訂元數據（Metadata）**：
* 進入該知識庫的 **「設定」->「元數據管理」**。
* 新增以下兩個欄位（名稱必須完全一致）：
* `file_uuid` (字串型態)：紀錄檔案的唯一識別碼。


* `pdf_url` (字串型態)：紀錄轉檔後的 PDF 檢視連結。






3. **取得 API Key 與 Dataset ID**：
* 點選知識庫側邊欄的 **「API 整合」**，建立並複製 API Key（格式如 `dataset-xxx`）。


* 知識庫網址列中的 UUID 即為該知識庫的 `dataset_id`。



---

## 二、 API 服務說明 (`/upload-to-dify`)

服務啟動成功後，即可呼叫核心 API 端點，亦可透過 Swagger UI 介面進行線上測試與檢視完整的 API 規格。

* **Swagger API 文件網址**：`http://YOUR_SERVER_IP:8000/docs`
* **核心 API 端點**：`POST /upload-to-dify`
* **請求格式**：`multipart/form-data`

* **請求參數**：
* `file` (File, 必填)：要處理並同步至 Dify 的 Word 檔案（格式僅支援 `.docx`）。


* `dataset_id` (Form Data, 必填)：目標 Dify 知識庫的 ID。





---

## 三、 環境變數設定 (`.env`)

在專案根目錄建立 `.env` 檔案，並填入以下資訊（請將 `YOUR_SERVER_IP` 替換為實際的測試機 IP）：

```env
# ==========================================
# 網路與服務位址設定
# ==========================================

# 本服務公開存取網址 (用於產生寫入 Dify Chunk 內文的 PDF 預覽連結)
# 警告：請替換為你的伺服器實體 IP，禁止設定為 localhost 或 127.0.0.1
SERVER_BASE_URL=http://YOUR_SERVER_IP:8000

# Dify API 進入點位址 (根據實際 Dify 部署位址與 Port 調整)
# 若 Dify 與本服務位於不同容器或伺服器，請填寫 Dify 伺服器 IP
DIFY_API_BASE=http://YOUR_SERVER_IP/v1


# ==========================================
# Dify 認證與授權
# ==========================================

# Dify 知識庫專用 API Key (前綴通常為 dataset-)
DIFY_API_KEY=dataset-YOUR_DIFY_API_KEY_HERE


# ==========================================
# 系統執行環境設定
# ==========================================

# Windows 原生開發環境之 LibreOffice 執行檔絕對路徑
# 備註：部署至 Docker/Linux 環境時，系統會自動切換為原生指令，忽略此項設定
LIBREOFFICE_WIN_PATH=C:\Program Files\LibreOffice\program\soffice.exe

```

---

## 四、 Docker 部署指南

### 1. 複製專案與設定環境

```bash
# 1. 下載專案原始碼
git clone <YOUR_GITHUB_REPOSITORY_URL>
cd <YOUR_REPOSITORY_DIR>

# 2. 建立 .env 設定檔並填入正確參數
nano .env

# 3. 建立實體持久化資料夾 (選填，Docker 啟動時亦會自動建立)
mkdir -p storage/docx_original storage/docx_modified storage/pdf

```

### 2. 啟動 Docker 容器

```bash
# 建置並於背景啟動容器
docker-compose up -d --build

# 查看容器啟動狀態
docker ps

# 查看應用程式 Log
docker logs -f dify-upload-word-service

```

---

## 五、 `app.py` 程式處理流程說明

當發送 POST 請求至 `/upload-to-dify` 端點時，`app.py` 會依序執行以下自動化流程：

1. **檔案格式檢查與初始化**：
   * 檢查檔案副檔名是否為 `.docx`。
   * 使用 `uuid.uuid4()` 生成全域唯一識別碼 (`file_uuid`) 作為後續本地檔案儲存檔名。
   * 將上傳的檔案暫存至 `storage/docx/{file_uuid}.docx`。

2. **Word 文件內容修正（核心雙步驟）**：
   * `fix_list_indentation_for_libreoffice()`：修正 Word XML 的 `w:leftChars` 屬性，將 `numbering.xml` 真正定義的縮排數值寫死至段落自身的 `w:ind` 上，防止 LibreOffice 轉檔時縮排全部貼齊左邊界。
   * `normalize_and_solidify_list_numbering()`：
     * 依據段落順序與清單識別碼，維護 `(num_id, ilvl)` 複合鍵計數器（以 `num_id` 隔離不同清單，避免同層級的不同獨立清單相互干擾造成編號偏移，如 `a.` 誤變成 `b.`）。
     * 解析 `numbering.xml` 中的樣式樣板（`w:numFmt` 與 `w:lvlText`），將動態算出的階層編號轉為靜態文字寫死在段落文字最前端。
     * 移除段落中的動態編號標記 `<w:numPr>`，防止 Word 或 LibreOffice 重複繪製編號。
   * **儲存檔案**：修正後的內容會直接覆蓋存回 `storage/docx/{file_uuid}.docx`（本服務不另外保存未處理的原始檔）。

3. **PDF 轉檔與靜態掛載**：
   * 呼叫系統內建 LibreOffice Headless 模式，將修正後的 Word 轉為 PDF 並存入 `storage/pdf/{file_uuid}.pdf`。
   * FastAPI 以 `/static/pdf/` 路由掛載該目錄供外部存取預覽，確保外部無法存取與下載原始 Word 檔。

4. **Dify 舊檔清理** (`delete_existing_document_and_clean_local`)：
   * 向 Dify 查詢該知識庫中是否已有同名文件。
   * 若存在舊檔，先刪除 Dify 上的該筆文件，並利用 metadata（或內文解析）找出舊 UUID，同時刪除本地對應的舊 `.docx` 與 `.pdf` 實體檔案，確保每一次上傳都是 Clean State。

5. **上傳至 Dify 知識庫 (Hierarchical)**：
   * 將處理後的 Word 檔透過 `create-by-file` 端點上傳至 Dify。
   * 設定為父子層級模式 (`hierarchical_model`)，並沿用文件內使用者手動標註的 `===CHUNK===` 進行父切塊分割。

6. **Metadata 自動回寫與狀態輪詢**：
   * `set_document_metadata()`：動態查詢並快取 `file_uuid` 與 `pdf_url` 的欄位 ID，將資訊合併寫入該 Dify 文件的元數據。
   * `wait_for_dify_indexing()`：定期輪詢 Dify 索引狀態直至 `completed`。若索引失敗或超時，會主動向 Dify 請求刪除該新建文件並清理本地暫存檔。