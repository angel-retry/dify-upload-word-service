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

1. **檔案接收與隔離儲存 (`/storage`)**：
* 產生全域唯一識別碼 (`file_uuid`)。


* 將原始 Word 寫入 `storage/docx_original/`，修改用副本寫入 `storage/docx_modified/`。




2. **Word 縮排與編號修正**：
* `fix_list_indentation_for_libreoffice()`：修復 Word XML 的 `w:leftChars` 屬性，防止 LibreOffice 轉檔時縮排貼齊左邊界。


* `normalize_and_solidify_list_numbering()`：將動態階層編號（如 1.1, 1.1.1）寫死為實體文字，並移除 XML 中的 `w:numPr` 標記。




3. **PDF 轉檔與靜態掛載**：
* 呼叫系統內建 LibreOffice Headless 模式，將修正後的 Word 轉為 PDF 並存入 `storage/pdf/`。


* FastAPI 以 `/static/pdf/` 路由掛載該目錄，外部僅能存取 PDF，保護原始 Word 檔不外洩。




4. **Chunk 標記與 PDF 網址注入**：
* 在 Word 文件頂端注入 `【資料來源 PDF】: [檔名](SERVER_BASE_URL/static/pdf/uuid.pdf)` 與 `===CHUNK===` 自訂切割標記。




5. **Dify 舊檔清理與階層式索引 (Hierarchical)**：
* `delete_existing_document_and_clean_local()`：檢查知識庫，若存在同名文件則先執行刪除並清理本地舊檔，確保每次以 Clean State 建立。


* 呼叫 Dify API，以父子分段模式（Hierarchical Model）建立新文件。




6. **Metadata 自動回寫與輪詢**：
* 動態查詢並快取 `file_uuid` 與 `pdf_url` 的欄位 ID，回寫至該 Dify 文件的元數據。


* 定期輪詢 Dify 索引狀態直至 `completed`。