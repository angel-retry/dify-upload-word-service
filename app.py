import os
import requests
import json
import uuid
import re
import gc
import platform
import subprocess
import asyncio
from pathlib import Path
from fastapi import FastAPI, UploadFile, File, Form
from fastapi.staticfiles import StaticFiles
from dotenv import load_dotenv
import docx
from docx.oxml.ns import qn

load_dotenv()

app = FastAPI()

BASE_DIR = Path(__file__).resolve().parent

# 維持三分立的資料夾結構，保護原始檔不外洩
STORAGE_DIR = BASE_DIR / "storage"
DOCX_ORIGINAL_DIR = STORAGE_DIR / "docx_original"
DOCX_MODIFIED_DIR = STORAGE_DIR / "docx_modified"
PDF_DIR = STORAGE_DIR / "pdf"

DOCX_ORIGINAL_DIR.mkdir(parents=True, exist_ok=True)
DOCX_MODIFIED_DIR.mkdir(parents=True, exist_ok=True)
PDF_DIR.mkdir(parents=True, exist_ok=True)

# 僅掛載 PDF 目錄為靜態路由，確保外部無法存取原始 Word 檔
app.mount("/static/pdf", StaticFiles(directory=str(PDF_DIR)), name="static_pdf")

SERVER_BASE_URL = os.getenv("SERVER_BASE_URL", "http://localhost:8000")
DIFY_API_BASE = os.getenv("DIFY_API_BASE", "http://localhost/v1")
DIFY_API_KEY = os.getenv("DIFY_API_KEY")

if not DIFY_API_KEY:
    raise ValueError("未設定 DIFY_API_KEY，請確認 .env 檔案設定。")

# 新增：metadata 欄位 ID 的記憶體快取，不用再手動把 ID 貼進 .env。
# 第一次用到某個欄位名稱時才會打一次查詢 API，之後同一次程式執行期間直接沿用快取，
# 不會每次上傳都重打請求。若欄位被刪除重建、快取失效，會自動重新查詢。
_METADATA_FIELD_CACHE = {}


def get_metadata_field_id(dataset_id: str, field_name: str, headers: dict, force_refresh: bool = False) -> str:
    """
    動態取得某個 metadata 欄位（例如 file_uuid、pdf_url）在 Dify 裡的欄位 ID。
    優先從記憶體快取拿；快取沒有時才打 API 查詢一次並存進快取。
    快取 key 改用 (dataset_id, field_name)，避免不同知識庫的欄位 ID 互相污染。
    為避免觸發 Dify 1.16.0-rc1 資料庫寫入異常，已移除自動建立欄位的邏輯。
    """
    cache_key = (dataset_id, field_name)
    if not force_refresh and cache_key in _METADATA_FIELD_CACHE:
        return _METADATA_FIELD_CACHE[cache_key]

    print(f"[Metadata查詢] 快取中沒有 '{field_name}'，向 Dify 查詢目前所有欄位...")
    url = f"{DIFY_API_BASE}/datasets/{dataset_id}/metadata"
    try:
        resp = requests.get(url, headers=headers)
        if resp.status_code == 200:
            fields = resp.json().get("doc_metadata", [])
            for f in fields:
                if f.get("name") == field_name:
                    _METADATA_FIELD_CACHE[cache_key] = f.get("id")
                    print(f"[Metadata查詢] 找到欄位 '{field_name}'，ID={f.get('id')}，已存入快取。")
                    return f.get("id")
        else:
            print(f"[Metadata查詢] 查詢欄位列表失敗，狀態碼: {resp.status_code}")
    except Exception as e:
        print(f"[Metadata查詢] 發生例外錯誤: {str(e)}")

    print(f"[Metadata查詢] 欄位 '{field_name}' 不存在，為避免資料庫異常，不執行自動建立。")
    return None

LIBREOFFICE_WIN_PATH = os.getenv(
    "LIBREOFFICE_WIN_PATH",
    r"C:\Program Files\LibreOffice\program\soffice.exe"
)


def safe_unlink(file_path: Path):
    if file_path.exists():
        try:
            file_path.unlink()
            print(f"[檔案刪除] 已成功刪除本地檔案: {file_path}")
        except Exception as e:
            print(f"[檔案刪除] 刪除本地檔案失敗 ({file_path}): {str(e)}")


def set_document_metadata(dataset_id: str, document_id: str, file_uuid: str, pdf_url: str, headers: dict) -> bool:
    """
    把 file_uuid 與 pdf_url 一起寫進該 Dify 文件的 metadata（合併成一次請求）。
    欄位 ID 改為動態查詢＋快取（見 get_metadata_field_id），不需要手動在 .env 設定 ID，
    也不怕欄位被刪除重建後 ID 對不上的問題。

    注意：這支端點官方文件目前未列出（屬於社群驗證可用、但未正式收錄的端點），
    不同 Dify 版本行為可能有落差，建議先用 curl/Postman 對你的環境測試過一次。
    """
    field_ids = {
        "file_uuid": get_metadata_field_id(dataset_id, "file_uuid", headers),
        "pdf_url": get_metadata_field_id(dataset_id, "pdf_url", headers),
    }
    values = {"file_uuid": file_uuid, "pdf_url": pdf_url}

    metadata_list = [
        {"id": field_ids[name], "name": name, "value": values[name]}
        for name in ("file_uuid", "pdf_url")
        if field_ids[name]
    ]

    if not metadata_list:
        print("[Metadata寫入] 找不到任何可用的 metadata 欄位 ID，略過寫入。")
        return False

    def _do_write(list_to_write):
        url = f"{DIFY_API_BASE}/datasets/{dataset_id}/documents/metadata"
        payload = {
            "operation_data": [
                {
                    "document_id": document_id,
                    "metadata_list": list_to_write,
                    "partial_update": True
                }
            ]
        }
        return requests.post(url, headers=headers, json=payload)

    try:
        resp = _do_write(metadata_list)
        if resp.status_code in (200, 202):
            print(f"[Metadata寫入] 成功將 file_uuid={file_uuid}, pdf_url={pdf_url} 寫入文件 {document_id}")
            return True

        # 若失敗，可能是快取裡的欄位 ID 已經過期（欄位被刪除重建過），
        # 強制重新查詢一次欄位 ID 後再重試一次。
        print(f"[Metadata寫入] 第一次嘗試失敗，狀態碼: {resp.status_code}，內容: {resp.text}")
        print("[Metadata寫入] 懷疑快取的欄位 ID 已過期，強制重新查詢後重試一次...")
        field_ids = {
            "file_uuid": get_metadata_field_id(dataset_id, "file_uuid", headers, force_refresh=True),
            "pdf_url": get_metadata_field_id(dataset_id, "pdf_url", headers, force_refresh=True),
        }
        metadata_list = [
            {"id": field_ids[name], "name": name, "value": values[name]}
            for name in ("file_uuid", "pdf_url")
            if field_ids[name]
        ]
        if not metadata_list:
            return False

        resp2 = _do_write(metadata_list)
        if resp2.status_code in (200, 202):
            print(f"[Metadata寫入] 重試成功，已寫入文件 {document_id}")
            return True
        print(f"[Metadata寫入] 重試仍失敗，狀態碼: {resp2.status_code}，內容: {resp2.text}")
        return False
    except Exception as e:
        print(f"[Metadata寫入] 發生例外錯誤: {str(e)}")
        return False


def get_file_uuid_from_metadata(doc: dict) -> str:
    """
    新增：嘗試從「取得文件列表」API 回傳的單筆 doc 物件中，
    解析出 file_uuid metadata 的值。
    Dify 的 doc_metadata 欄位常見格式為 list，例如：
    [{"id": "...", "name": "file_uuid", "value": "xxxx-xxxx"}]
    若你的 Dify 版本格式不同，請依實際回傳結構調整這裡。
    """
    doc_metadata = doc.get("doc_metadata") or doc.get("metadata") or []
    if isinstance(doc_metadata, list):
        for item in doc_metadata:
            if isinstance(item, dict) and item.get("name") == "file_uuid" and item.get("value"):
                return item["value"]
    elif isinstance(doc_metadata, dict):
        if doc_metadata.get("file_uuid"):
            return doc_metadata["file_uuid"]
    return None


# ============================================================
# 修正 w:leftChars 造成的 LibreOffice 縮排跑版問題（完全保留未動）
# ============================================================
def fix_list_indentation_for_libreoffice(doc: docx.Document):
    """
    問題根源：
    段落 pPr 裡常見 <w:ind w:leftChars="0"/>（以字元為單位的縮排覆蓋值）。
    Word 能正確運算這個值疊加 numbering.xml 定義的階層縮排，
    但 LibreOffice 對 leftChars 的相容性較差，常常直接把它當成
    「縮排=0」，導致所有層級（1. / 1.1 / 1.1.1）全部貼齊左邊界，
    階層感消失。

    解法：
    完全不動 w:numPr（自動編號機制維持原樣，不會有編號重複問題），
    只針對每個有編號的段落，去查 numbering.xml 中對應
    numId + ilvl（或 lvlOverride）該有的 left / hanging 縮排值，
    直接寫死覆蓋到段落自己的 w:ind 上，並移除會讓 LibreOffice
    誤判的 leftChars / rightChars 屬性。
    """
    numbering_part = doc.part.numbering_part
    numbering_root = numbering_part.element

    # 1) numId -> abstractNumId 對照表（含 lvlOverride 的獨立 ind 設定）
    num_to_abstract = {}
    num_lvl_overrides = {}  # numId -> {ilvl: (left, hanging)}

    for num_el in numbering_root.findall(qn('w:num')):
        num_id = num_el.get(qn('w:numId'))
        abstract_el = num_el.find(qn('w:abstractNumId'))
        if abstract_el is not None:
            num_to_abstract[num_id] = abstract_el.get(qn('w:val'))

        overrides = {}
        for lvl_override in num_el.findall(qn('w:lvlOverride')):
            ilvl = lvl_override.get(qn('w:ilvl'))
            lvl_el = lvl_override.find(qn('w:lvl'))
            if lvl_el is not None:
                pPr_el = lvl_el.find(qn('w:pPr'))
                ind_el = pPr_el.find(qn('w:ind')) if pPr_el is not None else None
                if ind_el is not None:
                    left = ind_el.get(qn('w:left'))
                    hanging = ind_el.get(qn('w:hanging'))
                    if left is not None:
                        overrides[ilvl] = (left, hanging)
        if overrides:
            num_lvl_overrides[num_id] = overrides

    # 2) abstractNumId -> {ilvl: (left, hanging)} 對照表
    abstract_lvl_ind = {}
    for abstract_el in numbering_root.findall(qn('w:abstractNum')):
        abstract_id = abstract_el.get(qn('w:abstractNumId'))
        lvl_map = {}
        for lvl_el in abstract_el.findall(qn('w:lvl')):
            ilvl = lvl_el.get(qn('w:ilvl'))
            pPr_el = lvl_el.find(qn('w:pPr'))
            ind_el = pPr_el.find(qn('w:ind')) if pPr_el is not None else None
            if ind_el is not None:
                left = ind_el.get(qn('w:left'))
                hanging = ind_el.get(qn('w:hanging'))
                if left is not None:
                    lvl_map[ilvl] = (left, hanging)
        abstract_lvl_ind[abstract_id] = lvl_map

    fixed_count = 0

    for p in doc.paragraphs:
        pPr = p._p.pPr
        if pPr is None:
            continue

        numPr = pPr.find(qn('w:numPr'))
        if numPr is None:
            continue  # 沒有自動編號，跳過，不動它的縮排

        ilvl_el = numPr.find(qn('w:ilvl'))
        numId_el = numPr.find(qn('w:numId'))
        if numId_el is None:
            continue

        ilvl = ilvl_el.get(qn('w:val')) if ilvl_el is not None else '0'
        num_id = numId_el.get(qn('w:val'))

        # 優先找 lvlOverride，其次找 abstractNum 預設層級縮排
        left = hanging = None
        if num_id in num_lvl_overrides and ilvl in num_lvl_overrides[num_id]:
            left, hanging = num_lvl_overrides[num_id][ilvl]
        else:
            abstract_id = num_to_abstract.get(num_id)
            if abstract_id is not None:
                lvl_map = abstract_lvl_ind.get(abstract_id, {})
                if ilvl in lvl_map:
                    left, hanging = lvl_map[ilvl]

        if left is None:
            continue  # 查不到對照值，保守跳過，避免誤改

        # 3) 覆蓋段落自己的 w:ind：寫死 left / hanging（twips），
        #    並移除 leftChars / rightChars，避免 LibreOffice 誤判
        ind_el = pPr.find(qn('w:ind'))
        if ind_el is None:
            ind_el = pPr.makeelement(qn('w:ind'), {})
            pPr.append(ind_el)

        ind_el.set(qn('w:left'), left)
        if hanging is not None:
            ind_el.set(qn('w:hanging'), hanging)

        for attr in ('w:leftChars', 'w:rightChars', 'w:hangingChars', 'w:firstLineChars'):
            if ind_el.get(qn(attr)) is not None:
                del ind_el.attrib[qn(attr)]

        fixed_count += 1

    print(f"[縮排修正] 共修正 {fixed_count} 個段落的縮排數值")


def convert_docx_to_pdf(docx_path: Path, output_pdf_path: Path) -> Path:
    print(f"[PDF轉檔] 開始轉換 PDF: {docx_path.name}")
    try:
        if platform.system() == "Windows":
            if not os.path.exists(LIBREOFFICE_WIN_PATH):
                raise FileNotFoundError(
                    f"找不到 LibreOffice 執行檔，請確認已安裝於: {LIBREOFFICE_WIN_PATH}"
                )
            cmd = [
                LIBREOFFICE_WIN_PATH,
                "--headless",
                "--convert-to", "pdf",
                "--outdir", str(output_pdf_path.parent.resolve()),
                str(docx_path.resolve())
            ]
        else:
            cmd = [
                "libreoffice",
                "--headless",
                "--convert-to", "pdf",
                "--outdir", str(output_pdf_path.parent.resolve()),
                str(docx_path.resolve())
            ]

        subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=True)

        generated_pdf = output_pdf_path.parent / f"{docx_path.stem}.pdf"
        if generated_pdf.exists() and generated_pdf != output_pdf_path:
            generated_pdf.replace(output_pdf_path)

        print(f"[PDF轉檔] 轉換成功: {output_pdf_path.name}")
        return output_pdf_path
    except Exception as e:
        print(f"[PDF轉檔] 發生錯誤: {str(e)}")
        raise RuntimeError(f"LibreOffice 轉檔失敗: {str(e)}")


def delete_existing_document_and_clean_local(dataset_id: str, original_filename: str, headers: dict) -> bool:
    """
    查詢 Dify 知識庫中是否存在同名文件。
    若存在，會：
      1. 呼叫 Dify 刪除文件 API，把舊文件從知識庫移除
      2. 解析其舊 UUID 並刪除本地三個資料夾對應的舊實體檔案
    回傳是否有刪除到舊文件（True/False）。

    改為「刪除舊文件＋一律新增」而非「update-by-file」，
    原因：update-by-file 無法變更既有文件的 doc_form（一般 / 父子模式），
    若舊文件是在改用 hierarchical 模式之前建立的，更新時會被 Dify 拒絕
    （錯誤：doc_form is different from the dataset doc_form）。
    刪除後重新建立，可確保每次都是乾淨的 hierarchical 模式文件。

    嚴謹處理：若查詢文件列表失敗、或找到同名文件卻刪除失敗，
    一律直接丟出例外中止整個上傳流程，避免 Dify 裡同時存在
    「刪不掉的舊文件」與「新建立的文件」兩份同名重複文件。
    """
    print(f"[狀態查詢] 檢查 Dify 是否已有同名文件: '{original_filename}'")
    list_url = f"{DIFY_API_BASE}/datasets/{dataset_id}/documents"
    params = {"page": 1, "limit": 100}
    response = requests.get(list_url, headers=headers, params=params)

    if response.status_code != 200:
        raise Exception(
            f"查詢知識庫文件列表失敗，狀態碼: {response.status_code}，內容: {response.text}"
        )

    data = response.json()
    documents = data.get("data", [])

    for doc in documents:
        if doc.get("name") == original_filename:
            doc_id = doc.get("id")
            print(f"[狀態查詢] 找到舊版本文件，ID: {doc_id}，開始清理...")

            # 先刪除 Dify 上的舊文件；失敗就直接中止，不繼續往下建立新文件，
            # 避免同一份檔名在知識庫裡同時存在新舊兩份。
            delete_url = f"{DIFY_API_BASE}/datasets/{dataset_id}/documents/{doc_id}"
            del_res = requests.delete(delete_url, headers=headers)
            if del_res.status_code in (200, 204):
                print(f"[Dify刪除] 已成功刪除舊文件，ID: {doc_id}")
            else:
                raise Exception(
                    f"刪除舊文件失敗（ID: {doc_id}），狀態碼: {del_res.status_code}，"
                    f"內容: {del_res.text}，已中止上傳流程以避免產生重複文件。"
                )

            # 優先嘗試：直接從文件列表回傳的 metadata 解析舊 UUID
            old_uuid = get_file_uuid_from_metadata(doc)

            if old_uuid:
                print(f"[狀態查詢] 從 metadata 解析出舊檔案 UUID: {old_uuid}")
                safe_unlink(PDF_DIR / f"{old_uuid}.pdf")
                safe_unlink(DOCX_MODIFIED_DIR / f"{old_uuid}.docx")
                safe_unlink(DOCX_ORIGINAL_DIR / f"{old_uuid}.docx")
            else:
                # 退回舊機制：從 segment 內文用 regex 解析 PDF 網址取得 UUID
                print("[狀態查詢] metadata 中未找到 file_uuid，退回舊的內文解析機制。")
                segments_url = f"{DIFY_API_BASE}/datasets/{dataset_id}/documents/{doc_id}/segments"
                seg_res = requests.get(segments_url, headers=headers)

                if seg_res.status_code == 200:
                    seg_data = seg_res.json().get("data", [])
                    if seg_data:
                        first_chunk_text = seg_data[0].get("content", "")
                        match = re.search(r'/static/pdf/([a-f0-9\-]+\.pdf)', first_chunk_text)
                        if match:
                            old_pdf_filename = match.group(1)
                            old_uuid = Path(old_pdf_filename).stem
                            print(f"[狀態查詢] 解析出舊檔案 UUID: {old_uuid}")
                            safe_unlink(PDF_DIR / old_pdf_filename)
                            safe_unlink(DOCX_MODIFIED_DIR / f"{old_uuid}.docx")
                            safe_unlink(DOCX_ORIGINAL_DIR / f"{old_uuid}.docx")
                        else:
                            print("[狀態查詢] 警告：無法從舊文件的 segment 內容解析出 UUID。")

            return True

    print("[狀態查詢] 知識庫中未找到同名文件，將執行新增流程。")
    return False


def normalize_and_solidify_list_numbering(doc: docx.Document):
    """
    1. 算出動態多層級數字 (1., 1.1., 1.3.1.) 並寫死進內文文字首位。
    2. 刪除原本 Word XML 內的 w:numPr (動態縮排與編號標記)。
    """
    level_counters = {}

    for p in doc.paragraphs:
        pPr = p._p.get_or_add_pPr()
        numPr = pPr.find(docx.oxml.ns.qn('w:numPr'))

        if numPr is not None:
            ilvl_elem = numPr.find(docx.oxml.ns.qn('w:ilvl'))
            ilvl = int(ilvl_elem.get(docx.oxml.ns.qn('w:val'))) if ilvl_elem is not None else 0

            keys_to_del = [k for k in level_counters.keys() if k > ilvl]
            for k in keys_to_del:
                del level_counters[k]

            level_counters[ilvl] = level_counters.get(ilvl, 0) + 1

            full_number_parts = [str(level_counters[i]) for i in range(ilvl + 1) if i in level_counters]
            prefix = ".".join(full_number_parts) + " "

            text_content = p.text.strip()
            has_hardcoded_number = re.match(r'^(\d+(\.\d+)*[\.、\)]|\(\d+\))', text_content)

            if text_content and not has_hardcoded_number:
                if p.runs:
                    p.runs[0].text = prefix + p.runs[0].text
                else:
                    p.text = prefix + p.text

            pPr.remove(numPr)


def inject_pdf_url_and_chunk_tag_safe(docx_path: Path, pdf_url: str, display_filename: str):
    print(f"[Word注入] 開始實體化編號並注入 PDF 網址: {docx_path.name}")
    doc = docx.Document(docx_path)

    normalize_and_solidify_list_numbering(doc)

    pdf_text = f"【資料來源 PDF】: [{display_filename}]({pdf_url})"

    p_first = doc.paragraphs[0]
    p_first.insert_paragraph_before("===CHUNK===")
    p_first.insert_paragraph_before(pdf_text)
    p_first.insert_paragraph_before("===CHUNK===")

    doc.save(docx_path)
    del doc
    gc.collect()
    print("[Word注入] 注入完成並儲存檔案。")


async def wait_for_dify_indexing(dataset_id: str, document_id: str, headers: dict, max_retries: int = 60, delay: int = 5) -> dict:
    print(f"[狀態輪詢] 開始輪詢文件索引狀態，文件 ID: {document_id}")
    for i in range(max_retries):
        url = f"{DIFY_API_BASE}/datasets/{dataset_id}/documents/{document_id}"
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            doc_info = response.json()
            doc_data = doc_info.get("document", doc_info)
            status = doc_data.get("indexing_status")
            
            print(f"[狀態輪詢] 第 {i+1} 次檢查，目前狀態: {status}")
            
            if status == "completed":
                print("[狀態輪詢] 索引完成！")
                return doc_info
            elif status == "error":
                print("[狀態輪詢] 索引失敗，Dify 返回錯誤狀態。")
                return None
        else:
            print(f"[狀態輪詢] 請求失敗，狀態碼: {response.status_code}")
            
        await asyncio.sleep(delay)
        
    print("[狀態輪詢] 達到最大重試次數，索引發生超時。")
    return None


@app.post("/upload-to-dify")
async def process_and_upload_to_dify(file: UploadFile = File(...), dataset_id: str = Form(...)):
    # 變更：將 HTTPException 改為回傳帶有 status: error 的 JSON 物件
    if not file.filename.endswith(".docx"):
        return {
            "status": "error",
            "detail": "僅支援 .docx 格式檔案"
        }

    original_filename = file.filename
    print(f"\n========== 開始處理上傳請求: {original_filename}（知識庫: {dataset_id}）==========")
    headers = {"Authorization": f"Bearer {DIFY_API_KEY}"}

    file_uuid = str(uuid.uuid4())
    print(f"[初始化] 生成新檔案 UUID: {file_uuid}")
    
    original_docx_path = DOCX_ORIGINAL_DIR / f"{file_uuid}.docx"
    modified_docx_path = DOCX_MODIFIED_DIR / f"{file_uuid}.docx"
    pdf_path = PDF_DIR / f"{file_uuid}.pdf"

    try:
        file_bytes = await file.read()
        with open(original_docx_path, "wb") as buffer:
            buffer.write(file_bytes)
        with open(modified_docx_path, "wb") as buffer:
            buffer.write(file_bytes)
        print("[儲存檔案] 原始檔與處理檔已寫入本地儲存區。")

        doc_for_indent_fix = docx.Document(modified_docx_path)
        fix_list_indentation_for_libreoffice(doc_for_indent_fix)
        doc_for_indent_fix.save(modified_docx_path)
        del doc_for_indent_fix
        gc.collect()

        convert_docx_to_pdf(modified_docx_path, pdf_path)

        pdf_url = f"{SERVER_BASE_URL}/static/pdf/{file_uuid}.pdf"

        inject_pdf_url_and_chunk_tag_safe(modified_docx_path, pdf_url, original_filename)

        # 若知識庫中已有同名文件，先刪除舊文件並清掉對應的本地舊檔案，
        # 之後一律走「建立新文件」流程（不再用 update-by-file，
        # 因為 update-by-file 無法變更既有文件的 doc_form）。
        delete_existing_document_and_clean_local(dataset_id, original_filename, headers)

        data = {
            "process_rule": {
                "mode": "hierarchical",
                "rules": {
                    "pre_processing_rules": [
                        {"id": "remove_extra_spaces", "enabled": True},
                        {"id": "remove_urls_emails", "enabled": False}
                    ],
                    "segmentation": {
                        "separator": "===CHUNK===",
                        "max_tokens": 4000
                    },
                    "parent_mode": "paragraph",
                    "subchunk_segmentation": {
                        "separator": "\n\n",
                        "max_tokens": 4000
                    }
                }
            },
            "indexing_technique": "high_quality",
            "doc_form": "hierarchical_model",
            "retrieval_model": {
                "search_method": "semantic_search",
                "top_k": 3,
                "reranking_enable": False,
                "score_threshold_enabled": False
            }
        }

        with open(modified_docx_path, "rb") as docx_file:
            files = {
                "file": (original_filename, docx_file, "application/vnd.openxmlformats-officedocument.wordprocessingml.document"),
                "data": (None, json.dumps(data), "application/json")
            }

            print("[Dify上傳] 開始呼叫 API 建立新文件")
            api_url = f"{DIFY_API_BASE}/datasets/{dataset_id}/document/create-by-file"

            response = requests.post(api_url, headers=headers, files=files)

        if response.status_code != 200:
            print(f"[API失敗] 回應內容: {response.text}")
            raise Exception(f"Dify API 操作失敗: {response.text}")

        dify_response_data = response.json()
        target_doc_id = dify_response_data.get("document", {}).get("id")

        if not target_doc_id:
            raise Exception("無法從 Dify 回應中取得文件 ID")

        # 新增：文件建立/更新成功後，把這次的 file_uuid、pdf_url 一起寫入該文件的 metadata，
        # 之後要對照本地三個資料夾的檔案，就不用再靠內文 regex 解析。
        set_document_metadata(dataset_id, target_doc_id, file_uuid, pdf_url, headers)

        final_dify_response = await wait_for_dify_indexing(dataset_id, target_doc_id, headers)

        if final_dify_response:
            print("========== 處理流程全數成功完成 ==========\n")
            return {
                "status": "success",
                "original_filename": original_filename,
                "dataset_id": dataset_id,
                "pdf_url": pdf_url,
                "dify_response": final_dify_response
            }
        else:
            raise Exception("Dify 索引失敗或發生超時錯誤")

    except Exception as e:
        print(f"\n[錯誤處理] 流程發生例外，清理本次產生的暫存檔，錯誤內容: {str(e)}")
        safe_unlink(original_docx_path)
        safe_unlink(modified_docx_path)
        safe_unlink(pdf_path)
        print("========== 錯誤處理執行完畢 ==========\n")
        # 變更：移除 raise HTTPException，改回傳 JSON 讓 Dify 判讀
        return {
            "status": "error",
            "detail": str(e)
        }