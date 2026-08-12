FROM python:3.11-slim

# 安裝 LibreOffice 與中文字型，防止轉檔時中文變成亂碼框框
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice \
    fonts-noto-cjk \
    fonts-wqy-zenhei \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# 建立程式碼實際使用的三個子目錄
RUN mkdir -p storage/docx_original storage/docx_modified storage/pdf

EXPOSE 8000

CMD ["uvicorn", "app:app", "--host", "0.0.0.0", "--port", "8000"]