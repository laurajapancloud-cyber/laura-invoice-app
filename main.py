import os
import json
import secrets
import psycopg2
from psycopg2.extras import RealDictCursor
from typing import Annotated, List, Dict, Any, Optional
import base64
from io import BytesIO
import datetime
from zoneinfo import ZoneInfo
import time
import re
import requests
from supabase import create_client, Client as SupabaseClient

from fastapi import FastAPI, UploadFile, File, Depends, HTTPException, status, Response, Request, Form
from fastapi.security import HTTPBasic, HTTPBasicCredentials
from fastapi.templating import Jinja2Templates
from fastapi.responses import HTMLResponse, JSONResponse
import google.generativeai as genai
from weasyprint import HTML
from dotenv import load_dotenv
import openpyxl
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.utils.units import pixels_to_EMU
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker
from openpyxl.drawing.xdr import XDRPositiveSize2D
import barcode
from barcode.writer import ImageWriter
from pydantic import BaseModel

def get_jst_now():
    return datetime.datetime.now(ZoneInfo("Asia/Tokyo"))

# Conditional imports
try:
    from openai import OpenAI, AzureOpenAI
except ImportError:
    OpenAI = None
    AzureOpenAI = None

try:
    from google.oauth2 import service_account
    from google.cloud import vision
    from googleapiclient.discovery import build as gdrive_build
    from googleapiclient.http import MediaInMemoryUpload
except ImportError:
    vision = None
    gdrive_build = None
    MediaInMemoryUpload = None

# Load environment variables
load_dotenv()

APP_USERNAME = os.getenv("APP_USERNAME")
APP_PASSWORD = os.getenv("APP_PASSWORD")
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
GOOGLE_CREDENTIALS_JSON = os.getenv("GOOGLE_CREDENTIALS_JSON")
AZURE_OPENAI_KEY = os.getenv("AZURE_OPENAI_KEY")
AZURE_OPENAI_ENDPOINT = os.getenv("AZURE_OPENAI_ENDPOINT")
GDRIVE_FOLDER_ID = os.getenv("GDRIVE_FOLDER_ID")
GDRIVE_WEBHOOK_URL = os.getenv("GDRIVE_WEBHOOK_URL")

SUPABASE_URL = os.getenv("SUPABASE_URL")
SUPABASE_SERVICE_KEY = os.getenv("SUPABASE_SERVICE_KEY")
STORAGE_BUCKET = "invoices"

supabase_client: Optional[SupabaseClient] = None
if SUPABASE_URL and SUPABASE_SERVICE_KEY:
    supabase_client = create_client(SUPABASE_URL, SUPABASE_SERVICE_KEY)

def storage_upload(path: str, data: bytes, mime: str) -> str:
    """Supabase Storage にアップロードし、保存パスを返す"""
    if not supabase_client:
        raise Exception("Supabase Storage が未設定です")
    try:
        supabase_client.storage.from_(STORAGE_BUCKET).remove([path])
    except: pass
    supabase_client.storage.from_(STORAGE_BUCKET).upload(
        path, data, file_options={"content-type": mime, "upsert": "true"}
    )
    return path

def storage_download(path: str) -> bytes:
    """Storage から bytes を取得"""
    if not supabase_client:
        raise Exception("Supabase Storage が未設定です")
    return supabase_client.storage.from_(STORAGE_BUCKET).download(path)
    print("Warning: Missing required environment variables.")

# AI Initialization
genai.configure(api_key=GEMINI_API_KEY)
gemini_model = genai.GenerativeModel("gemini-2.5-flash")

openai_client = None
if OPENAI_API_KEY and OpenAI:
    openai_client = OpenAI(api_key=OPENAI_API_KEY)

azure_client = None
if AZURE_OPENAI_KEY and AZURE_OPENAI_ENDPOINT and AzureOpenAI:
    # URLの末尾が /openai の場合は自動調整 (AzureOpenAIが内部で付与するため)
    base_endpoint = AZURE_OPENAI_ENDPOINT.rstrip('/')
    if base_endpoint.endswith('/openai'):
        base_endpoint = base_endpoint[:-7]
        
    azure_client = AzureOpenAI(
        api_key=AZURE_OPENAI_KEY,
        api_version="2024-02-01",
        azure_endpoint=base_endpoint,
        default_headers={"Ocp-Apim-Subscription-Key": AZURE_OPENAI_KEY}
    )

vision_client = None
if GOOGLE_CREDENTIALS_JSON and vision:
    try:
        creds_dict = json.loads(GOOGLE_CREDENTIALS_JSON)
        creds = service_account.Credentials.from_service_account_info(creds_dict)
        vision_client = vision.ImageAnnotatorClient(credentials=creds)
    except Exception as e:
        print(f"Warning: Cloud Vision init failed: {e}")

def get_drive_service():
    """Google Drive API service (uses same service account JSON as Vision)."""
    if not GOOGLE_CREDENTIALS_JSON or not gdrive_build:
        return None
    try:
        creds_dict = json.loads(GOOGLE_CREDENTIALS_JSON)
        creds = service_account.Credentials.from_service_account_info(
            creds_dict,
            scopes=["https://www.googleapis.com/auth/drive.file"]
        )
        return gdrive_build("drive", "v3", credentials=creds, cache_discovery=False)
    except Exception as e:
        print(f"Drive init failed: {e}")
        return None

# ==================== PostgreSQL (Supabase) Database ====================
DATABASE_URL = os.getenv("DATABASE_URL")

def get_db():
    if not DATABASE_URL:
        raise Exception("DATABASE_URL environment variable is not set.")
    # SupabaseのURIに含まれるSSLモードに対応
    conn = psycopg2.connect(DATABASE_URL, cursor_factory=RealDictCursor)
    return conn

def init_db():
    # テーブル作成はSupabaseのSQL Editorで行うことを推奨しますが、念のためここでも試行します
    conn = get_db()
    with conn.cursor() as cur:
        cur.execute("""
            CREATE TABLE IF NOT EXISTS customers (
                id SERIAL PRIMARY KEY,
                name TEXT NOT NULL,
                discount_rate INTEGER NOT NULL DEFAULT 35,
                created_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP
            );
            CREATE TABLE IF NOT EXISTS invoices (
                id SERIAL PRIMARY KEY,
                invoice_number TEXT NOT NULL UNIQUE,
                customer_name TEXT NOT NULL,
                discount_rate INTEGER NOT NULL,
                total_net_amount INTEGER DEFAULT 0,
                total_tax_amount INTEGER DEFAULT 0,
                total_grand_total INTEGER DEFAULT 0,
                item_count INTEGER DEFAULT 0,
                pdf_data BYTEA,
                excel_data BYTEA,
                detail_pdf_data BYTEA,
                detail_excel_data BYTEA,
                created_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP
            );
            CREATE TABLE IF NOT EXISTS invoice_items (
                id SERIAL PRIMARY KEY,
                invoice_id INTEGER NOT NULL REFERENCES invoices(id) ON DELETE CASCADE,
                code TEXT,
                color TEXT,
                size TEXT,
                unit_price INTEGER DEFAULT 0,
                quantity INTEGER DEFAULT 1,
                net_amount INTEGER DEFAULT 0
            );
            CREATE TABLE IF NOT EXISTS api_usage (
                id SERIAL PRIMARY KEY,
                ai_model TEXT NOT NULL,
                image_count INTEGER DEFAULT 0,
                created_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP
            );
        """)
        
        # 初期データの投入（空の場合のみ）
        cur.execute("SELECT COUNT(*) FROM customers")
        if cur.fetchone()['count'] == 0:
            cur.executemany("INSERT INTO customers (name, discount_rate) VALUES (%s, %s)", [
                ("株式会社 タム 御中", 35),
                ("株式会社 サンプル 御中", 40),
            ])
    conn.commit()
    conn.close()

def generate_invoice_number():
    now = get_jst_now()
    prefix = f"LJ-{now.strftime('%Y%m')}"
    conn = get_db()
    with conn.cursor() as cur:
        cur.execute("SELECT COUNT(*) AS cnt FROM invoices WHERE invoice_number LIKE %s", (f"{prefix}%",))
        row = cur.fetchone()
        seq = (row['cnt'] or 0) + 1
    conn.close()
    return f"{prefix}-{seq:04d}"

# Initialize DB on startup
init_db()

# ==================== FastAPI App ====================
app = FastAPI()
security = HTTPBasic()
templates = Jinja2Templates(directory="templates")

def authenticate(credentials: Annotated[HTTPBasicCredentials, Depends(security)]):
    if not APP_USERNAME or not APP_PASSWORD:
        raise HTTPException(status_code=status.HTTP_500_INTERNAL_SERVER_ERROR, detail="Server config error.")
    is_correct_username = secrets.compare_digest(credentials.username.encode("utf8"), APP_USERNAME.encode("utf8"))
    is_correct_password = secrets.compare_digest(credentials.password.encode("utf8"), APP_PASSWORD.encode("utf8"))
    if not (is_correct_username and is_correct_password):
        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="Incorrect credentials", headers={"WWW-Authenticate": "Basic"})
    return credentials.username

@app.get("/", response_class=HTMLResponse)
async def index(request: Request, username: Annotated[str, Depends(authenticate)]):
    return templates.TemplateResponse(request=request, name="index.html")

@app.get("/health")
async def health():
    return {"status": "alive", "time": get_jst_now().isoformat()}

# ==================== Dashboard API ====================
@app.get("/api/dashboard")
async def get_dashboard(username: Annotated[str, Depends(authenticate)]):
    conn = get_db()
    with conn.cursor() as cur:
        now = get_jst_now()
        month_start = now.strftime("%Y-%m-01")
        
        # This month stats
        cur.execute(
            "SELECT COUNT(*) as cnt, COALESCE(SUM(total_grand_total),0) as total FROM invoices WHERE created_at >= %s",
            (month_start,)
        )
        inv_stats = cur.fetchone()
        
        # Monthly history (last 6 months)
        # Postgres uses to_char for formatting
        cur.execute(
            "SELECT to_char(created_at, 'YYYY-MM') as month, COUNT(*) as cnt, COALESCE(SUM(total_grand_total),0) as total FROM invoices GROUP BY month ORDER BY month DESC LIMIT 6"
        )
        monthly = cur.fetchall()
        
        # API usage this month
        cur.execute(
            "SELECT ai_model, COUNT(*) as cnt, COALESCE(SUM(image_count),0) as imgs FROM api_usage WHERE created_at >= %s GROUP BY ai_model",
            (month_start,)
        )
        usage = cur.fetchall()
    conn.close()
    
    # Estimate costs
    usage_list = []
    for u in usage:
        model = u["ai_model"]
        imgs = u["imgs"]
        if model == "gemini":
            cost = round(imgs * 0.1, 1)
        elif model in ["azure", "openai"]:
            cost = round(imgs * 0.5, 1)
        else:
            cost = 0
        usage_list.append({"model": model, "requests": u["cnt"], "images": imgs, "est_cost_yen": cost})
    
    return {
        "this_month": {"invoices": inv_stats["cnt"], "total_amount": inv_stats["total"]},
        "monthly": [dict(m) for m in monthly],
        "api_usage": usage_list
    }

# ==================== Customer Master API ====================
@app.get("/api/customers")
async def get_customers(username: Annotated[str, Depends(authenticate)]):
    conn = get_db()
    with conn.cursor() as cur:
        cur.execute("SELECT id, name, discount_rate FROM customers ORDER BY id")
        rows = cur.fetchall()
    conn.close()
    return [dict(r) for r in rows]

@app.post("/api/customers")
async def add_customer(username: Annotated[str, Depends(authenticate)], name: str = Form(...), discount_rate: int = Form(35)):
    conn = get_db()
    with conn.cursor() as cur:
        cur.execute("INSERT INTO customers (name, discount_rate) VALUES (%s, %s)", (name, discount_rate))
    conn.commit()
    conn.close()
    return {"status": "ok"}

@app.delete("/api/customers/{cid}")
async def delete_customer(cid: int, username: Annotated[str, Depends(authenticate)]):
    conn = get_db()
    with conn.cursor() as cur:
        cur.execute("DELETE FROM customers WHERE id = %s", (cid,))
    conn.commit()
    conn.close()
    return {"status": "ok"}

# ==================== History API ====================
@app.get("/api/history")
async def get_history(username: Annotated[str, Depends(authenticate)]):
    conn = get_db()
    with conn.cursor() as cur:
        cur.execute(
            "SELECT id, invoice_number, customer_name, total_grand_total, item_count, to_char(created_at, 'YYYY-MM-DD\"T\"HH24:MI:SS') as created_at FROM invoices ORDER BY id DESC"
        )
        rows = cur.fetchall()
    conn.close()
    return [dict(r) for r in rows]

@app.get("/api/history/{inv_id}/pdf")
async def download_history_pdf(inv_id: int, username: Annotated[str, Depends(authenticate)]):
    conn = get_db()
    with conn.cursor() as cur:
        cur.execute("SELECT pdf_data, invoice_number FROM invoices WHERE id = %s", (inv_id,))
        row = cur.fetchone()
    conn.close()
    if not row or not row["pdf_data"]:
        raise HTTPException(status_code=404, detail="PDF not found")
    return Response(content=row["pdf_data"], media_type="application/pdf",
                    headers={"Content-Disposition": f'attachment; filename="{row["invoice_number"]}.pdf"'})

@app.get("/api/history/{inv_id}/excel")
async def download_history_excel(inv_id: int, username: Annotated[str, Depends(authenticate)]):
    conn = get_db()
    with conn.cursor() as cur:
        cur.execute("SELECT excel_data, invoice_number FROM invoices WHERE id = %s", (inv_id,))
        row = cur.fetchone()
    conn.close()
    if not row or not row["excel_data"]:
        raise HTTPException(status_code=404, detail="Excel not found")
    return Response(content=row["excel_data"],
                    media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    headers={"Content-Disposition": f'attachment; filename="{row["invoice_number"]}.xlsx"'})

@app.get("/api/history/{inv_id}/detail-pdf")
async def download_detail_pdf(inv_id: int, username: Annotated[str, Depends(authenticate)]):
    conn = get_db()
    with conn.cursor() as cur:
        cur.execute("SELECT detail_pdf_data, invoice_number FROM invoices WHERE id = %s", (inv_id,))
        row = cur.fetchone()
    conn.close()
    if not row or not row["detail_pdf_data"]:
        raise HTTPException(status_code=404, detail="Detail PDF not found")
    return Response(content=row["detail_pdf_data"], media_type="application/pdf",
                    headers={"Content-Disposition": f'attachment; filename="{row["invoice_number"]}_detail.pdf"'})

@app.get("/api/history/{inv_id}/detail-excel")
async def download_detail_excel(inv_id: int, username: Annotated[str, Depends(authenticate)]):
    conn = get_db()
    with conn.cursor() as cur:
        cur.execute("SELECT detail_excel_data, invoice_number FROM invoices WHERE id = %s", (inv_id,))
        row = cur.fetchone()
    conn.close()
    if not row or not row["detail_excel_data"]:
        raise HTTPException(status_code=404, detail="Detail Excel not found")
    return Response(content=row["detail_excel_data"], media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    headers={"Content-Disposition": f'attachment; filename="{row["invoice_number"]}_detail.xlsx"'})

@app.get("/api/history/{inv_id}/items")
async def get_history_items(inv_id: int, username: Annotated[str, Depends(authenticate)]):
    conn = get_db()
    with conn.cursor() as cur:
        cur.execute("SELECT code, color, size, unit_price, quantity FROM invoice_items WHERE invoice_id = %s", (inv_id,))
        rows = cur.fetchall()
        cur.execute("SELECT invoice_number, customer_name, discount_rate FROM invoices WHERE id = %s", (inv_id,))
        inv = cur.fetchone()
    conn.close()
    if not inv:
        raise HTTPException(status_code=404, detail="Invoice not found")
    return {"invoice_number": inv["invoice_number"], "customer_name": inv["customer_name"], "discount_rate": inv["discount_rate"], "items": [dict(r) for r in rows]}

@app.delete("/api/history/{inv_id}")
async def delete_history(inv_id: int, username: Annotated[str, Depends(authenticate)]):
    conn = get_db()
    with conn.cursor() as cur:
        cur.execute("DELETE FROM invoice_items WHERE invoice_id = %s", (inv_id,))
        cur.execute("DELETE FROM invoices WHERE id = %s", (inv_id,))
    conn.commit()
    conn.close()
    return {"status": "ok"}

# ==================== AI Analysis API ====================
@app.post("/analyze-images")
async def analyze_images(
    request: Request,
    username: Annotated[str, Depends(authenticate)],
    files: List[UploadFile] = File(...),
    ai_model: str = Form("gemini")
):
    try:
        image_parts = []
        for file in files:
            image_bytes = await file.read()
            b64 = base64.b64encode(image_bytes).decode("utf-8")
            mime_type = file.content_type or "image/jpeg"
            image_parts.append({"mime_type": mime_type, "data": image_bytes, "base64": b64})

        prompt = """
あなたはアパレルブランドのデータ入力アシスタントです。提供された複数の商品タグの画像からそれぞれの情報を抽出し、厳密にJSON配列（リスト）の形式のみで出力してください。マークダウンの装飾(```jsonなど)は含めないでください。
スキーマ: [{"code": "品番(例:148-3101)", "color": "カラー(例:24)", "size": "サイズ(例:46)", "unit_price": 単価の数値(例:38000)}, ...]
複数の商品がある場合は、配列内に複数のオブジェクトを含めてください。
""".strip()

        items_data = []

        if ai_model == "vision":
            if not vision_client:
                raise HTTPException(status_code=500, detail="Cloud VisionのJSONキーが未設定です。")
            for part in image_parts:
                image = vision.Image(content=part["data"])
                response = vision_client.text_detection(image=image)
                if response.error.message:
                    raise HTTPException(status_code=500, detail=f"Vision Error: {response.error.message}")
                raw_text = response.text_annotations[0].description if response.text_annotations else ""
                code_match = re.search(r'[A-Za-z0-9]+-[A-Za-z0-9]+', raw_text)
                col_match = re.search(r'(?i)col(?:or)?[\s\.:]*(\\d{1,3})', raw_text)
                sz_match = re.search(r'(?i)size[\s\\.:]*([\w]+)', raw_text)
                if not sz_match:
                    sz_match = re.search(r'\b(3[68]|4[02468]|50)\b', raw_text)
                price_match = re.search(r'[¥￥]\s*([\d,]+)', raw_text)
                if not price_match:
                    price_match = re.search(r'\b(\d{1,3}(?:,\d{3})+)\b', raw_text)
                unit_price = 0
                if price_match:
                    try: unit_price = int(price_match.group(1).replace(',', ''))
                    except: pass
                items_data.append({
                    "code": code_match.group(0) if code_match else "-",
                    "color": col_match.group(1) if col_match else "-",
                    "size": sz_match.group(1) if sz_match else "-",
                    "unit_price": unit_price
                })
            return JSONResponse({"items": items_data})

        elif ai_model == "azure":
            if not azure_client:
                raise HTTPException(status_code=500, detail="Azure OpenAIのキー/エンドポイントが未設定です。")
            content_list = [{"type": "text", "text": prompt}]
            for part in image_parts:
                content_list.append({"type": "image_url", "image_url": {"url": f"data:{part['mime_type']};base64,{part['base64']}"}})
            response = azure_client.chat.completions.create(model="gpt-4o", messages=[{"role": "user", "content": content_list}], max_completion_tokens=2000)
            raw_text = response.choices[0].message.content.strip()

        elif ai_model == "openai":
            if not openai_client:
                raise HTTPException(status_code=500, detail="OpenAI APIキーが未設定です。")
            content_list = [{"type": "text", "text": prompt}]
            for part in image_parts:
                content_list.append({"type": "image_url", "image_url": {"url": f"data:{part['mime_type']};base64,{part['base64']}"}})
            response = openai_client.chat.completions.create(model="gpt-4o-mini", messages=[{"role": "user", "content": content_list}], max_tokens=2000)
            raw_text = response.choices[0].message.content.strip()

        else:
            contents = [prompt] + [{"mime_type": p["mime_type"], "data": p["data"]} for p in image_parts]
            try:
                response = gemini_model.generate_content(contents)
                raw_text = response.text.strip()
            except Exception as e:
                error_str = str(e)
                print(f"[GEMINI ERROR] {error_str}")
                if "429" in error_str or "RESOURCE_EXHAUSTED" in error_str or "Quota exceeded" in error_str or "rate_limit" in error_str:
                    raise HTTPException(status_code=429, detail="Geminiの無料枠上限に達しました。設定からAIモデルを切り替えてください。")
                raise HTTPException(status_code=500, detail=f"Gemini Error: {error_str}")

        if raw_text.startswith("```json"):
            raw_text = raw_text.replace("```json", "", 1).replace("```", "", 1).strip()
        elif raw_text.startswith("```"):
            raw_text = raw_text.replace("```", "", 2).strip()
        chunk_data = json.loads(raw_text)
        if not isinstance(chunk_data, list):
            chunk_data = [chunk_data]
        
        # Record API usage
        try:
            conn = get_db()
            with conn.cursor() as cur:
                cur.execute("INSERT INTO api_usage (ai_model, image_count) VALUES (%s, %s)", (ai_model, len(files)))
            conn.commit()
            conn.close()
        except: pass
        
        return JSONResponse({"items": chunk_data})

    except HTTPException:
        raise
    except Exception as e:
        import traceback
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(e))

# ==================== Product Detail Sheet Logic ====================
SIZE_COLUMNS = ["44", "46", "48", "50", "52"]

def build_detail_excel(invoice_number: str, customer_name: str, items: list) -> bytes:
    """商品明細表 Excel (サイズ別内訳)"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "商品明細表"

    fill_header = PatternFill(start_color="F4ECD8", end_color="F4ECD8", fill_type="solid")
    fill_meta = PatternFill(start_color="FFF8E7", end_color="FFF8E7", fill_type="solid")
    border_thin = Border(
        left=Side(style='thin', color='888888'),
        right=Side(style='thin', color='888888'),
        top=Side(style='thin', color='888888'),
        bottom=Side(style='thin', color='888888')
    )

    # 列幅
    widths = {'A': 4, 'B': 14, 'C': 7, 'D': 9, 'E': 7, 'F': 5, 'G': 5, 'H': 5, 'I': 5, 'J': 5, 'K': 14}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

    now = get_jst_now()
    reiwa = now.year - 2018

    # タイトル
    ws["A1"] = "商 品 明 細 表"
    ws["A1"].font = Font(bold=True, size=14)
    ws.merge_cells("A1:E1")
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")

    ws["G1"] = f"伝票番号: {invoice_number}"
    ws["G1"].font = Font(size=10, color="555555")
    ws.merge_cells("G1:K1")
    ws["G1"].alignment = Alignment(horizontal="right", vertical="center")

    ws["A2"] = f"取引先: {customer_name}"
    ws["A2"].font = Font(size=11)
    ws.merge_cells("A2:E2")

    ROWS_PER_SECTION = 5
    sections = max(1, (len(items) + ROWS_PER_SECTION - 1) // ROWS_PER_SECTION)

    section_start = 4
    for s in range(sections):
        date_row = section_start
        ws[f"I{date_row}"] = f"{reiwa}年"
        ws[f"J{date_row}"] = f"{now.month}月"
        ws[f"K{date_row}"] = f"{now.day}日"
        for c in ['I','J','K']:
            ws[f"{c}{date_row}"].alignment = Alignment(horizontal="center")
            ws[f"{c}{date_row}"].font = Font(size=9, color="555555")

        header_row = section_start + 1
        ws.row_dimensions[header_row].height = 22
        headers = {'A':'No.', 'B':'品番', 'C':'枚数', 'D':'上代', 'E':'カラー', 'F':'44', 'G':'46', 'H':'48', 'I':'50', 'J':'52', 'K':'備考'}
        for col, label in headers.items():
            cell = ws[f"{col}{header_row}"]
            cell.value = label
            cell.fill = fill_header
            cell.font = Font(bold=True, size=10)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = border_thin

        for i in range(ROWS_PER_SECTION):
            r = header_row + 1 + i
            ws.row_dimensions[r].height = 22
            item_idx = s * ROWS_PER_SECTION + i
            item = items[item_idx] if item_idx < len(items) else None
            ws[f"A{r}"] = i + 1
            ws[f"A{r}"].alignment = Alignment(horizontal="center", vertical="center")
            ws[f"A{r}"].font = Font(size=9, color="888888")

            if item:
                ws[f"B{r}"] = item.get("code", "")
                ws[f"C{r}"] = item.get("quantity", 0)
    ws.title = "伝票"
    fill_yellow = PatternFill(start_color="FFE699", end_color="FFE699", fill_type="solid")
    fill_blue = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    border_thin = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    
    ws.merge_cells("A1:I1")
    ws["A1"] = "御 納 品 書"
    ws["A1"].font = Font(size=24, bold=True)
    ws["A1"].alignment = Alignment(horizontal="center")
    
    ws["F1"] = "No."
    ws["G1"] = invoice_data['invoice_number']
    ws["G1"].fill = fill_yellow
    ws["F3"] = "日付"
    ws["G3"] = invoice_data['date']
    ws["A4"] = "店名"
    ws["A5"] = invoice_data['customer_name']
    ws["A5"].font = Font(size=14, bold=True)
    ws["A5"].fill = fill_yellow

    headers = ["品番", "カラー", "サイズ", "バーコード", "", "数量", "単価", "金額", "掛率"]
    h_cols = ["A", "B", "C", "D", "E", "F", "G", "H", "I"]
    start_row = 8
    for col, txt in zip(h_cols, headers):
        if not txt: continue
        cell = ws[f"{col}{start_row}"]
        cell.value = txt
        cell.fill = fill_blue
        cell.border = border_thin
        cell.alignment = Alignment(horizontal="center")

    ws.column_dimensions['A'].width = 15
    ws.column_dimensions['B'].width = 10
    ws.column_dimensions['C'].width = 10
    ws.column_dimensions['D'].width = 30
    ws.column_dimensions['F'].width = 10
    ws.column_dimensions['G'].width = 15
    ws.column_dimensions['H'].width = 15

    for i, item in enumerate(invoice_data["items"]):
        r = start_row + 1 + i
        ws.row_dimensions[r].height = 45
        ws[f"A{r}"] = item["code"]
        ws[f"B{r}"] = item["color"]
        ws[f"C{r}"] = item["size"]
        ws[f"F{r}"] = item["quantity"]
        ws[f"G{r}"] = item["unit_price"]
        ws[f"G{r}"].number_format = '#,##0'
        ws[f"H{r}"] = item["net_amount"]
        ws[f"H{r}"].number_format = '#,##0'
        ws[f"I{r}"] = f"{invoice_data['discount_rate']}%"
        
        for col in ["A","B","C","D","F","G","H","I"]:
            ws[f"{col}{r}"].border = border_thin
        
        # 簡易バーコード
        try:
            bc_str = str(item["code"])
            ean = barcode.get('code128', bc_str, writer=ImageWriter())
            bc_io = BytesIO()
            ean.write(bc_io)
            img = ExcelImage(bc_io)
            img.width, img.height = 120, 30
            marker = AnchorMarker(col=3, colOff=pixels_to_EMU(10), row=r-1, rowOff=pixels_to_EMU(5))
            img.anchor = OneCellAnchor(_marker=marker, ext=XDRPositiveSize2D(cx=pixels_to_EMU(120), cy=pixels_to_EMU(30)))
            ws.add_image(img)
        except: pass

    last_r = start_row + len(invoice_data["items"]) + 2
    ws.cell(row=last_r, column=7, value="小計")
    ws.cell(row=last_r, column=8, value=invoice_data["total_net_amount"])
    ws.cell(row=last_r+1, column=7, value="消費税")
    ws.cell(row=last_r+1, column=8, value=invoice_data["total_tax_amount"])
    ws.cell(row=last_r+2, column=7, value="合計金額")
    ws.cell(row=last_r+2, column=8, value=invoice_data["total_grand_total"])

    out = BytesIO()
    wb.save(out)
    return out.getvalue()

def build_all_files(invoice_data: dict) -> dict:
    """4ファイルまとめて生成"""
    return {
        "pdf": build_invoice_pdf(invoice_data),
        "excel": build_invoice_excel(invoice_data),
        "detail_pdf": build_detail_pdf(invoice_data["invoice_number"], invoice_data["customer_name"], invoice_data["items"]),
        "detail_excel": build_detail_excel(invoice_data["invoice_number"], invoice_data["customer_name"], invoice_data["items"]),
    }

def assemble_invoice_data(inv_info: dict, items_input: list, discount_rate: int) -> dict:
    """生成用データ準備"""
    processed = []
    total_net = total_tax = total_grand = 0
    for it in items_input:
        up = it.get("unit_price", 0)
        if isinstance(up, str): up = int(up.replace(',','').replace('¥','').strip() or '0')
        qty = max(1, int(it.get("quantity") or 1))
        net = int(up * (discount_rate / 100) * qty)
        tax = int(net * 0.1)
        grand = net + tax
        processed.append({
            "code": it.get("code") or "-", "color": it.get("color") or "-", "size": it.get("size") or "-",
            "unit_price": up, "quantity": qty, "net_amount": net, "tax_amount": tax, "grand_total": grand
        })
        total_net += net; total_tax += tax; total_grand += grand
    
    dt = inv_info.get("locked_at") or inv_info.get("created_at") or get_jst_now()
    if isinstance(dt, str):
        try: dt = datetime.datetime.fromisoformat(dt.replace("Z", "+00:00"))
        except: dt = get_jst_now()

    return {
        "invoice_number": inv_info["invoice_number"], "customer_name": inv_info["customer_name"],
        "discount_rate": discount_rate, "items": processed,
        "total_net_amount": total_net, "total_tax_amount": total_tax, "total_grand_total": total_grand,
        "issuer": "株式会社 ラウラジャパン", "date": dt.strftime("%Y年%m月%d日")
    }

# ==================== Document Generation API ====================
class DocumentRequest(BaseModel):
    invoice_id: Optional[int] = None
    customer_name: str
    discount_rate: int
    items: List[Dict[str, Any]]

@app.post("/generate-documents")
async def generate_documents(request: Request, username: Annotated[str, Depends(authenticate)], payload: DocumentRequest):
    try:
        conn = get_db()
        with conn.cursor() as cur:
            if payload.invoice_id:
                cur.execute("SELECT invoice_number, status FROM invoices WHERE id = %s", (payload.invoice_id,))
                row = cur.fetchone()
                if not row: raise HTTPException(404, "Invoice not found")
                if row["status"] == "locked": raise HTTPException(400, "確定済みの伝票は編集できません。")
                invoice_number = row["invoice_number"]
            else:
                invoice_number = generate_invoice_number()

            inv_data = assemble_invoice_data({"invoice_number": invoice_number, "customer_name": payload.customer_name}, payload.items, payload.discount_rate)

            if payload.invoice_id:
                cur.execute("""
                    UPDATE invoices SET customer_name=%s, discount_rate=%s, total_net_amount=%s, total_tax_amount=%s, total_grand_total=%s, 
                    item_count=%s, status='draft', locked_at=NULL,
                    pdf_storage_path=NULL, excel_storage_path=NULL, detail_pdf_storage_path=NULL, detail_excel_storage_path=NULL,
                    created_at=CURRENT_TIMESTAMP WHERE id=%s
                """, (payload.customer_name, payload.discount_rate, inv_data["total_net_amount"], inv_data["total_tax_amount"], inv_data["total_grand_total"], len(inv_data["items"]), payload.invoice_id))
                cur.execute("DELETE FROM invoice_items WHERE invoice_id=%s", (payload.invoice_id,))
                inv_id = payload.invoice_id
            else:
                cur.execute("""
                    INSERT INTO invoices (invoice_number, customer_name, discount_rate, total_net_amount, total_tax_amount, total_grand_total, item_count, status) 
                    VALUES (%s,%s,%s,%s,%s,%s,%s,'draft') RETURNING id
                """, (invoice_number, payload.customer_name, payload.discount_rate, inv_data["total_net_amount"], inv_data["total_tax_amount"], inv_data["total_grand_total"], len(inv_data["items"])))
                inv_id = cur.fetchone()['id']
            
            for item in inv_data["items"]:
                cur.execute("INSERT INTO invoice_items (invoice_id, code, color, size, unit_price, quantity, net_amount) VALUES (%s,%s,%s,%s,%s,%s,%s)",
                             (inv_id, item["code"], item["color"], item["size"], item["unit_price"], item["quantity"], item["net_amount"]))
        conn.commit()
        conn.close()

        files = build_all_files(inv_data)
        return JSONResponse({
            "invoice_id": inv_id, "invoice_number": invoice_number, "status": "draft",
            "pdf_base64": base64.b64encode(files["pdf"]).decode(), "excel_base64": base64.b64encode(files["excel"]).decode(),
            "detail_pdf_base64": base64.b64encode(files["detail_pdf"]).decode(), "detail_excel_base64": base64.b64encode(files["detail_excel"]).decode(),
        })
    except HTTPException: raise
    except Exception as e:
        import traceback; traceback.print_exc()
        raise HTTPException(500, str(e))

@app.post("/api/history/{inv_id}/lock")
async def lock_invoice(inv_id: int, username: Annotated[str, Depends(authenticate)]):
    conn = get_db()
    with conn.cursor() as cur:
        cur.execute("SELECT * FROM invoices WHERE id=%s", (inv_id,))
        inv = cur.fetchone()
        if not inv: raise HTTPException(404, "Not found")
        if inv["status"] == "locked": return {"status": "already_locked"}
        cur.execute("SELECT code,color,size,unit_price,quantity FROM invoice_items WHERE invoice_id=%s", (inv_id,))
        items = [dict(r) for r in cur.fetchall()]
    
    inv_data = assemble_invoice_data(dict(inv), items, inv["discount_rate"])
    files = build_all_files(inv_data)
    
    base = f"locked/{inv['invoice_number']}"
    p = {
        "pdf": storage_upload(f"{base}/invoice.pdf", files["pdf"], "application/pdf"),
        "excel": storage_upload(f"{base}/invoice.xlsx", files["excel"], "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
        "detail_pdf": storage_upload(f"{base}/detail.pdf", files["detail_pdf"], "application/pdf"),
        "detail_excel": storage_upload(f"{base}/detail.xlsx", files["detail_excel"], "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
    }
    
    with conn.cursor() as cur:
        cur.execute("""
            UPDATE invoices SET status='locked', locked_at=NOW(),
            pdf_storage_path=%s, excel_storage_path=%s, detail_pdf_storage_path=%s, detail_excel_storage_path=%s
            WHERE id=%s
        """, (p["pdf"], p["excel"], p["detail_pdf"], p["detail_excel"], inv_id))
        conn.commit()
    conn.close()
    return {"status": "locked"}

@app.post("/api/history/{inv_id}/unlock")
async def unlock_invoice(inv_id: int, username: Annotated[str, Depends(authenticate)]):
    conn = get_db()
    with conn.cursor() as cur:
        cur.execute("UPDATE invoices SET status='draft', locked_at=NULL WHERE id=%s", (inv_id,))
        conn.commit()
    conn.close()
    return {"status": "draft"}

# ==================== History & Serving API ====================
@app.get("/api/history")
async def get_history(username: Annotated[str, Depends(authenticate)]):
    conn = get_db()
    with conn.cursor() as cur:
        cur.execute("""
            SELECT id, invoice_number, customer_name, total_grand_total, item_count, status, 
            to_char(locked_at, 'YYYY-MM-DD"T"HH24:MI:SS') as locked_at, 
            to_char(created_at, 'YYYY-MM-DD\"T\"HH24:MI:SS') as created_at FROM invoices ORDER BY id DESC
        """)
        rows = cur.fetchall()
    conn.close()
    return [dict(r) for r in rows]

def _serve_file(inv_id: int, kind: str):
    conn = get_db()
    with conn.cursor() as cur:
        cur.execute("SELECT * FROM invoices WHERE id=%s", (inv_id,))
        inv = cur.fetchone()
        if not inv: raise HTTPException(404, "Not found")
        cur.execute("SELECT code,color,size,unit_price,quantity FROM invoice_items WHERE invoice_id=%s", (inv_id,))
        items = [dict(r) for r in cur.fetchall()]
    conn.close()
    
    path_field = f"{kind.replace('-','_')}_storage_path"
    mime_map = {
        "pdf": ("application/pdf", "pdf"),
        "excel": ("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "xlsx"),
        "detail-pdf": ("application/pdf", "pdf"),
        "detail-excel": ("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "xlsx"),
    }
    mime, ext = mime_map[kind]
    
    if inv["status"] == "locked" and inv.get(path_field):
        try:
            data = storage_download(inv[path_field])
            return Response(content=data, media_type=mime, headers={"Content-Disposition": f'attachment; filename="{inv["invoice_number"]}_{kind}.{ext}"'})
        except: pass
    
    # Draft or Storage missing -> Re-generate
    inv_data = assemble_invoice_data(dict(inv), items, inv["discount_rate"])
    files = build_all_files(inv_data)
    key = kind.replace("-", "_")
    return Response(content=files[key], media_type=mime, headers={"Content-Disposition": f'attachment; filename="{inv["invoice_number"]}_{kind}.{ext}"'})

@app.get("/api/history/{inv_id}/pdf")
async def dl_pdf(inv_id: int, username: Annotated[str, Depends(authenticate)]): return _serve_file(inv_id, "pdf")
@app.get("/api/history/{inv_id}/excel")
async def dl_excel(inv_id: int, username: Annotated[str, Depends(authenticate)]): return _serve_file(inv_id, "excel")
@app.get("/api/history/{inv_id}/detail-pdf")
async def dl_dpdf(inv_id: int, username: Annotated[str, Depends(authenticate)]): return _serve_file(inv_id, "detail-pdf")
@app.get("/api/history/{inv_id}/detail-excel")
async def dl_dxl(inv_id: int, username: Annotated[str, Depends(authenticate)]): return _serve_file(inv_id, "detail-excel")

@app.get("/api/history/{inv_id}/items")
async def get_history_items(inv_id: int, username: Annotated[str, Depends(authenticate)]):
    conn = get_db()
    with conn.cursor() as cur:
        cur.execute("SELECT code, color, size, unit_price, quantity FROM invoice_items WHERE invoice_id = %s", (inv_id,))
        rows = cur.fetchall()
        cur.execute("SELECT invoice_number, customer_name, discount_rate FROM invoices WHERE id = %s", (inv_id,))
        inv = cur.fetchone()
    conn.close()
    if not inv: raise HTTPException(status_code=404, detail="Invoice not found")
    return {"invoice_number": inv["invoice_number"], "customer_name": inv["customer_name"], "discount_rate": inv["discount_rate"], "items": [dict(r) for r in rows]}

@app.delete("/api/history/{inv_id}")
async def delete_history(inv_id: int, username: Annotated[str, Depends(authenticate)]):
    conn = get_db()
    with conn.cursor() as cur:
        cur.execute("DELETE FROM invoice_items WHERE invoice_id = %s", (inv_id,))
        cur.execute("DELETE FROM invoices WHERE id = %s", (inv_id,))
    conn.commit()
    conn.close()
    return {"status": "ok"}

@app.post("/api/history/{inv_id}/upload-drive")
async def upload_to_drive(inv_id: int, username: Annotated[str, Depends(authenticate)]):
    if not GDRIVE_WEBHOOK_URL: raise HTTPException(500, "GDRIVE_WEBHOOK_URLが未設定です")
    conn = get_db()
    with conn.cursor() as cur:
        cur.execute("SELECT * FROM invoices WHERE id=%s", (inv_id,))
        inv = cur.fetchone()
        cur.execute("SELECT code,color,size,unit_price,quantity FROM invoice_items WHERE invoice_id=%s", (inv_id,))
        items = [dict(r) for r in cur.fetchall()]
    conn.close()
    if not inv: raise HTTPException(404, "Not found")

    if inv["status"] == "locked" and inv.get("pdf_storage_path"):
        files = {
            "pdf": storage_download(inv["pdf_storage_path"]),
            "excel": storage_download(inv["excel_storage_path"]),
            "detail_pdf": storage_download(inv["detail_pdf_storage_path"]),
            "detail_excel": storage_download(inv["detail_excel_storage_path"]),
        }
    else:
        # 割引率がNULLの場合は100%として扱う
        dr = inv.get("discount_rate")
        if dr is None: dr = 100
        inv_data = assemble_invoice_data(dict(inv), items, dr)
        files = build_all_files(inv_data)

    inv_num = inv["invoice_number"]
    cust = (inv["customer_name"] or "無名").replace("/", "_").replace("\\", "_")
    uploaded = []

    def _up_sync(fn, mime, data):
        resp = requests.post(GDRIVE_WEBHOOK_URL, json={
            "folderId": GDRIVE_FOLDER_ID, "filename": fn, "mime": mime, "base64": base64.b64encode(data).decode()
        }, timeout=30)
        if resp.status_code != 200: raise Exception(f"GAS error: {resp.text}")
        return resp.json()

    try:
        targets = [
            ("pdf", f"{inv_num}_{cust}_伝票.pdf", "application/pdf", files["pdf"]),
            ("excel", f"{inv_num}_{cust}_伝票.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", files["excel"]),
            ("detail_pdf", f"{inv_num}_{cust}_明細表.pdf", "application/pdf", files["detail_pdf"]),
            ("detail_excel", f"{inv_num}_{cust}_明細表.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", files["detail_excel"]),
        ]
        for typ, fn, mime, data in targets:
            res = _up_sync(fn, mime, data)
            uploaded.append({"type": typ, "name": fn, "url": res.get("url")})

        return {"status": "ok", "uploaded": uploaded}
    except Exception as e:
        import traceback; traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"Drive保存失敗(GAS): {e}")

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
