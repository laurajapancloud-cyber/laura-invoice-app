import os
import json
import secrets
import sqlite3
from typing import Annotated, List, Dict, Any, Optional
import base64
from io import BytesIO
import datetime
from zoneinfo import ZoneInfo
import time
import re

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

if not all([APP_USERNAME, APP_PASSWORD, GEMINI_API_KEY]):
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

# ==================== SQLite Database ====================
DB_PATH = "invoices.db"

def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    return conn

def init_db():
    conn = get_db()
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS customers (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            discount_rate INTEGER NOT NULL DEFAULT 35,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
        CREATE TABLE IF NOT EXISTS invoices (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            invoice_number TEXT NOT NULL UNIQUE,
            customer_name TEXT NOT NULL,
            discount_rate INTEGER NOT NULL,
            total_net_amount INTEGER DEFAULT 0,
            total_tax_amount INTEGER DEFAULT 0,
            total_grand_total INTEGER DEFAULT 0,
            item_count INTEGER DEFAULT 0,
            pdf_data BLOB,
            excel_data BLOB,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
        CREATE TABLE IF NOT EXISTS invoice_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            invoice_id INTEGER NOT NULL,
            code TEXT,
            color TEXT,
            size TEXT,
            unit_price INTEGER DEFAULT 0,
            quantity INTEGER DEFAULT 1,
            net_amount INTEGER DEFAULT 0,
            FOREIGN KEY (invoice_id) REFERENCES invoices(id) ON DELETE CASCADE
        );
        CREATE TABLE IF NOT EXISTS api_usage (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ai_model TEXT NOT NULL,
            image_count INTEGER DEFAULT 0,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
    """)
    # Insert default customers if table is empty
    count = conn.execute("SELECT COUNT(*) FROM customers").fetchone()[0]
    if count == 0:
        conn.executemany("INSERT INTO customers (name, discount_rate) VALUES (?, ?)", [
            ("株式会社 タム 御中", 35),
            ("株式会社 サンプル 御中", 40),
        ])
    conn.commit()
    conn.close()

def generate_invoice_number():
    now = get_jst_now()
    prefix = f"LJ-{now.strftime('%Y%m')}"
    conn = get_db()
    row = conn.execute(
        "SELECT COUNT(*) FROM invoices WHERE invoice_number LIKE ?",
        (f"{prefix}%",)
    ).fetchone()
    seq = (row[0] or 0) + 1
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
    now = get_jst_now()
    month_start = now.strftime("%Y-%m-01")
    
    # This month stats
    inv_stats = conn.execute(
        "SELECT COUNT(*) as cnt, COALESCE(SUM(total_grand_total),0) as total FROM invoices WHERE created_at >= ?",
        (month_start,)
    ).fetchone()
    
    # Monthly history (last 6 months)
    monthly = conn.execute(
        "SELECT strftime('%Y-%m', created_at) as month, COUNT(*) as cnt, COALESCE(SUM(total_grand_total),0) as total FROM invoices GROUP BY month ORDER BY month DESC LIMIT 6"
    ).fetchall()
    
    # API usage this month
    usage = conn.execute(
        "SELECT ai_model, COUNT(*) as cnt, COALESCE(SUM(image_count),0) as imgs FROM api_usage WHERE created_at >= ? GROUP BY ai_model",
        (month_start,)
    ).fetchall()
    
    conn.close()
    
    # Estimate costs
    usage_list = []
    for u in usage:
        model = u["ai_model"]
        imgs = u["imgs"]
        if model == "gemini":
            cost = round(imgs * 0.1, 1)  # ~0.1 yen per image
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
    rows = conn.execute("SELECT id, name, discount_rate FROM customers ORDER BY id").fetchall()
    conn.close()
    return [dict(r) for r in rows]

@app.post("/api/customers")
async def add_customer(username: Annotated[str, Depends(authenticate)], name: str = Form(...), discount_rate: int = Form(35)):
    conn = get_db()
    conn.execute("INSERT INTO customers (name, discount_rate) VALUES (?, ?)", (name, discount_rate))
    conn.commit()
    conn.close()
    return {"status": "ok"}

@app.delete("/api/customers/{cid}")
async def delete_customer(cid: int, username: Annotated[str, Depends(authenticate)]):
    conn = get_db()
    conn.execute("DELETE FROM customers WHERE id = ?", (cid,))
    conn.commit()
    conn.close()
    return {"status": "ok"}

# ==================== History API ====================
@app.get("/api/history")
async def get_history(username: Annotated[str, Depends(authenticate)]):
    conn = get_db()
    rows = conn.execute(
        "SELECT id, invoice_number, customer_name, total_grand_total, item_count, created_at FROM invoices ORDER BY id DESC"
    ).fetchall()
    conn.close()
    return [dict(r) for r in rows]

@app.get("/api/history/{inv_id}/pdf")
async def download_history_pdf(inv_id: int, username: Annotated[str, Depends(authenticate)]):
    conn = get_db()
    row = conn.execute("SELECT pdf_data, invoice_number FROM invoices WHERE id = ?", (inv_id,)).fetchone()
    conn.close()
    if not row or not row["pdf_data"]:
        raise HTTPException(status_code=404, detail="PDF not found")
    return Response(content=row["pdf_data"], media_type="application/pdf",
                    headers={"Content-Disposition": f'attachment; filename="{row["invoice_number"]}.pdf"'})

@app.get("/api/history/{inv_id}/excel")
async def download_history_excel(inv_id: int, username: Annotated[str, Depends(authenticate)]):
    conn = get_db()
    row = conn.execute("SELECT excel_data, invoice_number FROM invoices WHERE id = ?", (inv_id,)).fetchone()
    conn.close()
    if not row or not row["excel_data"]:
        raise HTTPException(status_code=404, detail="Excel not found")
    return Response(content=row["excel_data"],
                    media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    headers={"Content-Disposition": f'attachment; filename="{row["invoice_number"]}.xlsx"'})

@app.get("/api/history/{inv_id}/items")
async def get_history_items(inv_id: int, username: Annotated[str, Depends(authenticate)]):
    conn = get_db()
    rows = conn.execute("SELECT code, color, size, unit_price, quantity FROM invoice_items WHERE invoice_id = ?", (inv_id,)).fetchall()
    inv = conn.execute("SELECT invoice_number, customer_name, discount_rate FROM invoices WHERE id = ?", (inv_id,)).fetchone()
    conn.close()
    if not inv:
        raise HTTPException(status_code=404, detail="Invoice not found")
    return {"invoice_number": inv["invoice_number"], "customer_name": inv["customer_name"], "discount_rate": inv["discount_rate"], "items": [dict(r) for r in rows]}

@app.delete("/api/history/{inv_id}")
async def delete_history(inv_id: int, username: Annotated[str, Depends(authenticate)]):
    conn = get_db()
    conn.execute("DELETE FROM invoice_items WHERE invoice_id = ?", (inv_id,))
    conn.execute("DELETE FROM invoices WHERE id = ?", (inv_id,))
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
            conn.execute("INSERT INTO api_usage (ai_model, image_count) VALUES (?, ?)", (ai_model, len(files)))
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

# ==================== Document Generation API ====================
class DocumentRequest(BaseModel):
    invoice_id: Optional[int] = None
    customer_name: str
    discount_rate: int
    items: List[Dict[str, Any]]

@app.post("/generate-documents")
async def generate_documents(
    request: Request,
    username: Annotated[str, Depends(authenticate)],
    payload: DocumentRequest
):
    try:
        items_data = payload.items
        customer_name = payload.customer_name
        discount_rate = payload.discount_rate
        
        conn = get_db()
        if payload.invoice_id:
            row = conn.execute("SELECT invoice_number FROM invoices WHERE id = ?", (payload.invoice_id,)).fetchone()
            if not row:
                conn.close()
                raise HTTPException(status_code=404, detail="Invoice not found for update")
            invoice_number = row["invoice_number"]
        else:
            invoice_number = generate_invoice_number()
        conn.close()

        processed_items = []
        total_net_amount = 0
        total_tax_amount = 0
        total_grand_total = 0

        for data in items_data:
            unit_price = data.get("unit_price", 0)
            if isinstance(unit_price, str):
                unit_price = int(unit_price.replace(',', '').replace('¥', '').strip() or '0')
            quantity = data.get("quantity", 1)
            if isinstance(quantity, str):
                quantity = int(quantity or '1')
            if quantity < 1: quantity = 1
            net_amount = int(unit_price * (discount_rate / 100) * quantity)
            tax_amount = int(net_amount * 0.1)
            grand_total = net_amount + tax_amount
            processed_items.append({
                "code": data.get("code", "-"), "color": data.get("color", "-"),
                "size": data.get("size", "-"), "unit_price": unit_price,
                "quantity": quantity,
                "net_amount": net_amount, "tax_amount": tax_amount, "grand_total": grand_total
            })
            total_net_amount += net_amount
            total_tax_amount += tax_amount
            total_grand_total += grand_total

        invoice_data = {
            "invoice_number": invoice_number, "customer_name": customer_name,
            "discount_rate": discount_rate, "items": processed_items,
            "total_net_amount": total_net_amount, "total_tax_amount": total_tax_amount,
            "total_grand_total": total_grand_total, "issuer": "株式会社 ラウラジャパン",
            "date": get_jst_now().strftime("%Y年%m月%d日")
        }

        # Generate PDF
        html_content = templates.get_template("invoice_template.html").render(invoice_data)
        pdf_bytes = HTML(string=html_content).write_pdf()

        # Generate Excel
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "伝票"
        fill_yellow = PatternFill(start_color="FFE699", end_color="FFE699", fill_type="solid")
        fill_blue = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        border_thin = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

        ws["A1"] = "伝票"
        ws["A1"].font = Font(bold=True, size=14)
        ws["B2"] = "1売上"
        ws["A4"] = "コード"
        ws["A5"] = "SZ2006"
        ws["A5"].fill = fill_yellow
        ws["C4"] = "店名"
        ws["C5"] = customer_name
        ws["C5"].fill = fill_yellow
        ws["F1"] = "No."
        ws["G1"] = invoice_number
        ws["G1"].fill = fill_yellow
        now = get_jst_now()
        reiwa_year = now.year - 2018
        ws["F3"] = "日付"
        ws["F3"].fill = fill_yellow
        ws["G3"] = f"{reiwa_year}年"
        ws["G3"].fill = fill_yellow
        ws["H3"] = f"{now.month}月"
        ws["H3"].fill = fill_yellow
        ws["I3"] = f"{now.day}日"
        ws["I3"].fill = fill_yellow

        start_row = 8
        for col, text in {"A": "品番", "B": "カラー", "C": "サイズ", "D": "バーコード", "F": "数量", "G": "単価", "H": "金額", "I": "掛率"}.items():
            cell = ws[f"{col}{start_row}"]
            cell.value = text
            cell.border = border_thin
            cell.alignment = Alignment(horizontal="center")

        ws.column_dimensions['A'].width = 15
        ws.column_dimensions['B'].width = 10
        ws.column_dimensions['C'].width = 10
        ws.column_dimensions['D'].width = 30
        ws.column_dimensions['E'].width = 2
        ws.column_dimensions['F'].width = 10
        ws.column_dimensions['G'].width = 15
        ws.column_dimensions['H'].width = 15
        ws.column_dimensions['I'].width = 10

        current_row = start_row + 1
        for i, item in enumerate(processed_items):
            ws.row_dimensions[current_row].height = 48
            ws[f"A{current_row}"] = item['code']
            ws[f"B{current_row}"] = item["color"]
            ws[f"C{current_row}"] = item["size"]
            ws[f"D{current_row}"] = ""
            ws[f"F{current_row}"] = item["quantity"]
            ws[f"G{current_row}"] = item["unit_price"]
            ws[f"G{current_row}"].number_format = '#,##0'
            ws[f"H{current_row}"] = item["net_amount"]
            ws[f"H{current_row}"].number_format = '#,##0'
            ws[f"I{current_row}"] = f"{discount_rate}%"
            for col in ['A', 'B', 'C', 'D', 'F', 'G', 'H', 'I']:
                cell = ws[f"{col}{current_row}"]
                cell.border = border_thin
                cell.alignment = Alignment(vertical="center", horizontal="center" if col not in ["A", "D"] else "left")
                if col in ['A', 'B', 'C', 'F', 'G', 'I']: cell.fill = fill_yellow
                if col == 'H': cell.fill = fill_blue
            
            # Generate Barcode Image
            bc_str = f"{item['code']}{item['color']}{item['size']}".replace('-', '').upper()
            bc_str = re.sub(r'[^A-Z0-9\-\.\ \$\/\+\%]', '', bc_str)
            if not bc_str: bc_str = "000"
            try:
                code39 = barcode.get_barcode_class('code39')
                bc_io = BytesIO()
                code39(bc_str, writer=ImageWriter(), add_checksum=False).write(
                    bc_io,
                    options={"module_width":0.22, "module_height":6.0, "font_size":7, "text_distance":2.5, "quiet_zone":1}
                )
                bc_io.seek(0)
                img = ExcelImage(bc_io)
                img_w_px, img_h_px = 175, 44

                # D列(0-indexed=3) の current_row(0-indexed=current_row-1) にオフセット付きで配置
                img.anchor = OneCellAnchor(
                    _from=AnchorMarker(
                        col=3, colOff=pixels_to_EMU(10),
                        row=current_row - 1, rowOff=pixels_to_EMU(6)
                    ),
                    ext=XDRPositiveSize2D(
                        cx=pixels_to_EMU(img_w_px),
                        cy=pixels_to_EMU(img_h_px)
                    )
                )
                ws.add_image(img)
            except Exception as e:
                print(f"Barcode gen error: {e}")
                
            current_row += 1

        excel_io = BytesIO()
        wb.save(excel_io)
        excel_bytes = excel_io.getvalue()

        # Save to SQLite
        conn = get_db()
        if payload.invoice_id:
            conn.execute(
                "UPDATE invoices SET customer_name=?, discount_rate=?, total_net_amount=?, total_tax_amount=?, total_grand_total=?, item_count=?, pdf_data=?, excel_data=?, created_at=CURRENT_TIMESTAMP WHERE id=?",
                (customer_name, discount_rate, total_net_amount, total_tax_amount, total_grand_total, len(processed_items), pdf_bytes, excel_bytes, payload.invoice_id)
            )
            conn.execute("DELETE FROM invoice_items WHERE invoice_id=?", (payload.invoice_id,))
            inv_id = payload.invoice_id
        else:
            cursor = conn.execute(
                "INSERT INTO invoices (invoice_number, customer_name, discount_rate, total_net_amount, total_tax_amount, total_grand_total, item_count, pdf_data, excel_data) VALUES (?,?,?,?,?,?,?,?,?)",
                (invoice_number, customer_name, discount_rate, total_net_amount, total_tax_amount, total_grand_total, len(processed_items), pdf_bytes, excel_bytes)
            )
            inv_id = cursor.lastrowid
            
        for item in processed_items:
            conn.execute("INSERT INTO invoice_items (invoice_id, code, color, size, unit_price, quantity, net_amount) VALUES (?,?,?,?,?,?,?)",
                         (inv_id, item["code"], item["color"], item["size"], item["unit_price"], item["quantity"], item["net_amount"]))
        conn.commit()
        conn.close()

        pdf_b64 = base64.b64encode(pdf_bytes).decode('utf-8')
        excel_b64 = base64.b64encode(excel_bytes).decode('utf-8')

        return JSONResponse({
            "invoice_id": inv_id,
            "invoice_number": invoice_number,
            "pdf_base64": pdf_b64,
            "excel_base64": excel_b64
        })
    except Exception as e:
        import traceback
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/api/history/{inv_id}/upload-drive")
async def upload_to_drive(inv_id: int, username: Annotated[str, Depends(authenticate)]):
    if not GDRIVE_WEBHOOK_URL:
        raise HTTPException(status_code=500, detail="GDRIVE_WEBHOOK_URLが未設定です。Renderの設定を確認してください。")

    conn = get_db()
    row = conn.execute(
        "SELECT invoice_number, customer_name, pdf_data, excel_data FROM invoices WHERE id = ?",
        (inv_id,)
    ).fetchone()
    conn.close()
    if not row:
        raise HTTPException(status_code=404, detail="Invoice not found")

    inv_num = row["invoice_number"]
    cust = (row["customer_name"] or "").replace("/", "_").replace("\\", "_")
    uploaded = []

    async def _upload_via_gas(filename, mime, data):
        import base64
        import requests
        payload = {
            "folderId": GDRIVE_FOLDER_ID,
            "filename": filename,
            "mime": mime,
            "base64": base64.b64encode(data).decode('utf-8')
        }
        resp = requests.post(GDRIVE_WEBHOOK_URL, json=payload, timeout=30)
        if resp.status_code != 200:
            raise Exception(f"GAS error: {resp.status_code} - {resp.text}")
        return resp.json()

    try:
        if row["pdf_data"]:
            res = await _upload_via_gas(f"{inv_num}_{cust}.pdf", "application/pdf", row["pdf_data"])
            uploaded.append({"type": "pdf", "name": f"{inv_num}_{cust}.pdf", "url": res.get("url")})
        if row["excel_data"]:
            res = await _upload_via_gas(
                f"{inv_num}_{cust}.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                row["excel_data"]
            )
            uploaded.append({"type": "excel", "name": f"{inv_num}_{cust}.xlsx", "url": res.get("url")})

        return {"status": "ok", "uploaded": uploaded}
    except Exception as e:
        import traceback; traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"Drive保存失敗(GAS): {e}")

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
