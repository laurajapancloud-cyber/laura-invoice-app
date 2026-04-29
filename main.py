import os
import json
import secrets
from typing import Annotated, List
import base64
from io import BytesIO
import datetime

from fastapi import FastAPI, UploadFile, File, Depends, HTTPException, status, Response, Request, Form
from fastapi.security import HTTPBasic, HTTPBasicCredentials
from fastapi.templating import Jinja2Templates
from fastapi.responses import HTMLResponse, JSONResponse
import google.generativeai as genai
from weasyprint import HTML
from dotenv import load_dotenv
import openpyxl
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font

# Load environment variables
load_dotenv()

# Configuration
APP_USERNAME = os.getenv("APP_USERNAME")
APP_PASSWORD = os.getenv("APP_PASSWORD")
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")

if not all([APP_USERNAME, APP_PASSWORD, GEMINI_API_KEY]):
    print("Warning: Missing required environment variables.")

# Gemini Initialization
genai.configure(api_key=GEMINI_API_KEY)
model = genai.GenerativeModel("gemini-2.5-flash")

# FastAPI App
app = FastAPI()
security = HTTPBasic()
templates = Jinja2Templates(directory="templates")

def authenticate(credentials: Annotated[HTTPBasicCredentials, Depends(security)]):
    if not APP_USERNAME or not APP_PASSWORD:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail="Server configuration error: Auth credentials not set."
        )
    
    current_username_bytes = credentials.username.encode("utf8")
    correct_username_bytes = APP_USERNAME.encode("utf8")
    is_correct_username = secrets.compare_digest(
        current_username_bytes, correct_username_bytes
    )
    current_password_bytes = credentials.password.encode("utf8")
    correct_password_bytes = APP_PASSWORD.encode("utf8")
    is_correct_password = secrets.compare_digest(
        current_password_bytes, correct_password_bytes
    )
    if not (is_correct_username and is_correct_password):
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Incorrect email or password",
            headers={"WWW-Authenticate": "Basic"},
        )
    return credentials.username

@app.get("/", response_class=HTMLResponse)
async def index(request: Request, username: Annotated[str, Depends(authenticate)]):
    return templates.TemplateResponse(request=request, name="index.html")

@app.post("/generate-documents")
async def generate_documents(
    request: Request,
    username: Annotated[str, Depends(authenticate)],
    files: List[UploadFile] = File(...),
    customer_name: str = Form("株式会社 タム 御中"),
    discount_rate: int = Form(35)
):
    try:
        # 1. Read all Images
        image_parts = []
        for file in files:
            image_bytes = await file.read()
            image_parts.append({"mime_type": file.content_type or "image/jpeg", "data": image_bytes})
        
        # 2. Gemini API Request
        prompt = """
あなたはアパレルブランドのデータ入力アシスタントです。提供された複数の商品タグの画像からそれぞれの情報を抽出し、厳密にJSON配列（リスト）の形式のみで出力してください。マークダウンの装飾(```jsonなど)は含めないでください。
スキーマ: [{"code": "品番(例:148-3101)", "color": "カラー(例:24)", "size": "サイズ(例:46)", "unit_price": 単価の数値(例:38000)}, ...]
複数の商品がある場合は、配列内に複数のオブジェクトを含めてください。
""".strip()

        contents = [prompt] + image_parts
        
        try:
            response = model.generate_content(contents)
        except Exception as e:
            error_str = str(e)
            if "429" in error_str or "Quota exceeded" in error_str or "rate limits" in error_str:
                raise HTTPException(status_code=429, detail="AI（Gemini）の一時的な利用制限に達しました。しばらく（約30秒）待ってから再度お試しください。")
            raise HTTPException(status_code=500, detail=f"Gemini API Request Error: {error_str}")
        
        # Extract JSON from response
        try:
            raw_text = response.text.strip()
            if raw_text.startswith("```json"):
                raw_text = raw_text.replace("```json", "", 1).replace("```", "", 1).strip()
            elif raw_text.startswith("```"):
                raw_text = raw_text.replace("```", "", 2).strip()
            
            items_data = json.loads(raw_text)
            if not isinstance(items_data, list):
                items_data = [items_data]
        except Exception as e:
            raise HTTPException(status_code=500, detail=f"Gemini output parsing failed: {str(e)}\nRaw: {response.text}")

        # 3. Calculation Logic for all items
        processed_items = []
        total_net_amount = 0
        total_tax_amount = 0
        total_grand_total = 0
        
        for data in items_data:
            unit_price = data.get("unit_price", 0)
            if isinstance(unit_price, str):
                unit_price = int(unit_price.replace(',', '').replace('¥', '').strip())
            net_amount = int(unit_price * (discount_rate / 100))
            tax_amount = int(net_amount * 0.1)
            grand_total = net_amount + tax_amount
            
            processed_items.append({
                "code": data.get("code", "-"),
                "color": data.get("color", "-"),
                "size": data.get("size", "-"),
                "unit_price": unit_price,
                "net_amount": net_amount,
                "tax_amount": tax_amount,
                "grand_total": grand_total
            })
            
            total_net_amount += net_amount
            total_tax_amount += tax_amount
            total_grand_total += grand_total
        
        invoice_data = {
            "customer_name": customer_name,
            "discount_rate": discount_rate,
            "items": processed_items,
            "total_net_amount": total_net_amount,
            "total_tax_amount": total_tax_amount,
            "total_grand_total": total_grand_total,
            "issuer": "株式会社 ラウラジャパン",
            "date": datetime.datetime.now().strftime("%Y年%m月%d日")
        }

        # 4. Generate PDF with WeasyPrint
        html_content = templates.get_template("invoice_template.html").render(invoice_data)
        pdf_bytes = HTML(string=html_content).write_pdf()

        # 5. Generate Excel with openpyxl
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "伝票"

        # Styles
        fill_yellow = PatternFill(start_color="FFE699", end_color="FFE699", fill_type="solid")
        fill_blue = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        border_thin = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

        # Header values
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
        ws["G1"].fill = fill_yellow
        
        now = datetime.datetime.now()
        reiwa_year = now.year - 2018
        ws["F3"] = "日付"
        ws["F3"].fill = fill_yellow
        ws["G3"] = f"{reiwa_year}年"
        ws["G3"].fill = fill_yellow
        ws["H3"] = f"{now.month}月"
        ws["H3"].fill = fill_yellow
        ws["I3"] = f"{now.day}日"
        ws["I3"].fill = fill_yellow

        # Table Headers
        start_row = 8
        header_mapping = {
            "A": "品番", "B": "カラー", "C": "サイズ", "D": "バーコード",
            "F": "数量", "G": "単価", "H": "金額", "I": "掛率"
        }
        
        for col, text in header_mapping.items():
            cell = ws[f"{col}{start_row}"]
            cell.value = text
            cell.border = border_thin
            cell.alignment = Alignment(horizontal="center")
            
        # Adjust column widths
        ws.column_dimensions['A'].width = 15
        ws.column_dimensions['B'].width = 10
        ws.column_dimensions['C'].width = 10
        ws.column_dimensions['D'].width = 30
        ws.column_dimensions['E'].width = 2
        ws.column_dimensions['F'].width = 10
        ws.column_dimensions['G'].width = 15
        ws.column_dimensions['H'].width = 15
        ws.column_dimensions['I'].width = 10

        # Data Rows
        current_row = start_row + 1
        for i, item in enumerate(processed_items):
            ws.row_dimensions[current_row].height = 40
            
            # 枝番 (Line number on the left)
            ws[f"A{current_row}"] = f"{i+1} {item['code']}"
            ws[f"B{current_row}"] = item["color"]
            ws[f"C{current_row}"] = item["size"]
            ws[f"D{current_row}"] = item["code"]
            ws[f"F{current_row}"] = 1
            ws[f"G{current_row}"] = item["unit_price"]
            ws[f"G{current_row}"].number_format = '#,##0'
            ws[f"H{current_row}"] = item["net_amount"]
            ws[f"H{current_row}"].number_format = '#,##0'
            ws[f"I{current_row}"] = f"{discount_rate}%"
            
            for col in ['A', 'B', 'C', 'D', 'F', 'G', 'H', 'I']:
                cell = ws[f"{col}{current_row}"]
                cell.border = border_thin
                cell.alignment = Alignment(vertical="center", horizontal="center" if col not in ["A", "D"] else "left")
                
                if col in ['A', 'B', 'C', 'F', 'G', 'I']:
                    cell.fill = fill_yellow
                if col == 'H':
                    cell.fill = fill_blue
                
                if col == 'D':
                    cell.font = Font(name="Code39", size=24)
            
            current_row += 1

        # Save Excel to BytesIO
        excel_io = BytesIO()
        wb.save(excel_io)
        excel_bytes = excel_io.getvalue()

        # 6. Return Response
        pdf_b64 = base64.b64encode(pdf_bytes).decode('utf-8')
        excel_b64 = base64.b64encode(excel_bytes).decode('utf-8')

        return JSONResponse({
            "pdf_filename": "LAURA_JAPAN_invoice.pdf",
            "pdf_base64": pdf_b64,
            "excel_filename": "LAURA_JAPAN_data.xlsx",
            "excel_base64": excel_b64
        })

    except Exception as e:
        import traceback
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(e))

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
