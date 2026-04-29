import os
import json
import secrets
from typing import Annotated

from fastapi import FastAPI, UploadFile, File, Depends, HTTPException, status, Response, Request
from fastapi.security import HTTPBasic, HTTPBasicCredentials
from fastapi.templating import Jinja2Templates
from fastapi.responses import HTMLResponse
import google.generativeai as genai
from weasyprint import HTML
from dotenv import load_dotenv

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
model = genai.GenerativeModel("gemini-1.5-flash")

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
    return templates.TemplateResponse("index.html", {"request": request})

@app.post("/generate-pdf")
    try:
        # 1. Read Image
        image_bytes = await file.read()
        
        # 2. Gemini API Request
        prompt = """
あなたはアパレルブランドのデータ入力アシスタントです。提供された商品タグの画像から以下の情報を抽出し、厳密にJSON形式のみで出力してください。マークダウンの装飾(```jsonなど)は含めないでください。
スキーマ: {"code": "品番(例:148-3101)", "color": "カラー(例:24)", "size": "サイズ(例:46)", "unit_price": 単価の数値(例:38000)}
""".strip()

        response = model.generate_content([
            prompt,
            {"mime_type": "image/jpeg", "data": image_bytes}
        ])
        
        # Extract JSON from response
        try:
            raw_text = response.text.strip()
            # Clean up potential markdown formatting if Gemini ignored the instruction
            if raw_text.startswith("```json"):
                raw_text = raw_text.replace("```json", "", 1).replace("```", "", 1).strip()
            elif raw_text.startswith("```"):
                raw_text = raw_text.replace("```", "", 2).strip()
            
            data = json.loads(raw_text)
        except Exception as e:
            raise HTTPException(status_code=500, detail=f"Gemini output parsing failed: {str(e)}\nRaw: {response.text}")

        # 3. Calculation Logic
        unit_price = data.get("unit_price", 0)
        net_amount = int(unit_price * 0.35)
        tax_amount = int(net_amount * 0.1)
        grand_total = net_amount + tax_amount
        
        invoice_data = {
            "customer_name": "株式会社 タム 御中",
            "code": data.get("code", "-"),
            "color": data.get("color", "-"),
            "size": data.get("size", "-"),
            "unit_price": unit_price,
            "net_amount": net_amount,
            "tax_amount": tax_amount,
            "grand_total": grand_total,
            "issuer": "株式会社 ラウラジャパン"
        }

        # 4. Render HTML for PDF
        html_content = templates.get_template("invoice_template.html").render(invoice_data)

        # 5. Generate PDF with WeasyPrint
        pdf_bytes = HTML(string=html_content).write_pdf()

        # 6. Return Response
        return Response(
            content=pdf_bytes,
            media_type="application/pdf",
            headers={
                "Content-Disposition": "attachment; filename=LAURA_JAPAN_invoice.pdf"
            }
        )

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
