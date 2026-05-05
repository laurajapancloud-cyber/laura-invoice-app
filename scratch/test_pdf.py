import os
import sys
import base64

# Add the current directory to sys.path
sys.path.append(os.getcwd())

from main import build_invoice_pdf, assemble_invoice_data

test_inv = {
    "invoice_number": "TEST-202605-0001",
    "customer_name": "テスト株式会社",
    "discount_rate": 35,
    "status": "draft",
    "doc_type": "delivery",
    "id": 1,
    "created_at": "2026-05-05T00:00:00Z"
}
test_items = [
    {"code": "ABC-123", "color": "RED", "size": "M", "unit_price": 1000, "quantity": 2}
]

try:
    inv_data = assemble_invoice_data(test_inv, test_items, 35, "delivery")
    pdf_bytes = build_invoice_pdf(inv_data)
    print(f"PDF generated successfully: {len(pdf_bytes)} bytes")
except Exception as e:
    import traceback
    print(f"PDF generation failed: {e}")
    traceback.print_exc()
