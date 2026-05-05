import os
import sys
import base64

# Add the current directory to sys.path
sys.path.append(os.getcwd())

from main import build_all_files, assemble_invoice_data

test_inv = {
    "id": 39,
    "invoice_number": "TEST-39",
    "customer_name": "テスト株式会社",
    "discount_rate": 35,
    "status": "draft",
    "doc_type": "delivery",
    "created_at": "2026-05-05T00:00:00Z"
}
test_items = [
    {"code": "ABC-123", "color": "RED", "size": "44", "unit_price": 1000, "quantity": 2}
]

try:
    print("Assembling data...")
    inv_data = assemble_invoice_data(test_inv, test_items, 35, "delivery")
    print("Building all files...")
    files = build_all_files(inv_data)
    print(f"Success! PDF size: {len(files['pdf'])}, Excel size: {len(files['excel'])}")
except Exception as e:
    import traceback
    print(f"FAILED: {e}")
    traceback.print_exc()
