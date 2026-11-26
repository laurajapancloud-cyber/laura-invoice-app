from main import build_invoice_excel

invoice_data = {
    "doc_type": "delivery",
    "customer_code": "AM2005",
    "customer_name": "アクティブY 様",
    "date": "2026年5月27日",
    "user_name": "テスト担当",
    "discount_rate": 35,
    "total_net_amount": 100000,
    "total_tax_amount": 10000,
    "total_grand_total": 110000,
    "items": []
}

for i in range(45):
    invoice_data["items"].append({
        "code": f"ITEM-{i+1:03d}",
        "color": "BLK",
        "size": "M",
        "quantity": 2,
        "unit_price": 5000,
        "net_amount": 10000
    })

try:
    excel_bytes = build_invoice_excel(invoice_data, is_preview=False)
    with open("scratch/test_excel.xlsx", "wb") as f:
        f.write(excel_bytes)
    print("Excel generation SUCCESS")
except Exception as e:
    import traceback
    traceback.print_exc()
    print("Excel generation FAILED")
