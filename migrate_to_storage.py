"""既存の BYTEA データを Supabase Storage に流し込むワンショットスクリプト"""
import os
import base64
import psycopg2
from psycopg2.extras import RealDictCursor
from supabase import create_client
from dotenv import load_dotenv

# .env ファイルを読み込む（ローカル実行時）
load_dotenv()

DATABASE_URL = os.getenv("DATABASE_URL")
SUPABASE_URL = os.getenv("SUPABASE_URL")
SUPABASE_SERVICE_KEY = os.getenv("SUPABASE_SERVICE_KEY")
BUCKET = "invoices"

if not all([DATABASE_URL, SUPABASE_URL, SUPABASE_SERVICE_KEY]):
    print("Error: Missing environment variables (DATABASE_URL, SUPABASE_URL, or SUPABASE_SERVICE_KEY)")
    exit(1)

conn = psycopg2.connect(DATABASE_URL, cursor_factory=RealDictCursor)
sb = create_client(SUPABASE_URL, SUPABASE_SERVICE_KEY)

print("--- Migration Started ---")

with conn.cursor() as cur:
    # まだ Storage 移行されていない、データがある行を取得
    cur.execute("""
        SELECT id, invoice_number, pdf_data, excel_data, detail_pdf_data, detail_excel_data, created_at
        FROM invoices
        WHERE pdf_storage_path IS NULL AND pdf_data IS NOT NULL
        ORDER BY id
    """)
    rows = cur.fetchall()
    
    if not rows:
        print("No invoices to migrate.")
        conn.close()
        exit(0)

    print(f"Migrating {len(rows)} invoices to Supabase Storage...")

    for r in rows:
        inv_num = r['invoice_number']
        base_path = f"locked/{inv_num}"
        paths = {}
        
        # 4つのファイルターゲット
        targets = [
            ("pdf", "invoice.pdf", "application/pdf", r["pdf_data"]),
            ("excel", "invoice.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", r["excel_data"]),
            ("detail_pdf", "detail.pdf", "application/pdf", r["detail_pdf_data"]),
            ("detail_excel", "detail.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", r["detail_excel_data"]),
        ]
        
        for typ, filename, mime, data in targets:
            if not data:
                continue
            
            full_path = f"{base_path}/{filename}"
            print(f"  Uploading {full_path} ...", end="")
            
            # 既存があれば念のため削除してからアップロード（upsert）
            try:
                sb.storage.from_(BUCKET).remove([full_path])
            except:
                pass
            
            try:
                sb.storage.from_(BUCKET).upload(
                    full_path, 
                    bytes(data), 
                    file_options={"content-type": mime, "upsert": "true"}
                )
                paths[typ] = full_path
                print(" OK")
            except Exception as e:
                print(f" FAIL: {e}")

        # DBのパス情報を更新し、ステータスを locked にする
        with conn.cursor() as cur2:
            cur2.execute("""
                UPDATE invoices SET 
                    status='locked', 
                    locked_at=COALESCE(locked_at, created_at),
                    pdf_storage_path=%s, 
                    excel_storage_path=%s, 
                    detail_pdf_storage_path=%s, 
                    detail_excel_storage_path=%s
                WHERE id=%s
            """, (
                paths.get("pdf"), 
                paths.get("excel"), 
                paths.get("detail_pdf"), 
                paths.get("detail_excel"), 
                r["id"]
            ))
            conn.commit()

print("\n--- Migration Completed ---")
conn.close()
