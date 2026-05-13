import psycopg2
import sys

db_url = "postgresql://postgres:6BAKx%238j_qZRaW5@db.drdfnaompatanosmiwqb.supabase.co:5432/postgres"

try:
    conn = psycopg2.connect(db_url)
    print("Connection successful!")
    
    with conn.cursor() as cur:
        cur.execute("SELECT table_name FROM information_schema.tables WHERE table_schema='public'")
        tables = [r[0] for r in cur.fetchall()]
        print("Tables in public schema:", tables)
        
    conn.close()
except Exception as e:
    print(f"Connection failed: {e}")
