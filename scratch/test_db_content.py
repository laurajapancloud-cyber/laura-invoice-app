import psycopg2
import sys

db_url = "postgresql://postgres:Laura713611021102@db.drdfnaompatanosmiwqb.supabase.co:6543/postgres"

try:
    conn = psycopg2.connect(db_url)
    with conn.cursor() as cur:
        cur.execute("SELECT table_name FROM information_schema.tables WHERE table_schema='public'")
        tables = [r[0] for r in cur.fetchall()]
        print("Tables in public schema:", tables)
        
        # Test if invoices table has data
        if 'invoices' in tables:
            cur.execute("SELECT count(*) FROM invoices")
            print("Invoices count:", cur.fetchone()[0])
            
        # Test if users table has data
        if 'users' in tables:
            cur.execute("SELECT count(*) FROM users")
            print("Users count:", cur.fetchone()[0])
            
    conn.close()
    sys.exit(0)
except Exception as e:
    print(f"Connection/Query failed: {e}")
