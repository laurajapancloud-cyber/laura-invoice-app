import psycopg2
import sys

db_url = "postgresql://postgres.drdfnaompatanosmiwqb:Laura713611021102@aws-0-ap-southeast-1.pooler.supabase.com:6543/postgres"

try:
    conn = psycopg2.connect(db_url)
    print("Connection to pooler successful!")
    
    with conn.cursor() as cur:
        cur.execute("SELECT 1")
        print("Query successful:", cur.fetchone())
        
    conn.close()
except Exception as e:
    print(f"Connection failed: {e}")
