import psycopg2
import sys

db_url = "postgresql://postgres:6BAKx#8j_qZRaW5@db.drdfnaompatanosmiwqb.supabase.co:5432/postgres"

try:
    conn = psycopg2.connect(db_url)
    print("Connection successful!")
    conn.close()
except Exception as e:
    print(f"Connection failed: {e}")
