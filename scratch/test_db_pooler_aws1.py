import psycopg2
import sys

db_url = "postgresql://postgres.drdfnaompatanosmiwqb:Laura713611021102@aws-1-ap-southeast-1.pooler.supabase.com:6543/postgres"

try:
    conn = psycopg2.connect(db_url)
    print("Connection to aws-1 pooler successful!")
    conn.close()
    sys.exit(0)
except Exception as e:
    print(f"Connection failed: {e}")
