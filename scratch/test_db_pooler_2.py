import psycopg2
import sys

# Test Supavisor format
db_url = "postgresql://postgres.drdfnaompatanosmiwqb:Laura713611021102@aws-0-ap-northeast-1.pooler.supabase.com:6543/postgres"

try:
    conn = psycopg2.connect(db_url)
    print("Connection to pooler (Tokyo) successful!")
    conn.close()
    sys.exit(0)
except Exception as e:
    pass

# Try with db.domain format but port 6543 (Supavisor IPv4)
db_url2 = "postgresql://postgres.drdfnaompatanosmiwqb:Laura713611021102@db.drdfnaompatanosmiwqb.supabase.co:6543/postgres"
try:
    conn = psycopg2.connect(db_url2)
    print("Connection to db domain port 6543 successful!")
    conn.close()
    sys.exit(0)
except Exception as e:
    print(f"Connection failed: {e}")
