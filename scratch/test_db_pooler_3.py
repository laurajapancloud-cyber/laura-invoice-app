import psycopg2
import sys

# Try with db.domain format but port 6543 (Supavisor IPv4) with user 'postgres'
db_url3 = "postgresql://postgres:Laura713611021102@db.drdfnaompatanosmiwqb.supabase.co:6543/postgres"
try:
    conn = psycopg2.connect(db_url3)
    print("Connection to db domain port 6543 with user 'postgres' successful!")
    conn.close()
    sys.exit(0)
except Exception as e:
    print(f"Connection failed: {e}")
