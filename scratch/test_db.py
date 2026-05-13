import os
import psycopg2
from dotenv import load_dotenv

load_dotenv()

db_url = os.getenv("DATABASE_URL")
print(f"DATABASE_URL: {db_url}")

if not db_url:
    print("DATABASE_URL is not set.")
    exit(1)

try:
    conn = psycopg2.connect(db_url)
    print("Successfully connected to the database.")
    cur = conn.cursor()
    cur.execute("SELECT version();")
    print(f"DB Version: {cur.fetchone()}")
    
    cur.execute("SELECT COUNT(*) FROM users;")
    print(f"Users count: {cur.fetchone()[0]}")
    
    cur.execute("SELECT COUNT(*) FROM customers;")
    print(f"Customers count: {cur.fetchone()[0]}")
    
    cur.execute("SELECT COUNT(*) FROM invoices;")
    print(f"Invoices count: {cur.fetchone()[0]}")
    
    conn.close()
except Exception as e:
    print(f"Failed to connect to the database: {e}")
