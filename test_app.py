from utils.db import get_db_connection

conn = get_db_connection()
cursor = conn.cursor()
cursor.execute("SELECT DB_NAME()")
print(cursor.fetchone())
cursor.close()
conn.close()