import sqlite3

conn = sqlite3.connect("database.db")
cursor = conn.cursor()

try:
    cursor.execute("ALTER TABLE unidades ADD COLUMN sigla TEXT")
    print("Coluna 'sigla' adicionada com sucesso.")
except sqlite3.OperationalError as e:
    print("Erro:", e)

conn.commit()
conn.close()
