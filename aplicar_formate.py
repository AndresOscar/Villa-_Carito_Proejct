import sqlite3

DB_NAME = "configuraciones.db"

conn = sqlite3.connect(DB_NAME)
cursor = conn.cursor()

# Agregar columna money_columns si no existe
try:
    cursor.execute("ALTER TABLE configuraciones_tablas ADD COLUMN money_columns TEXT")
    print("Columna money_columns agregada.")
except sqlite3.OperationalError:
    print("La columna money_columns ya existe.")

# Agregar columna header_rows si no existe
try:
    cursor.execute("ALTER TABLE configuraciones_tablas ADD COLUMN header_rows INTEGER")
    print("Columna header_rows agregada.")
except sqlite3.OperationalError:
    print("La columna header_rows ya existe.")

conn.commit()
conn.close()
