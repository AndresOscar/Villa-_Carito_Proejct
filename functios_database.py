import sqlite3


DB_NAME = "configuraciones.db"

def inicializar_base_datos():
    conn = sqlite3.connect('configuraciones.db')
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS configuraciones_tablas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            excel_file TEXT NOT NULL,
            sheet_name TEXT NOT NULL,
            excel_range TEXT NOT NULL,
            word_file TEXT NOT NULL,
            table_label TEXT NOT NULL,
            output_file TEXT,
            money_columns TEXT,      -- nuevo campo: ej. "1,2"
            header_rows INTEGER      -- nuevo campo: ej. 1
        )
    ''')
    conn.commit()
    conn.close()

def eliminar_configuracion(id_config):
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute('DELETE FROM configuraciones_tablas WHERE id=?', (id_config,))
    conn.commit()
    conn.close()


def guardar_configuracion(excel_file, sheet_name, excel_range, word_file, table_label, output_file=None, money_columns=None, header_rows=None):
    conn = sqlite3.connect('configuraciones.db')
    cursor = conn.cursor()
    cursor.execute('''
        INSERT INTO configuraciones_tablas (excel_file, sheet_name, excel_range, word_file, table_label, output_file, money_columns, header_rows)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    ''', (excel_file, sheet_name, excel_range, word_file, table_label, output_file, money_columns, header_rows))
    conn.commit()
    conn.close()


def obtener_configuraciones():
    conn = sqlite3.connect('configuraciones.db')
    cursor = conn.cursor()
    cursor.execute('SELECT * FROM configuraciones_tablas')
    configs = cursor.fetchall()
    conn.close()
    return configs


def actualizar_configuracion(id_config, excel_file, sheet_name, excel_range, word_file, table_label, output_file=None, money_columns=None, header_rows=None):
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute('''
        UPDATE configuraciones_tablas
        SET excel_file=?, sheet_name=?, excel_range=?, word_file=?, table_label=?, output_file=?, money_columns=?, header_rows=?
        WHERE id=?
    ''', (excel_file, sheet_name, excel_range, word_file, table_label, output_file, money_columns, header_rows, id_config))
    conn.commit()
    conn.close()



