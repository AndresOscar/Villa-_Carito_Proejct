from Functions_Backs import update_word_table_from_excel,main
from functios_database import inicializar_base_datos,guardar_configuracion,actualizar_configuracion,obtener_configuraciones,eliminar_configuracion

import sys
import sqlite3
from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QLineEdit, QPushButton, QFileDialog,
    QVBoxLayout, QHBoxLayout, QMessageBox, QGroupBox, QFormLayout, QTableWidget,
    QTableWidgetItem, QHeaderView
)
from PyQt5.QtGui import QFont
from PyQt5.QtCore import Qt



def actualizar_todas_las_tablas():
    configs = obtener_configuraciones()
    for config in configs:
        id_, excel_file, sheet_name, excel_range, word_file, table_label, output_file, money_columns_str, header_rows = config
        try:
            print(f"Actualizando tabla ID {id_}...")
            update_word_table_from_excel(excel_file, sheet_name, excel_range, word_file, table_label, output_file)
            print(f"Tabla ID {id_} actualizada correctamente.")

            # money_columns_str viene de la base de datos (ej. "1,2,3")
            if money_columns_str:
                money_columns = [int(x.strip()) - 1 for x in money_columns_str.split(",") if x.strip().isdigit()]
            else:
                money_columns = []

                
            header_rows = int(header_rows) if header_rows else 1

            # Aplica formato solo si hay columnas de dinero definidas
            target_doc = output_file if output_file else word_file


            if money_columns:
                main(target_doc, table_label, money_columns, header_rows)
                print(f"Formato aplicado a tabla ID {id_}.")

        except Exception as e:
            print(f"Error al actualizar tabla ID {id_}: {e}")


# Clase interfaz
class TableUpdaterGUI(QWidget):
    def __init__(self):
        super().__init__()
        inicializar_base_datos()
        self.init_ui()
        self.id_configuraciones = []
        self.cargar_configuraciones()
        


    def init_ui(self):
        self.setWindowTitle("Actualizador de Tablas Word-Excel")
        self.setStyleSheet("""
            QLabel { font-size: 14px; }
            QLineEdit { padding: 5px; font-size: 13px; }
            QPushButton { padding: 8px 15px; font-size: 13px; }
        """)

        title = QLabel("Actualizador de Tablas Word-Excel")
        title.setFont(QFont("Arial", 18, QFont.Bold))
        title.setAlignment(Qt.AlignCenter)

        # Grupo Excel
        excel_group = QGroupBox("Archivo de Excel")
        excel_layout = QFormLayout()

        self.excel_input = QLineEdit()
        excel_btn = QPushButton("Examinar")
        excel_btn.clicked.connect(self.browse_excel)
        excel_row = QHBoxLayout()
        excel_row.addWidget(self.excel_input)
        excel_row.addWidget(excel_btn)
        excel_layout.addRow("Ruta:", excel_row)

        self.sheet_input = QLineEdit()
        excel_layout.addRow("Hoja:", self.sheet_input)

        self.range_input = QLineEdit()
        excel_layout.addRow("Rango (ej B2:L14):", self.range_input)
        excel_group.setLayout(excel_layout)

        # Grupo Word
        word_group = QGroupBox("Archivo de Word")
        word_layout = QFormLayout()

        self.word_input = QLineEdit()
        word_btn = QPushButton("Examinar")
        word_btn.clicked.connect(self.browse_word)
        word_row = QHBoxLayout()
        word_row.addWidget(self.word_input)
        word_row.addWidget(word_btn)
        word_layout.addRow("Ruta:", word_row)

        self.label_input = QLineEdit()
        word_layout.addRow("Etiqueta:", self.label_input)
        word_group.setLayout(word_layout)

        # Grupo salida
        output_group = QGroupBox("Archivo de salida (opcional)")
        output_layout = QFormLayout()

        self.output_input = QLineEdit()
        output_btn = QPushButton("Examinar")
        output_btn.clicked.connect(self.browse_output)
        output_row = QHBoxLayout()
        output_row.addWidget(self.output_input)
        output_row.addWidget(output_btn)
        output_layout.addRow("Guardar como:", output_row)
        output_group.setLayout(output_layout)


                # Grupo formato adicional
        format_group = QGroupBox("Formato adicional")
        format_layout = QFormLayout()

        self.money_columns_input = QLineEdit()
        self.money_columns_input.setPlaceholderText("Ej. 1,2,3")
        format_layout.addRow("Columnas dinero (índices separados por coma):", self.money_columns_input)

        self.header_rows_input = QLineEdit()
        self.header_rows_input.setPlaceholderText("Ej. 1")
        format_layout.addRow("Filas de encabezado:", self.header_rows_input)

        format_group.setLayout(format_layout)


        # Botones acción
        guardar_btn = QPushButton("Guardar Configuración")
        guardar_btn.clicked.connect(self.guardar_config)

        actualizar_btn = QPushButton("Actualizar TODAS las tablas")
        actualizar_btn.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold;")
        actualizar_btn.clicked.connect(self.actualizar_todas)

        editar_btn = QPushButton("Actualizar Configuración Seleccionada")

        editar_btn.clicked.connect(self.actualizar_configuracion_seleccionada)

        eliminar_btn = QPushButton("Eliminar Configuración Seleccionada")
        eliminar_btn.setStyleSheet("background-color: #E53935; color: white; font-weight: bold;")
        eliminar_btn.clicked.connect(self.eliminar_configuracion_seleccionada)



        # Tabla configuraciones
        self.config_table = QTableWidget()
        self.config_table.setColumnCount(8)
        self.config_table.setHorizontalHeaderLabels(["Excel", "Hoja", "Rango", "Word", "Etiqueta", "Salida", "Columnas dinero", "n filas encabezado"])

        self.config_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)

        # Layout principal
        layout = QVBoxLayout()
        layout.addWidget(title)
        layout.addWidget(excel_group)
        layout.addWidget(word_group)
        layout.addWidget(output_group)
        layout.addWidget(format_group)
        layout.addWidget(guardar_btn)
        layout.addWidget(editar_btn)
        layout.addWidget(QLabel("Configuraciones guardadas:"))
        layout.addWidget(self.config_table)
        layout.addWidget(eliminar_btn)
        layout.addWidget(actualizar_btn)
        

     


        self.setLayout(layout)
        self.resize(700, 800)  # da un tamaño inicial, pero permite que se ajuste


    def browse_excel(self):
        file, _ = QFileDialog.getOpenFileName(self, "Seleccionar archivo Excel", "", "Archivos Excel (*.xlsx *.xls)")
        if file:
            self.excel_input.setText(file)

    def browse_word(self):
        file, _ = QFileDialog.getOpenFileName(self, "Seleccionar archivo Word", "", "Archivos Word (*.docx)")
        if file:
            self.word_input.setText(file)


    def eliminar_configuracion_seleccionada(self):
        selected_row = self.config_table.currentRow()
        if selected_row < 0:
            QMessageBox.warning(self, "Seleccionar", "Selecciona una configuración para eliminar.")
            return

        respuesta = QMessageBox.question(self, "Confirmar eliminación", "¿Estás seguro de que deseas eliminar esta configuración?", QMessageBox.Yes | QMessageBox.No)
        if respuesta == QMessageBox.No:
            return

        id_config = self.id_configuraciones[selected_row]
        eliminar_configuracion(id_config)
        QMessageBox.information(self, "Eliminado", "Configuración eliminada correctamente.")
        self.cargar_configuraciones()


    def browse_output(self):
        file, _ = QFileDialog.getSaveFileName(self, "Guardar como", "", "Archivos Word (*.docx)")
        if file:
            self.output_input.setText(file)

    def guardar_config(self):
        excel_file = self.excel_input.text().strip()
        sheet_name = self.sheet_input.text().strip()
        excel_range = self.range_input.text().strip()
        word_file = self.word_input.text().strip()
        table_label = self.label_input.text().strip()
        output_file = self.output_input.text().strip() or None
        money_columns = self.money_columns_input.text().strip() or None
        header_rows = self.header_rows_input.text().strip() or None



        if not all([excel_file, sheet_name, excel_range, word_file, table_label]):
            QMessageBox.warning(self, "Faltan datos", "Completa todos los campos obligatorios.")
            return

        guardar_configuracion(excel_file, sheet_name, excel_range, word_file, table_label, output_file, money_columns, header_rows)
        QMessageBox.information(self, "Guardado", "Configuración guardada correctamente.")
        self.cargar_configuraciones()


    def cargar_fila_en_campos(self, row_idx):
        self.excel_input.setText(self.config_table.item(row_idx, 0).text())
        self.sheet_input.setText(self.config_table.item(row_idx, 1).text())
        self.range_input.setText(self.config_table.item(row_idx, 2).text())
        self.word_input.setText(self.config_table.item(row_idx, 3).text())
        self.label_input.setText(self.config_table.item(row_idx, 4).text())
        self.output_input.setText(self.config_table.item(row_idx, 5).text())
        self.money_columns_input.setText(self.config_table.item(row_idx, 6).text())
        self.header_rows_input.setText(self.config_table.item(row_idx, 7).text() )

    def actualizar_configuracion_seleccionada(self):
        selected_row = self.config_table.currentRow()
        if selected_row < 0:
            QMessageBox.warning(self, "Seleccionar", "Selecciona una configuración para actualizar.")
            return

        id_config = self.id_configuraciones[selected_row]
        excel_file = self.excel_input.text().strip()
        sheet_name = self.sheet_input.text().strip()
        excel_range = self.range_input.text().strip()
        word_file = self.word_input.text().strip()
        table_label = self.label_input.text().strip()
        output_file = self.output_input.text().strip() or None
        money_columns = self.money_columns_input.text().strip() or None
        header_rows = self.header_rows_input.text().strip() or None

        if not all([excel_file, sheet_name, excel_range, word_file, table_label]):
            QMessageBox.warning(self, "Faltan datos", "Completa todos los campos obligatorios.")
            return

        actualizar_configuracion(id_config, excel_file, sheet_name, excel_range, word_file, table_label, output_file, money_columns, header_rows)
        QMessageBox.information(self, "Actualizado", "Configuración actualizada correctamente.")
        self.cargar_configuraciones()



    def cargar_configuraciones(self):
        configs = obtener_configuraciones()

        self.id_configuraciones = []  # 💥 limpiar antes de volver a llenar
        self.config_table.setRowCount(len(configs))

        for row_idx, config in enumerate(configs):
            id_config, excel_file, sheet_name, excel_range, word_file, table_label, output_file, money_columns, header_rows = config
            self.id_configuraciones.append(id_config)
        self.config_table.setItem(row_idx, 0, QTableWidgetItem(excel_file))
        self.config_table.setItem(row_idx, 1, QTableWidgetItem(sheet_name))
        self.config_table.setItem(row_idx, 2, QTableWidgetItem(excel_range))
        self.config_table.setItem(row_idx, 3, QTableWidgetItem(word_file))
        self.config_table.setItem(row_idx, 4, QTableWidgetItem(table_label))
        self.config_table.setItem(row_idx, 5, QTableWidgetItem(output_file if output_file else ""))
        self.config_table.setItem(row_idx, 6, QTableWidgetItem(money_columns if money_columns else ""))
        self.config_table.setItem(row_idx, 7, QTableWidgetItem(str(header_rows) if header_rows else ""))


    def actualizar_todas(self):
        try:
            actualizar_todas_las_tablas()
            QMessageBox.information(self, "Actualización completa", "Todas las tablas fueron actualizadas correctamente.")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Ocurrió un error:\n{str(e)}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = TableUpdaterGUI()
    window.show()
    sys.exit(app.exec_())
