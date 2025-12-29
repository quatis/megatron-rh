import sys
import os
import subprocess
from datetime import datetime, date
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton, QLabel,
    QFileDialog, QMessageBox, QComboBox, QInputDialog
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(BASE_DIR, "Config.xlsx")

changeDate = datetime.today().strftime('%d/%m/%Y')
LANGUAGE = "EN"

class RoboApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("🤖 Robozinho RH")
        self.setGeometry(400, 200, 500, 400)
        self.layout = QVBoxLayout()
        self.setLayout(self.layout)

        self.header = QLabel("👾 Our.Success - AUTOMATIZADO 👾\n🤖 Robozinho do RH 🤖")
        self.header.setStyleSheet("font-weight: bold; font-size: 16px; color: #00BCD4;")
        self.layout.addWidget(self.header)

        # Botões do menu
        self.btn_po = QPushButton("📅 Planned Occupation Date")
        self.btn_po.clicked.connect(lambda: self.run_script("po.py", "Planned Occupation Date"))

        self.btn_tbh = QPushButton("👤 To Be Hired")
        self.btn_tbh.clicked.connect(lambda: self.run_script("tbh.py", "To Be Hired"))

        self.btn_change_date = QPushButton("🔮 Alterar Data da Mudança")
        self.btn_change_date.clicked.connect(self.configure_change_date)

        self.btn_open_excel = QPushButton("📁 Abrir Arquivo Config")
        self.btn_open_excel.clicked.connect(self.open_config_file)

        self.btn_exit = QPushButton("🚪 Sair")
        self.btn_exit.clicked.connect(self.close)

        for btn in [self.btn_po, self.btn_tbh, self.btn_change_date, self.btn_open_excel, self.btn_exit]:
            btn.setStyleSheet("padding: 10px; font-size: 14px;")
            self.layout.addWidget(btn)

        self.status_label = QLabel(f"📅 Data atual: {changeDate} | 🌎 Idioma: {LANGUAGE}")
        self.layout.addWidget(self.status_label)

    def run_script(self, script_name, script_description):
        script_path = os.path.join(BASE_DIR, "script", script_name)
        if not os.path.exists(script_path):
            QMessageBox.critical(self, "Erro", f"Script {script_name} não encontrado!")
            return
        try:
            env = os.environ.copy()
            env["CHANGE_DATE"] = changeDate
            env["LANGUAGE"] = LANGUAGE
            subprocess.run([sys.executable, script_path], env=env)
            QMessageBox.information(self, "Sucesso", f"{script_description} executado com sucesso!")
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao executar script: {e}")

    def configure_change_date(self):
        global changeDate, LANGUAGE
        # Escolher idioma
        lang, ok = QInputDialog.getItem(self, "Idioma", "Selecione o idioma:", ["EN", "PT"], 0, False)
        if ok:
            LANGUAGE = lang

        # Escolher data
        today_option = QMessageBox.question(self, "Data de Mudança",
                                            "Deseja usar a data de hoje?",
                                            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if today_option == QMessageBox.StandardButton.Yes:
            changeDate = date.today().strftime("%d/%m/%Y") if LANGUAGE == "PT" else date.today().strftime("%m/%d/%Y")
        else:
            new_date, ok = QInputDialog.getText(self, "Nova Data", "Digite a nova data (dd/mm/yyyy):")
            if ok:
                try:
                    parsedDate = datetime.strptime(new_date, '%d/%m/%Y')
                    changeDate = parsedDate.strftime("%d/%m/%Y") if LANGUAGE == "PT" else parsedDate.strftime("%m/%d/%Y")
                except:
                    QMessageBox.warning(self, "Erro", "Data inválida!")

        self.status_label.setText(f"📅 Data atual: {changeDate} | 🌎 Idioma: {LANGUAGE}")

    def open_config_file(self):
        if os.path.exists(EXCEL_PATH):
            if os.name == 'nt':
                os.startfile(EXCEL_PATH)
            else:
                subprocess.run(['xdg-open', EXCEL_PATH])
        else:
            QMessageBox.warning(self, "Erro", "Arquivo Config.xlsx não encontrado!")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = RoboApp()
    window.show()
    sys.exit(app.exec())
