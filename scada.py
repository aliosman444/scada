import sys
import os
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QPushButton,
    QListWidget, QTableWidget, QTableWidgetItem, QMessageBox, QHeaderView, QProgressBar
)
from PyQt5.QtCore import Qt, QUrl, QThread, pyqtSignal
from PyQt5.QtWebEngineWidgets import QWebEngineView
import subprocess
import unicodedata
import re
from PyQt5.QtGui import QPixmap, QPalette, QColor
import paramiko

EXCEL_PATH = 'TURSOLAR_DATABASE_18.07.xlsx'

class PingWorker(QThread):
    result_signal = pyqtSignal(int, str)
    finished_signal = pyqtSignal()

    def __init__(self, df):
        super().__init__()
        self.df = df
        self._is_running = True

    def run(self):
        total = len(self.df)
        for idx, row in self.df.iterrows():
            if not self._is_running:
                break
            ip = str(row.get('STATİK IP', ''))
            if not ip:
                self.result_signal.emit(idx, '❌ IP yok')
                continue
            param = '-n' if sys.platform.startswith('win') else '-c'
            try:
                output = subprocess.check_output(['ping', param, '5', ip], universal_newlines=True, timeout=5)
                success = 'TTL=' in output or 'ttl=' in output
            except Exception:
                success = False
            if success:
                self.result_signal.emit(idx, '✅')
            else:
                self.result_signal.emit(idx, '❌')
        self.finished_signal.emit()

    def stop(self):
        self._is_running = False

class VPNTestWorker(QThread):
    result_signal = pyqtSignal(int, str)
    finished_signal = pyqtSignal()

    def __init__(self, df, hedef_ip='10.34.255.18'):
        super().__init__()
        self.df = df
        self.hedef_ip = hedef_ip
        self._is_running = True

    def run(self):
        for idx, row in self.df.iterrows():
            if not self._is_running:
                break
            saha_adi = str(row['SANTRAL ADI'])
            ip_adresi = str(row.get('STATİK IP', ''))
            kullanici_adi = str(row.get('ROUTER KULLANICI ADI', ''))
            sifre = str(row.get('ROUTER ŞİFRE', ''))
            port = row.get('SSH PORT', '')
            router = str(row.get('ROUTER', ''))

            try:
                port = int(float(port))
            except:
                self.result_signal.emit(idx, '❌ Hatalı port')
                continue

            if not ip_adresi:
                self.result_signal.emit(idx, '❌ IP yok')
                continue

            try:
                ssh = paramiko.SSHClient()
                ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
                ssh.connect(ip_adresi, port=port, username=kullanici_adi, password=sifre, timeout=5)

                if 'Four Faith' in router:
                    komut = f'ping -c5 {self.hedef_ip}'
                else:
                    komut = f'ping {self.hedef_ip}'

                stdin, stdout, stderr = ssh.exec_command(komut)
                cevap = stdout.read().decode().strip()

                if 'ttl=' in cevap or 'TTL=' in cevap:
                    self.result_signal.emit(idx, '✅')
                else:
                    self.result_signal.emit(idx, '❌ Yanıt yok')
            except Exception as e:
                self.result_signal.emit(idx, f'❌ Hata: {e}')

        self.finished_signal.emit()

    def stop(self):
        self._is_running = False


class SahaTakipArayuz(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Saha Takip Arayüzü')
        self.resize(1100, 700)
        self.df = pd.read_excel(EXCEL_PATH, dtype=str)
        self.df = self.df.fillna('')
        self.last_excel_mtime = self.get_excel_mtime()
        self.init_ui()

    def get_excel_mtime(self):
        try:
            return os.path.getmtime(EXCEL_PATH)
        except Exception:
            return None

    def init_ui(self):
        main_layout = QVBoxLayout()
        # Logo ekle
        logo_label = QLabel()
        pixmap = QPixmap('tursolar.png')
        logo_label.setPixmap(pixmap)
        logo_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        main_layout.addWidget(logo_label)
        # Arka plan ve buton renkleri için stil
        self.setStyleSheet('''
            QWidget {
                background-color: #fff;
            }
            QPushButton {
                background-color: #f37013;
                color: white;
                border-radius: 6px;
                padding: 6px 12px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #ff8c1a;
            }
            QLineEdit, QListWidget, QTableWidget {
                background-color: #fff;
                border: 1px solid #f37013;
                border-radius: 4px;
            }
        ''')
        search_layout = QHBoxLayout()

        self.search_label = QLabel('Santral Ara:')
        self.search_input = QLineEdit()
        self.search_input.textChanged.connect(self.update_list)
        search_layout.addWidget(self.search_label)
        search_layout.addWidget(self.search_input)

        self.refresh_button = QPushButton('Yenile')
        self.refresh_button.clicked.connect(self.refresh_excel)
        search_layout.addWidget(self.refresh_button)

        self.santral_list = QListWidget()
        self.santral_list.addItems(self.df['SANTRAL ADI'].astype(str).tolist())
        self.santral_list.currentTextChanged.connect(self.display_details)
        self.santral_list.setSelectionMode(self.santral_list.SingleSelection)

        self.details_table = QTableWidget()
        self.details_table.setColumnCount(len(self.df.columns))
        self.details_table.setHorizontalHeaderLabels(self.df.columns)
        header = self.details_table.horizontalHeader()
        if header is not None:
            header.setSectionResizeMode(QHeaderView.ResizeToContents)
        self.details_table.setRowCount(1)
        self.details_table.setEditTriggers(QTableWidget.NoEditTriggers)
        vheader = self.details_table.verticalHeader()
        if vheader is not None:
            vheader.setSectionResizeMode(QHeaderView.ResizeToContents)
        self.details_table.setWordWrap(True)
        self.details_table.setMinimumHeight(120)

        button_layout = QHBoxLayout()
        self.ping_button = QPushButton('Ping At')
        self.ping_button.clicked.connect(self.ping_selected)
        self.vpn_test_buton = QPushButton('VPN Test Et')
        self.vpn_test_buton.clicked.connect(self.vpn_test)
        self.web_button = QPushButton('Modem Arayüzüne Git')
        self.web_button.clicked.connect(self.open_modem_web)
        self.ekk_web_button = QPushButton('EKK Arayüzüne Git')
        self.ekk_web_button.clicked.connect(self.ekk_open_modem_web)
        self.excel_button = QPushButton('Excel Dosyasını Aç')
        self.excel_button.clicked.connect(self.open_excel_file)
        self.bulk_vpn_button = QPushButton('Toplu VPN Test Et')
        self.bulk_vpn_button.clicked.connect(self.bulk_vpn_test)
        self.bulk_ping_button = QPushButton('Toplu Ping At')
        self.bulk_ping_button.clicked.connect(self.bulk_ping)
        button_layout.addWidget(self.ping_button)
        button_layout.addWidget(self.vpn_test_buton)
        button_layout.addWidget(self.web_button)
        button_layout.addWidget(self.ekk_web_button)
        button_layout.addWidget(self.excel_button)
        button_layout.addWidget(self.bulk_vpn_button)
        button_layout.addWidget(self.bulk_ping_button)

        self.web_view = QWebEngineView()
        self.web_view.setMinimumHeight(0)
        self.web_view.setMaximumHeight(0)
        self.web_expanded = False
        self.web_view.setVisible(False)

        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)

        main_layout.addLayout(search_layout)
        main_layout.addWidget(self.santral_list)
        main_layout.addWidget(self.details_table)
        main_layout.addLayout(button_layout)
        main_layout.addWidget(self.progress_bar)
        main_layout.addWidget(self.web_view)
        self.setLayout(main_layout)

    def update_list(self, text):
        self.santral_list.clear()
        filtered = self.df[self.df['SANTRAL ADI'].str.contains(text, case=False, na=False)]
        self.santral_list.addItems(filtered['SANTRAL ADI'].astype(str).tolist())

    def display_details(self, santral_adi):
        if not santral_adi:
            self.selected_row = None
            return
        row = self.df[self.df['SANTRAL ADI'] == santral_adi]
        if row.empty:
            self.selected_row = None
            return
        row = row.iloc[0]
        self.details_table.setRowCount(1)
        for i, col in enumerate(self.df.columns):
            item = QTableWidgetItem(str(row[col]))
            item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
            item.setFlags(item.flags() ^ Qt.ItemIsEditable)
            item.setToolTip(str(row[col]))
            self.details_table.setItem(0, i, item)
        self.details_table.resizeRowsToContents()
        self.details_table.resizeColumnsToContents()
        self.selected_row = row

    def ping_selected(self):
        row = getattr(self, 'selected_row', None)
        if row is None:
            QMessageBox.warning(self, 'Uyarı', 'Lütfen önce bir santral seçin!')
            return
        ip = str(row['STATİK IP'])
        if not ip:
            QMessageBox.warning(self, 'Uyarı', 'IP adresi bulunamadı!')
            return
        param = '-n' if sys.platform.startswith('win') else '-c'
        try:
            output = subprocess.check_output(['ping', param, '4', ip], universal_newlines=True, timeout=5)
            QMessageBox.information(self, 'Ping Sonucu', output)
        except Exception as e:
            QMessageBox.warning(self, 'Ping Hatası', f'Ping atılamadı: {e}')

    def vpn_test(self):
        row = getattr(self, 'selected_row', None)
        if row is not None:
            saha_adi = str(row['SANTRAL ADI'])
            ip_adresi = str(row['STATİK IP'])
            kullanici_adi = str(row['ROUTER KULLANICI ADI'])
            sifre = str(row['ROUTER ŞİFRE'])
            port = str(row['SSH PORT'])
            hedef_ip = '10.34.255.18'
            router = str(row['ROUTER'])
        if row is None:
            QMessageBox.warning(self, ' Uyarı', 'Lütfen önce bir santral seçin!')
            return
        if not ip_adresi:
            QMessageBox.warning(self, 'Uyarı', 'IP adresi bulunamadı!')
            return
        try:
            ssh = paramiko.SSHClient()
            ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
            ssh.connect(ip_adresi, port=port, username=kullanici_adi, password=sifre, timeout=5)
            if 'Four Faith' in router:
                ping_komutu = f"ping -c5 {hedef_ip}"
                print(ping_komutu)
            else:
                ping_komutu = f"ping {hedef_ip}"
                print(ping_komutu)

            stdin, stdout, stderr = ssh.exec_command(ping_komutu)
            cevap = stdout.read().decode().strip()

            QMessageBox.information(self, 'Komut Çıktısı', cevap)

        except Exception as hata:
            QMessageBox.warning(self, 'Uyarı', f"[X] {saha_adi} - {ip_adresi} → Bağlantı Hatası: {hata}")

    def open_modem_web(self):
        row = getattr(self, 'selected_row', None)
        if row is None:
            QMessageBox.warning(self, 'Uyarı', 'Lütfen önce bir santral seçin!')
            return
        link = str(row.get('Modem Erişim', ''))
        if not link or link.strip() == '':
            QMessageBox.warning(self, 'Hata', 'Modem Erişim linki bulunamadı!')
            return
        import webbrowser
        webbrowser.open(link)

    def ekk_open_modem_web(self):
        row = getattr(self, 'selected_row', None)
        if row is None:
            QMessageBox.warning(self, 'Uyarı', 'Lütfen önce bir santral seçin!')
            return
        link = str(row.get('Ekk Erişim', ''))
        if not link or link.strip() == '':
            QMessageBox.warning(self, 'Hata', 'EKK Erişim linki bulunamadı!')
            return
        import webbrowser
        webbrowser.open(link)

    def open_excel_file(self):
        try:
            if sys.platform.startswith('win'):
                os.startfile(EXCEL_PATH)
            elif sys.platform.startswith('darwin'):
                subprocess.call(['open', EXCEL_PATH])
            else:
                subprocess.call(['xdg-open', EXCEL_PATH])
        except Exception as e:
            QMessageBox.warning(self, 'Hata', f'Excel dosyası açılamadı: {e}')

    def bulk_ping(self):
        from PyQt5.QtWidgets import QDialog, QVBoxLayout, QTableWidget, QTableWidgetItem, QHeaderView, QPushButton, QLabel
        from PyQt5.QtCore import Qt, QCoreApplication
        total = len(self.df)
        dialog = QDialog(self)
        dialog.setWindowTitle('Toplu Ping Sonuçları')
        # '?' butonunu kaldır
        dialog.setWindowFlags(dialog.windowFlags() & ~Qt.WindowContextHelpButtonHint)
        layout = QVBoxLayout()
        table = QTableWidget(total, 2)
        table.setHorizontalHeaderLabels(['Santral Adı', 'Ping Sonucu'])
        table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        layout.addWidget(table)
        info_label = QLabel('Ping atılıyor...')
        layout.addWidget(info_label)
        close_btn = QPushButton('Kapat')
        close_btn.clicked.connect(dialog.accept)
        layout.addWidget(close_btn)
        dialog.setLayout(layout)
        dialog.resize(400, 600)
        # Santral isimlerini ekle
        for idx, row in self.df.iterrows():
            name = str(row['SANTRAL ADI'])
            table.setItem(idx, 0, QTableWidgetItem(name))
            table.setItem(idx, 1, QTableWidgetItem('Bekleniyor...'))
        dialog.show()
        QCoreApplication.processEvents()
        self.ping_worker = PingWorker(self.df)
        def update_result(idx, result):
            if table is not None:
                table.setItem(idx, 1, QTableWidgetItem(result))
                QCoreApplication.processEvents()
        def finish():
            info_label.setText('Tüm pingler tamamlandı!')
            QCoreApplication.processEvents()
        self.ping_worker.result_signal.connect(update_result)
        self.ping_worker.finished_signal.connect(finish)
        self.ping_worker.start()
        def on_dialog_close():
            if self.ping_worker.isRunning():
                try:
                    self.ping_worker.result_signal.disconnect()
                except Exception:
                    pass
                try:
                    self.ping_worker.finished_signal.disconnect()
                except Exception:
                    pass
                self.ping_worker.stop()
                self.ping_worker.wait()
        dialog.finished.connect(on_dialog_close)
        dialog.exec_()

    def bulk_vpn_test(self):
        from PyQt5.QtWidgets import QDialog, QVBoxLayout, QTableWidget, QTableWidgetItem, QHeaderView, QPushButton, QLabel
        from PyQt5.QtCore import Qt, QCoreApplication

        total = len(self.df)
        dialog = QDialog(self)
        dialog.setWindowTitle('Toplu VPN Test Sonuçları')
        dialog.setWindowFlags(dialog.windowFlags() & ~Qt.WindowContextHelpButtonHint)

        layout = QVBoxLayout()
        table = QTableWidget(total, 2)
        table.setHorizontalHeaderLabels(['Santral Adı', 'VPN Durumu'])
        table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        layout.addWidget(table)

        info_label = QLabel('VPN testleri yapılıyor...')
        layout.addWidget(info_label)

        close_btn = QPushButton('Kapat')
        close_btn.clicked.connect(dialog.accept)
        layout.addWidget(close_btn)

        dialog.setLayout(layout)
        dialog.resize(450, 600)

        for idx, row in self.df.iterrows():
            name = str(row['SANTRAL ADI'])
            table.setItem(idx, 0, QTableWidgetItem(name))
            table.setItem(idx, 1, QTableWidgetItem('Bekleniyor...'))

        dialog.show()
        QCoreApplication.processEvents()

        self.vpn_worker = VPNTestWorker(self.df)
        
        def update_result(idx, result):
            table.setItem(idx, 1, QTableWidgetItem(result))
            QCoreApplication.processEvents()

        def finish():
            info_label.setText('Tüm VPN testleri tamamlandı!')
            QCoreApplication.processEvents()

        self.vpn_worker.result_signal.connect(update_result)
        self.vpn_worker.finished_signal.connect(finish)
        self.vpn_worker.start()

        def on_dialog_close():
            if self.vpn_worker.isRunning():
                try:
                    self.vpn_worker.result_signal.disconnect()
                except:
                    pass
                try:
                    self.vpn_worker.finished_signal.disconnect()
                except:
                    pass
                self.vpn_worker.stop()
                self.vpn_worker.wait()
        
        dialog.finished.connect(on_dialog_close)
        dialog.exec_()


    def refresh_excel(self):
        current_mtime = self.get_excel_mtime()
        if current_mtime != self.last_excel_mtime:
            try:
                new_df = pd.read_excel(EXCEL_PATH, dtype=str)
                new_df = new_df.fillna('')
                self.df = new_df
                self.last_excel_mtime = current_mtime
                self.santral_list.clear()
                self.santral_list.addItems(self.df['SANTRAL ADI'].astype(str).tolist())
                self.details_table.setColumnCount(len(self.df.columns))
                self.details_table.setHorizontalHeaderLabels(self.df.columns)
                self.details_table.clearContents()
                self.details_table.setRowCount(1)
                self.selected_row = None
                QMessageBox.information(self, 'Yenilendi', 'Excel dosyası güncellendi ve veriler yenilendi.')
            except Exception as e:
                QMessageBox.warning(self, 'Hata', f'Excel dosyası okunamadı: {e}')
        else:
            QMessageBox.information(self, 'Bilgi', 'Excel dosyasında bir değişiklik yok.')


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = SahaTakipArayuz()
    window.show()
    sys.exit(app.exec_())

