import sys
import os
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                            QHBoxLayout, QLabel, QPushButton, QFileDialog, 
                            QCalendarWidget, QComboBox, QMessageBox)
from PyQt5.QtGui import QPixmap, QFont
from PyQt5.QtCore import Qt, QDate
import pandas as pd
from survey123istek import send_survey_emails

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

class SurveyApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Esri Türkiye Eğitim Hizmetleri | Eğitim Değerlendirme Anketi")
        self.setGeometry(100, 100, 800, 600)
        
        # Eğitmen bilgileri
        self.egitmenler = {
            "Arda Çetinkaya": "101",
            "Esra Aydın Demiröz": "102",
            "İrfan Taşköprü": "103",
            "Tuğçe Ateş": "104"
        }

        # Eğitim Bilgileri
        self.egitimler = {
            "ArcGIS Online": "AGON",
            "ArcGIS Pro: Temel Uygulamalar": "APEW",
            "ArcGIS City Engine ile 3B Şehirler Oluşturma": "BECE",
            "ArcGIS Dashboards'un Temelleri": "DASH",
            "ArcGIS Enterprise ile İçeriklerin Paylaşılması": "ESHA",
            "ArcGIS Enterprise: Temel Dağıtım Yapılandırması": "EBAS",
            "ArcGIS ile CBS'ye Giriş": "GISA",
            "ArcGIS ile Coğrafi Verilerin Yönetimi": "GDAT",
            "ArcGIS ile Hikaye Oluşturma": "STRY",
            "ArcGIS ile Saha Verilerini Toplama ve Yönetme": "FIDA",
            "ArcGIS Online: Temel Uygulamalar": "APEW",
            "ArcGIS Pro ile Mekansal Analiz Uygulamaları": "SNAP",
            "ArcGIS Pro: Temel Uygulamalar": "APEW",
            "ArcGIS Pro'da Görüntü Analizi": "IMAP",
            "ArcGIS Utility Network'te Çalışma": "UTIL",
            "ArcMap'ten ArcGIS Pro'ya Geçiş": "PROM",
            "Madencilikte CBS Uygulamaları": "Madencilik",
            "Öğretmenler İçin CBS": "ÖCBS",
            "Şehir Planlamada ArcGIS": "PLAN",
            "TUCBS Uyumlu Veri Tabanı Altyapısı Oluşturma": "TUCBS",
            "Diğer": "OTR"
        }
        
        # Ana widget ve layout
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QVBoxLayout(main_widget)
        
        # Logo
        logo_label = QLabel()
        logo_path = resource_path("esri_logo.png")
        logo_pixmap = QPixmap(logo_path)
        if not logo_pixmap.isNull():
            logo_pixmap = logo_pixmap.scaled(415, 121, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            logo_label.setPixmap(logo_pixmap)
        else:
            logo_label.setText("Esri Logo")
            logo_label.setStyleSheet("font-size: 24px; font-weight: bold;")
        logo_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(logo_label)
        
        # Eğitim tarihi seçimi
        date_layout = QHBoxLayout()
        date_label = QLabel("Eğitimin Sonlandığı Tarihi Seçiniz:")
        date_label.setFont(QFont("Arial", 12))
        self.calendar = QCalendarWidget()
        self.calendar.setGridVisible(True)
        self.calendar.setMinimumDate(QDate.currentDate().addDays(-365))  # Son 1 yıl
        self.calendar.setMaximumDate(QDate.currentDate())
        date_layout.addWidget(date_label)
        date_layout.addWidget(self.calendar)
        main_layout.addLayout(date_layout)
        
        # Excel dosyası seçimi
        file_layout = QHBoxLayout()
        file_label = QLabel("Katılımcı Listesini Seçiniz:")
        file_label.setFont(QFont("Arial", 12))
        self.file_path_label = QLabel("Dosya seçilmedi")
        self.file_path_label.setFont(QFont("Arial", 10))
        self.file_path_label.setStyleSheet("color: gray;")
        browse_button = QPushButton("Dosya Seç")
        browse_button.clicked.connect(self.browse_file)
        file_layout.addWidget(file_label)
        file_layout.addWidget(self.file_path_label)
        file_layout.addWidget(browse_button)
        main_layout.addLayout(file_layout)
        
        # Eğitmen seçimi
        egitmen_layout = QHBoxLayout()
        egitmen_label = QLabel("Eğitmen Seçiniz:")
        egitmen_label.setFont(QFont("Arial", 12))
        self.egitmen_combo = QComboBox()
        self.egitmen_combo.setFont(QFont("Arial", 12))
        self.egitmen_combo.addItems(self.egitmenler.keys())
        egitmen_layout.addWidget(egitmen_label)
        egitmen_layout.addWidget(self.egitmen_combo)
        main_layout.addLayout(egitmen_layout)

        # Eğitim seçimi
        egitim_layout = QHBoxLayout()
        egitim_label = QLabel("Eğitim Seçiniz:")
        egitim_label.setFont(QFont("Arial", 12))
        self.egitim_combo = QComboBox()
        self.egitim_combo.setFont(QFont("Arial", 12))
        self.egitim_combo.addItems(self.egitimler.keys())
        egitim_layout.addWidget(egitim_label)
        egitim_layout.addWidget(self.egitim_combo)
        main_layout.addLayout(egitim_layout)
        
        # Gönder butonu
        send_button = QPushButton("Anket E-postalarını Gönder")
        send_button.setFont(QFont("Arial", 12, QFont.Bold))
        send_button.setStyleSheet("background-color: #0079c1; color: white; padding: 10px;")
        send_button.clicked.connect(self.send_emails)
        main_layout.addWidget(send_button)
        
        # Durum mesajı
        self.status_label = QLabel("")
        self.status_label.setAlignment(Qt.AlignCenter)
        self.status_label.setFont(QFont("Arial", 10))
        main_layout.addWidget(self.status_label)
        
        # Boşluk ekle
        main_layout.addStretch()
        
        # Excel dosya yolu
        self.excel_path = ""
        
    def browse_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Excel Dosyası Seç", "", "Excel Dosyaları (*.xlsx *.xls)"
        )
        if file_path:
            self.excel_path = file_path
            self.file_path_label.setText(os.path.basename(file_path))
            self.file_path_label.setStyleSheet("color: black;")
    
    def send_emails(self):
        if not self.excel_path:
            QMessageBox.warning(self, "Hata", "Lütfen bir Excel dosyası seçin.")
            return
        
        try:
            # Excel dosyasını oku
            df = pd.read_excel(self.excel_path)
            
            # Eğitim tarihini ayarla
            selected_date = self.calendar.selectedDate().toPyDate()
            df['egitim_tarihi'] = selected_date
            
            # Eğitmen kodunu ayarla
            selected_egitmen = self.egitmen_combo.currentText()
            egitmen_kodu = self.egitmenler[selected_egitmen]
            
            # Eğitim kodunu ayarla
            selected_egitim = self.egitim_combo.currentText()
            egitim_kodu = self.egitimler[selected_egitim]
            
            # Geçici Excel dosyası oluştur
            temp_excel_path = "temp_survey_data.xlsx"
            df.to_excel(temp_excel_path, index=False)
            
            # E-postaları gönder
            self.status_label.setText("E-postalar gönderiliyor...")
            self.status_label.setStyleSheet("color: blue;")
            QApplication.processEvents()
            
            send_survey_emails(temp_excel_path, egitmen_kodu, egitim_kodu)
            
            # Geçici dosyayı sil
            if os.path.exists(temp_excel_path):
                os.remove(temp_excel_path)
            
            self.status_label.setText("E-postalar başarıyla gönderildi!")
            self.status_label.setStyleSheet("color: green;")
            
        except Exception as e:
            self.status_label.setText(f"Hata oluştu: {str(e)}")
            self.status_label.setStyleSheet("color: red;")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = SurveyApp()
    window.show()
    sys.exit(app.exec_()) 