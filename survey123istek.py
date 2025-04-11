import pandas as pd
import smtplib
import hashlib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from urllib.parse import quote
from datetime import datetime
import logging
import sys
import codecs
import requests
import os
import base64

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

# Loglama ayarları
# UTF-8 encoding ile dosya handler'ı oluştur
file_handler = logging.FileHandler("survey_email_log.txt", encoding='utf-8')
file_handler.setLevel(logging.INFO)

# Console handler'ı oluştur ve hata yönetimi ekle
class SafeStreamHandler(logging.StreamHandler):
    def emit(self, record):
        try:
            msg = self.format(record)
            stream = self.stream
            # Windows console için özel işlem
            if sys.platform == 'win32' and hasattr(stream, 'fileno') and stream.fileno() == 1:
                # ASCII olmayan karakterleri kaldır
                msg = msg.encode('ascii', 'replace').decode('ascii')
            stream.write(msg + self.terminator)
            self.flush()
        except Exception:
            self.handleError(record)

console_handler = SafeStreamHandler()
console_handler.setLevel(logging.INFO)

# Formatter oluştur
formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
file_handler.setFormatter(formatter)
console_handler.setFormatter(formatter)

# Logger'ı yapılandır
logger = logging.getLogger()
logger.setLevel(logging.INFO)
logger.addHandler(file_handler)
logger.addHandler(console_handler)

def generate_encrypted_code(email, date):
    date_str = date.strftime("%Y%m%d")
    raw_string = f"{email}{date_str}".encode('utf-8')
    return hashlib.sha256(raw_string).hexdigest()[:12]

def check_url_exists(long_url):
    """
    URL'nin daha önce kısaltılıp kısaltılmadığını kontrol eder
    """
    api_url = "https://esritr.link/yourls-api.php"
    
    params = {
        "username": "esritr",
        "password": "2NC0BRs7vbU4UncUAiqYDqxLkPi2ziNoozkG1",
        "action": "url-stats",
        "format": "json",
        "url": long_url
    }
    
    try:
        response = requests.get(api_url, params=params)
        response.raise_for_status()
        
        data = response.json()
        
        if data.get("status") == "success" and data.get("link"):
            return data["link"]["shorturl"]
        else:
            return None
            
    except Exception as e:
        logging.error(f"URL kontrolu sirasinda hata olustu: URL kısaltılmadan gönderilecek")
        return None

def shorten_url(long_url):
    """
    URL'yi kısaltır, hata durumunda None döndürür
    """
    # Önce URL'nin daha önce kısaltılıp kısaltılmadığını kontrol et
    existing_short_url = check_url_exists(long_url)
    if existing_short_url:
        logging.info(f"URL zaten kisaltilmis: {existing_short_url}")
        return existing_short_url
    
    api_url = "https://esritr.link/yourls-api.php"
    
    params = {
        "username": "esritr",
        "password": "2NC0BRs7vbU4UncUAiqYDqxLkPi2ziNoozkG1",
        "action": "shorturl",
        "format": "json",
        "url": long_url
    }
    
    try:
        response = requests.get(api_url, params=params)
        
        # 404 hatası durumunda orijinal URL'yi kullan
        if response.status_code == 404:
            logging.warning("URL kisaltma servisi bulunamadi (404), orijinal URL kullanilacak")
            return None
            
        response.raise_for_status()  # Diğer HTTP hatalarını kontrol et
        
        data = response.json()
        
        if data.get("status") == "success":
            return data["shorturl"]
        else:
            logging.warning(f"URL kisaltma hatasi: orijinal URL kullanilacak")
            return None
            
    except Exception as e:
        logging.error(f"URL kisaltma istegi sirasinda hata olustu: orijinal URL kullanilacak")
        return None

def send_survey_emails(excel_path, egitmen_kodu=None, egitim_kodu=None):
    # Excel verilerini oku
    df = pd.read_excel(excel_path, parse_dates=['egitim_tarihi'])
    
    # SMTP ayarları
    smtp_server = "smtp.office365.com"
    port = 587
    sender_email = "donotreply@esri.com.tr"
    password = "Kaya1946"

    # Banner resmi için değişkeni başlangıçta None olarak tanımla
    img_base64 = None

    try:
        # SMTP bağlantısını başlat
        server = smtplib.SMTP(smtp_server, port)
        server.starttls()
        server.login(sender_email, password)
        logging.info("SMTP baglantisi basariyla kuruldu")

        # Banner resmini yükle
        banner_path = resource_path("banner.png")
        if os.path.exists(banner_path):
            try:
                with open(banner_path, 'rb') as f:
                    img_data = f.read()
                    img_base64 = base64.b64encode(img_data).decode('utf-8')
                logging.info("Banner resmi base64'e cevrildi")
            except Exception as e:
                logging.error(f"Banner resmi base64'e cevrilirken hata olustu: {str(e)}")
                img_base64 = None
        else:
            logging.error(f"Banner resmi bulunamadi: {banner_path}")

        for _, row in df.iterrows():
            try:
                # Kullanıcı bilgilerini al
                first_name = row['Adı']
                last_name = row['Soyadı']
                email = row['Mail Adresi']
                egitim_tarihi = row['egitim_tarihi']

                # Şifreli kodu oluştur
                encrypted_code = generate_encrypted_code(email, egitim_tarihi)
                
                # URL parametrelerini hazırla
                base_url = "https://survey123.arcgis.com/share/de2a297c3d73478ab170b88fdb68834b"
                params = {
                    "field:edate": egitim_tarihi.strftime("%Y-%m-%d"),
                    "field:kcode": encrypted_code,
                    "field:egt": egitmen_kodu,
                    "field:egt_k": egitim_kodu
                }
                
                # URL'yi oluştur
                encoded_params = "&".join(
                    [f"{key}={quote(str(value))}" for key, value in params.items()]
                )
                survey_url = f"{base_url}?{encoded_params}"
                
                logging.info(f"{email} icin anket URL'si olusturuldu")
                
                # URL'yi kısaltmayı dene
                short_url = shorten_url(survey_url)
                final_url = short_url if short_url else survey_url
                
                if short_url:
                    logging.info(f"{email} icin URL basariyla kisaltildi: {short_url}")
                else:
                    logging.info(f"{email} icin orijinal URL kullanilacak")

                # E-posta içeriğini oluştur
                if img_base64:
                    # Base64 ile gömülü resimli HTML
                    message = f"""<html>
<body style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px;">
<div style="text-align: left; margin-bottom: 20px;">
    <img src="data:image/png;base64,{img_base64}" alt="Esri Türkiye Banner" style="width:100%; max-width:600px;">
</div>

<p>Sayın {first_name} {last_name},</p>

<p>Sizi tamamlamış olduğunuz eğitim hakkındaki görüşlerinizi almak üzere </p>
<p>Esri Türkiye Eğitim Değerlendirme Anketimize katılmaya davet ediyoruz.</p>

<p>Değerlendirmeleriniz kurumumuzun eğitim kalitesinin arttırılması için önemlidir.</p>

<p>Ankete katılmak için, lütfen aşağıdaki bağlantıya tıklayınız.</p>

<p>Saygılarımızla,</p>

<p>Esri Türkiye Eğitim Ekibi</p>

<p style="text-align: left; margin: 20px 0;">
    <a href="{final_url}" style="color: #0079c1; text-decoration: underline;">Anketi doldurmak için lütfen buraya tıklayınız</a>
</p>

<p>------------------------------------------------------------------</p>

<p>Eğer yukarıdaki bağlantı çalışmıyorsa, </p>

<p>lütfen aşağıdaki URL'yi kopyalayınız ve tarayıcınızda açınız:</p>

<p>{final_url}</p>

<div style="text-align: left; margin-top: 20px;">
    <img src="data:image/png;base64,{img_base64}" alt="Esri Türkiye Banner" style="width:100%; max-width:600px;">
</div>
</body>
</html>"""
                else:
                    # Resim olmadan düz HTML
                    message = f"""<html>
<body style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px;">
<p>Sayin {first_name} {last_name},</p>

<p>Sizi tamamlamış olduğunuz eğitim hakkındaki görüşlerinizi almak üzere </p>
<p>Esri Türkiye Eğitim Değerlendirme Anketimize katılmaya davet ediyoruz.</p>

<p>Değerlendirmeleriniz kurumumuzun eğitim kalitesinin arttırılması için önemlidir.</p>

<p>Ankete katılmak için, lütfen aşağıdaki bağlantıya tıklayınız.</p>

<p>Saygılarımızla,</p>

<p>Esri Türkiye Eğitim Ekibi</p>

<p style="text-align: left; margin: 20px 0;">
    <a href="{final_url}" style="color: #0079c1; text-decoration: underline;">Anketi doldurmak için buraya tıklayınız</a>
</p>

<p>------------------------------------------------------------------</p>

<p>Eğer yukarıdaki bağlantı çalışmıyorsa, </p>

<p>lütfen aşağıdaki URL'yi kopyalayınız ve tarayıcınızda açınız:</p>

<p>{final_url}</p>

</body>
</html>"""

                # E-posta mesajını hazırla
                msg = MIMEMultipart('alternative')
                msg['From'] = sender_email
                msg['To'] = email
                msg['Subject'] = "Esri Turkiye Egitim Degerlendirme Anketi"
                
                # HTML içeriği ekle
                html_part = MIMEText(message, 'html')
                msg.attach(html_part)

                # E-postayı gönder
                server.sendmail(sender_email, email, msg.as_string())
                logging.info(f"{email} adresine basariyla mail gonderildi.")
                
            except Exception as e:
                logging.error(f"{email} icin islem sirasinda hata olustu: {str(e)}")
                continue

    except Exception as e:
        logging.error(f"Genel hata olustu: {str(e)}")
    finally:
        server.quit()
        logging.info("SMTP baglantisi kapatildi")

# Kullanım örneği
if __name__ == "__main__":
    logging.info("E-posta gonderme islemi baslatiliyor")
    send_survey_emails("C:/py/kullanicilar.xlsx")
    logging.info("E-posta gonderme islemi tamamlandi")