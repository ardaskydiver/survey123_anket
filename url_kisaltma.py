import requests

def check_url_exists(long_url):
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
        print("URL kontrolü sırasında hata oluştu:", str(e))
        return None

def shorten_url(long_url):
    # Önce URL'nin daha önce kısaltılıp kısaltılmadığını kontrol et
    existing_short_url = check_url_exists(long_url)
    if existing_short_url:
        print(f"URL zaten kısaltılmış: {existing_short_url}")
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
        response.raise_for_status()  # HTTP hatalarını kontrol et
        
        data = response.json()
        
        if data.get("status") == "success":
            return data["shorturl"]
        else:
            print("Hata:", data.get("message", "Bilinmeyen hata"))
            return None
            
    except Exception as e:
        print("İstek sırasında hata oluştu:", str(e))
        return None

# Kullanım Örneği
long_url = "https://egitim.maps.arcgis.com/home/item.html?id=de2a297c3d73478ab170b88fdb68834b"
short_url = shorten_url(long_url)

if short_url:
    print(f"Kısaltılmış URL: {short_url}")
else:
    print("URL kısaltma başarısız oldu")