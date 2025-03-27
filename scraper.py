from playwright.sync_api import sync_playwright
import pandas as pd
import time
import random
import os
import json
from datetime import datetime
import multiprocessing
import threading

# Thread-safe print fonksiyonu
print_lock = threading.Lock()
def safe_print(*args, **kwargs):
    with print_lock:
        print(*args, **kwargs)

def random_delay(min_sec=1, max_sec=2):
    """Rastgele bir süre bekler"""
    time.sleep(random.uniform(min_sec, max_sec))

def extract_officer_data(page, officer_url, search_term):
    """Yönetici sayfasından verileri çeker"""
    try:
        safe_print(f"Visiting: {officer_url}")
        page.goto(officer_url)
        random_delay()
        
        # Doğum tarihi kontrolü
        dob_element = page.locator("dl #officer-date-of-birth-value")
        if dob_element.count() == 0:
            safe_print(f"Doğum tarihi bulunamadı, bu kişi atlanıyor: {officer_url}")
            return None
        
        dob = dob_element.inner_text().strip()
        safe_print(f"Doğum tarihi bulundu: {dob}")
        
        # Uyruk bilgisini çek
        nationality = ""
        nationality_element = page.locator("dl #nationality-value1")
        if nationality_element.count() > 0:
            nationality = nationality_element.inner_text().strip()
            safe_print(f"Uyruk bilgisi bulundu: {nationality}")
        
        # Yönetici adı
        officer_name = page.locator(".heading-xlarge").inner_text().strip()
        safe_print(f"Officer name: {officer_name}")
        
        # Atamalar
        appointments = []
        appointment_elements = page.locator(".appointment-1").all()
        safe_print(f"Found {len(appointment_elements)} appointments")
        
        for element in appointment_elements:
            appointment = {}
            
            # Şirket adı ve numarası - GÜNCELLENMIŞ SELEKTÖR
            company_link = element.locator("a[href*='/company/']")
            if company_link.count() > 0:
                company_name = company_link.inner_text().strip()
                company_url = company_link.get_attribute("href")
                company_number = ""
                # Parantez içindeki şirket numarasını çıkar
                import re
                match = re.search(r'\((\d+)\)', company_name)
                if match:
                    company_number = match.group(1)
                
                appointment["Şirket Adı"] = company_name
                appointment["Şirket Numarası"] = company_number
            else:
                continue  # Şirket linki yoksa atla
            
            # Diğer bilgileri topla - GÜNCELLENMIŞ SELEKTÖRLER
            selectors = {
                "Şirket Durumu": "#company-status-value-1",
                "Yazışma Adresi": "#correspondence-address-value-1",
                "Rol": "#appointment-type-value1",
                "Atanma Tarihi": "#appointed-value1",
                "Yönetilen Kanun": "#legal-authority-value-1",
                "Yasal Form": "#legal-form-value-1"
            }
            
            for field_name, selector in selectors.items():
                field_element = element.locator(selector)
                if field_element.count() > 0:
                    appointment[field_name] = field_element.inner_text().strip()
                else:
                    appointment[field_name] = ""
            
            appointments.append(appointment)
        
        if appointments:
            return {
                "Arama Terimi": search_term,
                "İsim": officer_name,
                "Doğum Tarihi": dob,
                "Uyruk": nationality,
                "Atamalar": appointments,
                "URL": officer_url
            }
        else:
            safe_print(f"No appointments found for {officer_name}")
            return None
    
    except Exception as e:
        safe_print(f"Error extracting officer data: {str(e)}")
        return None

def process_name(browser_context, search_term, max_pages=20):
    """Belirli bir isim için yöneticileri arar ve verilerini çeker"""
    page = browser_context.new_page()
    page.set_default_timeout(30000)
    
    base_url = "https://find-and-update.company-information.service.gov.uk"
    officers_data = []
    
    try:
        for page_num in range(1, max_pages + 1):
            # Arama sayfasına git
            search_url = f"{base_url}/search/officers?q={search_term}&page={page_num}"
            safe_print(f"Searching page {page_num}: {search_url}")
            
            page.goto(search_url)
            random_delay()
            
            # Sonuç var mı kontrol et
            no_results = page.locator(".search-no-results").count() > 0
            if no_results:
                safe_print(f"No results found for {search_term} on page {page_num}")
                break
            
            # Yönetici linklerini bul - DÜZELTILMIŞ SELEKTÖR
            # Sayfadaki tüm linkleri doğrudan HTML'den al
            links = page.evaluate("""
                () => {
                    const links = Array.from(document.querySelectorAll('a.govuk-link[href*="/officers/"]'));
                    return links.map(link => {
                        return {
                            href: link.getAttribute('href'),
                            text: link.textContent.trim()
                        };
                    });
                }
            """)
            
            if not links:
                safe_print(f"No officer links found on page {page_num}")
                break
            
            safe_print(f"Found {len(links)} officers on page {page_num}")
            
            # Her yönetici için veri çek
            for i, link in enumerate(links):
                try:
                    safe_print(f"Processing officer {i+1}/{len(links)} on page {page_num}")
                    
                    href = link.get('href')
                    if href:
                        officer_url = base_url + href
                        officer_data = extract_officer_data(page, officer_url, search_term)
                        
                        if officer_data:
                            officers_data.append(officer_data)
                            safe_print(f"Successfully processed: {officer_data['İsim']}")
                    
                    random_delay()
                except Exception as e:
                    safe_print(f"Error processing officer link: {str(e)}")
                    continue
            
            # Sonraki sayfa var mı?
            next_button = page.locator("a.page-next")
            if next_button.count() == 0 or not next_button.is_visible():
                safe_print(f"No more pages for {search_term}")
                break
    
    except Exception as e:
        safe_print(f"Error during search: {str(e)}")
    
    finally:
        page.close()
    
    return officers_data

def save_to_excel(all_data, filename="turkish_officers"):
    """Verileri Excel dosyasına kaydeder"""
    if not all_data:
        safe_print("No data to save")
        return
    
    # Tüm atamalar için liste
    all_appointments = []
    
    # Yönetici özet bilgileri için liste
    officers_summary = []
    
    for officer in all_data:
        # Yönetici özet bilgisi
        officer_summary = {
            "Arama Terimi": officer["Arama Terimi"],
            "İsim": officer["İsim"],
            "Atama Sayısı": len(officer["Atamalar"]),
            "URL": officer["URL"]
        }
        
        # Doğum tarihi ekle
        if "Doğum Tarihi" in officer:
            officer_summary["Doğum Tarihi"] = officer["Doğum Tarihi"]
        
        # Uyruk ekle
        if "Uyruk" in officer:
            officer_summary["Uyruk"] = officer["Uyruk"]
        
        officers_summary.append(officer_summary)
        
        # Atama bilgileri
        for appointment in officer["Atamalar"]:
            appointment_data = {
                "Arama Terimi": officer["Arama Terimi"],
                "Yönetici İsmi": officer["İsim"],
                "Şirket Adı": appointment.get("Şirket Adı", ""),
                "Şirket Numarası": appointment.get("Şirket Numarası", ""),
                "Şirket Durumu": appointment.get("Şirket Durumu", ""),
                "Rol": appointment.get("Rol", ""),
                "Yazışma Adresi": appointment.get("Yazışma Adresi", ""),
                "Atanma Tarihi": appointment.get("Atanma Tarihi", ""),
                "Yönetilen Kanun": appointment.get("Yönetilen Kanun", ""),
                "Yasal Form": appointment.get("Yasal Form", "")
            }
            
            # Doğum tarihi ekle
            if "Doğum Tarihi" in officer:
                appointment_data["Doğum Tarihi"] = officer["Doğum Tarihi"]
            
            # Uyruk ekle
            if "Uyruk" in officer:
                appointment_data["Uyruk"] = officer["Uyruk"]
                
            all_appointments.append(appointment_data)
    
    # Excel dosyası oluştur
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    excel_filename = f"{filename}_{timestamp}.xlsx"
    
    with pd.ExcelWriter(excel_filename, engine='openpyxl') as writer:
        # Yöneticiler sayfası
        pd.DataFrame(officers_summary).to_excel(writer, sheet_name='Yöneticiler', index=False)
        
        # Atamalar sayfası
        pd.DataFrame(all_appointments).to_excel(writer, sheet_name='Atamalar', index=False)
    
    safe_print(f"Data saved to {excel_filename}")
    safe_print(f"Total: {len(officers_summary)} officers and {len(all_appointments)} appointments")

def process_single_name(name):
    """Tek bir ismi işleyen fonksiyon - multiprocessing için"""
    safe_print(f"\n{'='*50}")
    safe_print(f"Processing name: {name}")
    safe_print(f"{'='*50}")
    
    # Playwright'ı başlat
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        
        try:
            # Context oluştur
            context = browser.new_context(
                viewport={'width': 1920, 'height': 1080},
                user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36'
            )
            
            # İsmi işle
            officers_data = process_name(context, name)
            
            # Context'i kapat
            context.close()
            
            # Tarayıcıyı kapat
            browser.close()
            
            if officers_data:
                safe_print(f"Found {len(officers_data)} officers for {name}")
                
                # JSON dosyasına kaydet
                json_filename = f"results_{name}.json"
                with open(json_filename, 'w', encoding='utf-8') as f:
                    json.dump(officers_data, f, ensure_ascii=False, indent=4)
                
                safe_print(f"Data for {name} saved to {json_filename}")
                return True
            else:
                safe_print(f"No results found for {name}")
                return False
                
        except Exception as e:
            safe_print(f"Error processing name {name}: {str(e)}")
            
            # Tarayıcıyı kapat
            if 'browser' in locals() and browser:
                browser.close()
                
            return False

def main():
    # Multiprocessing için
    multiprocessing.freeze_support()
    
    turkish_names = [
        "Ahmet", "Mehmet", "Mustafa", "Ali", "Huseyin", 
        "Hasan", "Ibrahim", "Yusuf", "Emre", "Burak", 
        "Onur", "Kerem", "Can", "Efe", "Omer", 
        "Serkan", "Kaan", "Mert", "Enes", "Arda", 
        "Cem", "Taha", "Hakan", "Sinan", "Baris", 
        "Ayse", "Fatma", "Zeynep", "Elif", "Hatice", 
        "Emine", "Aylin", "Ceren", "Busra", "Irem", 
        "Ebru", "Hande", "Duygu", "Selin", "Deniz", 
        "Esra", "Gamze", "Yasemin", "Sibel", "Gozde", 
        "Melike", "Tugba", "Dilara", "Sevgi", "Seyma"
    ]
    
    # İngilizce karakter versiyonlarını ekle
    english_names = []
    for name in turkish_names:
        english_name = name.replace("ı", "i").replace("ö", "o").replace("ü", "u").replace("ğ", "g").replace("ş", "s").replace("ç", "c")
        if english_name != name:
            english_names.append(english_name)
    
    # Tüm isimleri birleştir
    all_names = turkish_names + english_names
    
    # Daha önce işlenmiş isimleri kontrol et
    processed_names = []
    for name in all_names:
        json_filename = f"results_{name}.json"
        if os.path.exists(json_filename):
            processed_names.append(name)
    
    safe_print(f"Daha önce işlenmiş isimler: {processed_names}")
    
    # İşlenecek isimleri belirle
    names_to_process = [name for name in all_names if name not in processed_names]
    safe_print(f"İşlenecek isim sayısı: {len(names_to_process)}")
    
    # Daha önce toplanan verileri yükle
    all_data = []
    for name in processed_names:
        json_filename = f"results_{name}.json"
        try:
            with open(json_filename, 'r', encoding='utf-8') as f:
                data = json.load(f)
                all_data.extend(data)
                safe_print(f"{name} için {len(data)} yönetici verisi yüklendi")
        except Exception as e:
            safe_print(f"Hata: {name} verisi yüklenemedi - {str(e)}")
    
    # Excel dosyasına kaydet
    if all_data:
        safe_print("Mevcut verileri Excel'e kaydediyorum...")
        save_to_excel(all_data, "turkish_officers_current")
        safe_print(f"Toplam {len(all_data)} yönetici verisi Excel'e kaydedildi")
    
    # Otomatik olarak devam et
    safe_print("Veri toplamaya devam ediliyor...")
    
    # Maksimum paralel işlem sayısı
    max_processes = min(4, len(names_to_process))  # En fazla 4 paralel işlem
    
    # İşlenecek isimleri gruplara ayır
    name_chunks = []
    chunk_size = max(1, len(names_to_process) // max_processes)
    
    for i in range(0, len(names_to_process), chunk_size):
        name_chunks.append(names_to_process[i:i + chunk_size])
    
    # Paralel işlem için havuz oluştur
    with multiprocessing.Pool(processes=max_processes) as pool:
        # İsimleri paralel olarak işle
        results = pool.map(process_single_name, names_to_process)
    
    # İşlem tamamlandıktan sonra tüm JSON dosyalarını yükle
    all_data = []
    for name in all_names:
        json_filename = f"results_{name}.json"
        if os.path.exists(json_filename):
            try:
                with open(json_filename, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    if isinstance(data, list):
                        all_data.extend(data)
                    else:
                        all_data.append(data)
            except Exception as e:
                safe_print(f"Hata: {json_filename} yüklenemedi - {str(e)}")
    
    # Tüm verileri kaydet
    if all_data:
        save_to_excel(all_data, "turkish_officers_final")
        safe_print(f"Toplam {len(all_data)} yönetici verisi Excel'e kaydedildi")
    else:
        safe_print("No data found for any name")

if __name__ == "__main__":
    main()
