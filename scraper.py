from playwright.sync_api import sync_playwright
import pandas as pd
import time
import random
import os
import json
from datetime import datetime

def random_delay(min_sec=1, max_sec=2):
    """Rastgele bir süre bekler"""
    time.sleep(random.uniform(min_sec, max_sec))

def extract_officer_data(page, officer_url, search_term):
    """Yönetici sayfasından verileri çeker"""
    try:
        print(f"Visiting: {officer_url}")
        page.goto(officer_url)
        random_delay()
        
        # Yönetici adı
        officer_name = page.locator(".heading-xlarge").inner_text().strip()
        print(f"Officer name: {officer_name}")
        
        # Atamalar
        appointments = []
        appointment_elements = page.locator(".appointment-1").all()
        print(f"Found {len(appointment_elements)} appointments")
        
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
                "Atamalar": appointments,
                "URL": officer_url
            }
        else:
            print(f"No appointments found for {officer_name}")
            return None
    
    except Exception as e:
        print(f"Error extracting officer data: {str(e)}")
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
            print(f"Searching page {page_num}: {search_url}")
            
            page.goto(search_url)
            random_delay()
            
            # Sonuç var mı kontrol et
            no_results = page.locator(".search-no-results").count() > 0
            if no_results:
                print(f"No results found for {search_term} on page {page_num}")
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
                print(f"No officer links found on page {page_num}")
                break
            
            print(f"Found {len(links)} officers on page {page_num}")
            
            # Her yönetici için veri çek
            for i, link in enumerate(links):
                try:
                    print(f"Processing officer {i+1}/{len(links)} on page {page_num}")
                    
                    href = link.get('href')
                    if href:
                        officer_url = base_url + href
                        officer_data = extract_officer_data(page, officer_url, search_term)
                        
                        if officer_data:
                            officers_data.append(officer_data)
                            print(f"Successfully processed: {officer_data['İsim']}")
                    
                    random_delay()
                except Exception as e:
                    print(f"Error processing officer link: {str(e)}")
                    continue
            
            # Sonraki sayfa var mı?
            next_button = page.locator("a.page-next")
            if next_button.count() == 0 or not next_button.is_visible():
                print(f"No more pages for {search_term}")
                break
    
    except Exception as e:
        print(f"Error during search: {str(e)}")
    
    finally:
        page.close()
    
    return officers_data

def save_to_excel(all_data, filename="turkish_officers"):
    """Verileri Excel dosyasına kaydeder"""
    if not all_data:
        print("No data to save")
        return
    
    # Tüm atamalar için liste
    all_appointments = []
    
    # Yönetici özet bilgileri için liste
    officers_summary = []
    
    for officer in all_data:
        # Yönetici özet bilgisi
        officers_summary.append({
            "Arama Terimi": officer["Arama Terimi"],
            "İsim": officer["İsim"],
            "Atama Sayısı": len(officer["Atamalar"]),
            "URL": officer["URL"]
        })
        
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
            all_appointments.append(appointment_data)
    
    # Excel dosyası oluştur
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    excel_filename = f"{filename}_{timestamp}.xlsx"
    
    with pd.ExcelWriter(excel_filename, engine='openpyxl') as writer:
        # Yöneticiler sayfası
        pd.DataFrame(officers_summary).to_excel(writer, sheet_name='Yöneticiler', index=False)
        
        # Atamalar sayfası
        pd.DataFrame(all_appointments).to_excel(writer, sheet_name='Atamalar', index=False)
    
    print(f"Data saved to {excel_filename}")
    print(f"Total: {len(officers_summary)} officers and {len(all_appointments)} appointments")

def main():
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
    
    print(f"Daha önce işlenmiş isimler: {processed_names}")
    
    # İşlenecek isimleri belirle
    names_to_process = [name for name in all_names if name not in processed_names]
    print(f"İşlenecek isim sayısı: {len(names_to_process)}")
    
    # Daha önce toplanan verileri yükle
    all_data = []
    for name in processed_names:
        json_filename = f"results_{name}.json"
        try:
            with open(json_filename, 'r', encoding='utf-8') as f:
                data = json.load(f)
                all_data.extend(data)
                print(f"{name} için {len(data)} yönetici verisi yüklendi")
        except Exception as e:
            print(f"Hata: {name} verisi yüklenemedi - {str(e)}")
    
    # Excel dosyasına kaydet
    if all_data:
        print("Mevcut verileri Excel'e kaydediyorum...")
        save_to_excel(all_data, "turkish_officers_current")
        print(f"Toplam {len(all_data)} yönetici verisi Excel'e kaydedildi")
    
    # Otomatik olarak devam et
    print("Veri toplamaya devam ediliyor...")
    
    # Tüm veriler
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)  # Görünür tarayıcı
        context = browser.new_context(
            viewport={'width': 1920, 'height': 1080},
            user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36'
        )
        
        # İşlenecek isimleri işle
        for i, name in enumerate(names_to_process):
            print(f"\n{'='*50}")
            print(f"Processing name {i+1}/{len(names_to_process)}: {name}")
            print(f"{'='*50}")
            
            officers_data = process_name(context, name)
            
            if officers_data:
                all_data.extend(officers_data)
                print(f"Found {len(officers_data)} officers for {name}")
                
                # JSON dosyasına kaydet
                json_filename = f"results_{name}.json"
                with open(json_filename, 'w', encoding='utf-8') as f:
                    json.dump(officers_data, f, ensure_ascii=False, indent=4)
                
                print(f"Data for {name} saved to {json_filename}")
                
                # Her 5 isimde bir ara sonuçları kaydet
                if (i + 1) % 5 == 0 and all_data:
                    interim_filename = f"turkish_officers_interim_{i+1}_names"
                    save_to_excel(all_data, interim_filename)
            else:
                print(f"No results found for {name}")
            
            # Kısa bir ara ver
            time.sleep(random.uniform(2, 3))
        
        browser.close()
    
    # Tüm verileri kaydet
    if all_data:
        save_to_excel(all_data, "turkish_officers_final")
    else:
        print("No data found for any name")

if __name__ == "__main__":
    main()
