import pandas as pd
import json
import os
import glob
from datetime import datetime

def main():
    # Tüm JSON dosyalarını bul
    json_files = glob.glob('results_*.json')
    print(f"Bulunan JSON dosyası sayısı: {len(json_files)}")
    
    # Tüm verileri topla
    all_data = []
    for json_file in json_files:
        try:
            with open(json_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
                if isinstance(data, list):
                    all_data.extend(data)
                else:
                    all_data.append(data)
                print(f"Yüklendi: {json_file}")
        except Exception as e:
            print(f"Hata: {json_file} yüklenemedi - {str(e)}")
    
    print(f"Toplam yüklenen veri sayısı: {len(all_data)}")
    
    # Yönetici özet bilgileri ve atamalar
    officers_summary = []
    all_appointments = []
    
    for officer in all_data:
        if not isinstance(officer, dict):
            continue
        
        if 'İsim' not in officer or 'Atamalar' not in officer:
            continue
        
        # Yönetici özet bilgisi
        officers_summary.append({
            "Arama Terimi": officer.get('Arama Terimi', ''),
            "İsim": officer.get('İsim', ''),
            "Atama Sayısı": len(officer.get('Atamalar', [])),
            "URL": officer.get('URL', '')
        })
        
        # Atama bilgileri
        for appointment in officer.get('Atamalar', []):
            if not isinstance(appointment, dict):
                continue
                
            appointment_data = {
                "Arama Terimi": officer.get('Arama Terimi', ''),
                "Yönetici İsmi": officer.get('İsim', ''),
                "Şirket Adı": appointment.get('Şirket Adı', ''),
                "Şirket Numarası": appointment.get('Şirket Numarası', ''),
                "Şirket Durumu": appointment.get('Şirket Durumu', ''),
                "Rol": appointment.get('Rol', ''),
                "Yazışma Adresi": appointment.get('Yazışma Adresi', ''),
                "Atanma Tarihi": appointment.get('Atanma Tarihi', ''),
                "Yönetilen Kanun": appointment.get('Yönetilen Kanun', ''),
                "Yasal Form": appointment.get('Yasal Form', '')
            }
            all_appointments.append(appointment_data)
    
    print(f"Yönetici sayısı: {len(officers_summary)}")
    print(f"Atama sayısı: {len(all_appointments)}")
    
    # Excel dosyası oluştur
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    excel_filename = f"turkish_officers_combined_{timestamp}.xlsx"
    
    with pd.ExcelWriter(excel_filename, engine='openpyxl') as writer:
        # Yöneticiler sayfası
        pd.DataFrame(officers_summary).to_excel(writer, sheet_name='Yöneticiler', index=False)
        
        # Atamalar sayfası
        pd.DataFrame(all_appointments).to_excel(writer, sheet_name='Atamalar', index=False)
    
    print(f"Veriler kaydedildi: {excel_filename}")
    print(f"Toplam: {len(officers_summary)} yönetici ve {len(all_appointments)} atama")

if __name__ == "__main__":
    main()
