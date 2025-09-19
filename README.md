# Excel Converter & Alarm Reporting Tool

This project is an **Excel conversion and reporting tool** designed for office use.  
Users can upload files through a browser, filter alarm reports, process them, and easily download the formatted Excel files.

---

## Features
- Upload alarm reports and device information  
- Filter based on user-specified time intervals  
- Special filtering for driver fatigue alerts  
- Filter alarms repeating every 3 hours  
- Limit daily alarms to a maximum of 3  
- Translate alarm types to Turkish and organize columns  
- Generate Excel links according to alarm time  
- Download results as a formatted Excel file  

---

## Usage
🔗 [You can try the application here](https://excel-otomasyon-sistemdestek.streamlit.app)  

1. Open the application in a browser  
2. Upload the **Alarm Report file** (`.xlsx`)  
3. Upload the **Device Information file** (`.xlsx`)  
4. Enter the **name of the output file**  
5. Specify the **time intervals** (e.g., `05.00-08.20`, `16.40-20.40`)  
6. Click the **"Process Report/Raporu İşle"** button  
7. Once processing is complete, download the **Excel file**

   To install the required dependencies:
```bash
pip install -r requirements.txt
```

# Excel Dönüştürücü & Alarm Raporlama Aracı

Bu proje, ofis kullanımına uygun olarak hazırlanmış bir **Excel dönüştürme ve raporlama aracı**dır.  
Kullanıcılar, tarayıcı üzerinden dosya yükleyip alarm raporlarını filtreleyebilir, işleyebilir ve düzenlenmiş Excel dosyalarını kolayca indirebilir.  

---

## Özellikler
- Alarm raporlarını ve cihaz bilgilerini yükleme  
- Kullanıcı tarafından belirlenen saat aralıklarına göre filtreleme  
- Sürücü esneme uyarılarını özel filtreleme  
- 3 saat aralıklı tekrar eden alarmların filtrelenmesi  
- Günlük maksimum 3 alarm sınırlaması  
- Alarm türlerini Türkçeye çevirme ve sütun düzenleme  
- Alarm zamanına göre Excel linkleri oluşturma  
- Sonuçları biçimlendirilmiş Excel dosyası olarak indirme  

---

## Kullanım
🔗 [Uygulamayı buradan deneyebilirsiniz](https://excel-otomasyon-sistemdestek.streamlit.app)  

1. Tarayıcı üzerinden uygulamayı açın 
2. **Alarm rapor dosyasını** yükleyin (`.xlsx`)  
3. **Cihaz bilgileri dosyasını** yükleyin (`.xlsx`)  
4. **Oluşturulacak dosya adını** girin  
5. **Saat aralıklarını** belirtin (örn: `05.00-08.20`, `16.40-20.40`)  
6. **"Raporu İşle"** butonuna tıklayın  
7. İşlem tamamlandığında **Excel dosyasını indirin**

   Gerekli bağımlılıkları yüklemek için:
```bash
pip install -r requirements.txt
```
