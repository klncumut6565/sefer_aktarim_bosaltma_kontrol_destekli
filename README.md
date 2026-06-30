# Sefer Aktarım — Boşaltma Kontrol Destekli (Web Sürümü)

PDF sefer bildirimlerini okuyup Excel taşıma kontrol listesine aktarır ve
isteğe bağlı olarak her sefer için Boşaltma Kontrol Dökümanı (Word/PDF) üretir.

## Çalıştırma

```bash
pip install -r requirements.txt
streamlit run app.py
```

DOCX → PDF dönüşümü için LibreOffice gereklidir (Streamlit Cloud'da
`packages.txt` ile otomatik kurulur; yerelde `apt install libreoffice` /
`brew install libreoffice` gibi bir yolla kurulabilir).
