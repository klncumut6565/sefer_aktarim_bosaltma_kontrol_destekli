FROM python:3.12-slim

# LibreOffice: Gönderim/Boşaltma Kontrol Dökümanlarının .docx -> .pdf dönüşümü için gerekli.
# apt-get update her build'de taze çalışır; Streamlit Cloud'daki gibi bozuk/eskimiş
# önbelleğe bağımlı değiliz.
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# Render, PORT ortam değişkenini dışarıdan atar; Streamlit'i buna bağlıyoruz.
ENV STREAMLIT_SERVER_HEADLESS=true \
    STREAMLIT_SERVER_ENABLE_CORS=false \
    STREAMLIT_SERVER_ENABLE_XSRF_PROTECTION=false \
    STREAMLIT_BROWSER_GATHER_USAGE_STATS=false

EXPOSE 8501

CMD streamlit run app.py --server.port=${PORT:-8501} --server.address=0.0.0.0
