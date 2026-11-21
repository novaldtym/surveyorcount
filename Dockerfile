FROM python:3.11-slim

WORKDIR /app

# Salin dependensi dan instal
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Salin seluruh kode proyek
COPY . .

# Buat folder uploads dengan permission yang benar
RUN mkdir -p uploads && chmod -R 777 uploads

# Port default untuk Hugging Face Spaces adalah 7860, atau Render via $PORT
ENV PORT=7860
EXPOSE 7860

CMD ["gunicorn", "-b", "0.0.0.0:7860", "--workers", "2", "--timeout", "120", "app:app"]
