FROM python:3.12-slim

# System libs required by matplotlib
RUN apt-get update && apt-get install -y --no-install-recommends \
    libfreetype6 \
    libpng16-16 \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt gunicorn

COPY app.py .
COPY templates/ templates/
COPY static/ static/

RUN mkdir -p uploads

# Unraid Docker Manager labels — enables WebUI button and icon in Unraid UI
LABEL net.unraid.docker.webui="http://[IP]:[PORT:5000]/"
LABEL net.unraid.docker.icon="https://raw.githubusercontent.com/danysgit/m365-report-app/main/static/favicon.png"
LABEL net.unraid.docker.managed="dockerman"

EXPOSE 5000
VOLUME ["/app/uploads"]

CMD ["gunicorn", "--bind", "0.0.0.0:5000", "--workers", "2", "--timeout", "120", "app:app"]
