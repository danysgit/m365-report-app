FROM python:3.12-slim
WORKDIR /app
RUN apt-get update && apt-get install -y --no-install-recommends \
    libfreetype6-dev libpng-dev && rm -rf /var/lib/apt/lists/*
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .
RUN mkdir -p uploads
EXPOSE 5000
ENV FLASK_APP=app.py
ENV SECRET_KEY=change-me-in-production
CMD ["python", "app.py"]
