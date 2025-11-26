FROM python:3.11-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1

WORKDIR /app

# copy requirements first for better caching
COPY requirements.txt ./

RUN pip install --no-cache-dir -r requirements.txt

# copy application
COPY . /app

EXPOSE 5000

# Use gunicorn for production-like server
CMD ["gunicorn", "-w", "4", "-b", "0.0.0.0:5000", "app:app"]
