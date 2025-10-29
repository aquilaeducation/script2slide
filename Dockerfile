FROM python:3.11-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1

WORKDIR /app
COPY requirements.txt /app/
RUN pip install --no-cache-dir -r requirements.txt

COPY . /app

# Informational on Render; Render will still inject $PORT
EXPOSE 8080

# Start Gunicorn bound to $PORT from the host
CMD ["python", "-m", "gunicorn", "-w", "2", "-b", "0.0.0.0:$PORT", "--access-logfile", "-", "--timeout", "60", "app:app"]
