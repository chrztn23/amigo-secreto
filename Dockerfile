FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    XDG_RUNTIME_DIR=/app/run

WORKDIR /app
COPY requirements.txt ./
RUN python -m pip install --no-cache-dir -r requirements.txt

COPY app.py draw.py ./
COPY templates/ ./templates/
COPY static/ ./static/

RUN mkdir -p /app/data /app/instance \
    && chown 10001:10001 /app/data /app/instance \
    && install -d -m 0700 -o 10001 -g 10001 /app/run
USER 10001:10001

EXPOSE 8000
CMD ["gunicorn", "--bind", "0.0.0.0:8000", "app:app"]
