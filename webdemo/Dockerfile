FROM python:3.11-slim

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY app.py .
COPY scripts ./scripts
COPY assets ./assets

RUN mkdir -p uploads salida

EXPOSE 10000

# Render inyecta la variable de entorno $PORT en tiempo de ejecución -- usamos shell form
# para que se expanda ese valor real (no un puerto fijo).
CMD gunicorn -b 0.0.0.0:${PORT:-10000} --timeout 120 app:app
