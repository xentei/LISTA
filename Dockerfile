FROM python:3.9-slim

# 1. Configuraciones de entorno
ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1

WORKDIR /app

# 2. INSTALAMOS CURL (ESTO FALTABA PARA EL HEALTHCHECK)
# Actualizamos listas, instalamos curl y limpiamos para no ocupar espacio
RUN apt-get update && apt-get install -y curl && rm -rf /var/lib/apt/lists/*

# 3. Copiamos e instalamos requerimientos
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# 4. Copiamos el resto del código
COPY . .

# 5. Exponemos el puerto de Streamlit
EXPOSE 8501

# 6. Chequeo de salud (Ahora sí funcionará porque instalamos curl)
HEALTHCHECK CMD curl --fail http://localhost:8501/_stcore/health || exit 1

# 7. Comando de inicio
CMD ["streamlit", "run", "main.py", "--server.port=8501", "--server.address=0.0.0.0"]
