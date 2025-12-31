# Usamos una imagen base de Python liviana
FROM python:3.9-slim

# Evita que Python genere archivos .pyc y mantiene los logs en tiempo real
ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1

# Directorio de trabajo en el contenedor
WORKDIR /app

# Copiamos los requerimientos e instalamos
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copiamos el resto del código
COPY . .

# Exponemos el puerto (Coolify usa el 8000 u 80 por defecto a veces, pero 8000 es seguro)
EXPOSE 8000

# COMANDO DE INICIO:
# Si es un script que corre una vez y termina (o un bucle infinito simple):
CMD ["python", "main.py"]

# O SI ES UNA APP FLASK/WEB (descomenta la linea de abajo y comenta la de arriba):
# CMD ["gunicorn", "-w", "2", "-b", "0.0.0.0:8000", "main:app"]
