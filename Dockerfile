FROM python:3.11-slim

# Instalar dependencias del sistema
RUN apt-get update && \
    apt-get install -y curl gnupg2 apt-transport-https software-properties-common && \
    curl https://packages.microsoft.com/keys/microsoft.asc | apt-key add - && \
    curl https://packages.microsoft.com/config/debian/11/prod.list > /etc/apt/sources.list.d/mssql-release.list && \
    ACCEPT_EULA=Y apt-get install -y msodbcsql18 unixodbc-dev

# Crear carpeta de la app
WORKDIR /app
COPY . /app

# Instalar dependencias de Python
RUN pip install --no-cache-dir -r requirements.txt

# Puerto y comando
EXPOSE 5000
CMD ["gunicorn", "app:app"]
