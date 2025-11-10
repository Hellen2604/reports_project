#!/usr/bin/env bash
set -ex

# Instala dependencias del sistema
apt-get update
apt-get install -y curl gnupg2 apt-transport-https software-properties-common

# Agrega el repositorio de Microsoft
curl https://packages.microsoft.com/keys/microsoft.asc | apt-key add -
curl https://packages.microsoft.com/config/debian/11/prod.list > /etc/apt/sources.list.d/mssql-release.list

# Instala el driver ODBC y dependencias
apt-get update
ACCEPT_EULA=Y apt-get install -y msodbcsql18 unixodbc-dev

# Limpieza opcional
apt-get clean
