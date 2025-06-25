@echo off
setlocal enabledelayedexpansion

:: Obtener el directorio desde donde se ejecutó el script
set "SCRIPT_DIR=%~dp0"
cd /d "%SCRIPT_DIR%"

:: Configurar color verde
color 0a

:: Verificar privilegios de administrador
echo Verificando privilegios de administrador...
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Este script requiere privilegios de administrador.
    echo Por favor, ejecute el script como administrador.
    pause
    exit /b 1
)
echo [OK] Ejecutando como administrador

:: Verificar conexion a internet
echo Verificando conexion a internet...
ping 8.8.8.8 -n 1 -w 1000 >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: No se detecto conexion a internet.
    echo Por favor, verifique su conexion e intente nuevamente.
    pause
    exit /b 1
)
echo [OK] Conexion a internet detectada

:: Verificar si Python esta instalado
echo Verificando instalacion de Python...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Python no esta instalado. Por favor, instale Python 3.8 o superior.
    pause
    exit /b 1
)
echo [OK] Python instalado correctamente

:: Crear entorno virtual
echo Creando entorno virtual...
if exist venv (
    echo Eliminando entorno virtual anterior...
    rmdir /s /q venv
)
python -m venv venv
if %errorlevel% neq 0 (
    echo ERROR: No se pudo crear el entorno virtual.
    pause
    exit /b 1
)
echo [OK] Entorno virtual creado

:: Activar entorno virtual
echo Activando entorno virtual...
call venv\Scripts\activate
if %errorlevel% neq 0 (
    echo ERROR: No se pudo activar el entorno virtual.
    pause
    exit /b 1
)
echo [OK] Entorno virtual activado

:: Verificar pip en el entorno virtual
echo Verificando pip en el entorno virtual...
pip --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: No se pudo verificar pip en el entorno virtual.
    pause
    exit /b 1
)
echo [OK] Pip verificado

:: Actualizar pip
echo Actualizando pip...
python -m pip install --upgrade pip
if %errorlevel% neq 0 (
    echo ERROR: No se pudo actualizar pip.
    pause
    exit /b 1
)
echo [OK] Pip actualizado

:: Verificar y crear requirements.txt si no existe
if not exist requirements.txt (
    echo Creando archivo requirements.txt...
    (
        echo pandas>=1.5.3
        echo selenium>=4.9.0
        echo webdriver-manager>=3.8.6
        echo python-dotenv>=1.0.0
        echo PyPDF2>=3.0.0
        echo openpyxl>=3.1.2
        echo pywin32>=306
    ) > requirements.txt
    echo [OK] Archivo requirements.txt creado
) else (
    echo [OK] Archivo requirements.txt ya existe
)

:: Instalar dependencias
echo Instalando dependencias desde %SCRIPT_DIR%requirements.txt...
color 0a
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo ERROR: No se pudieron instalar las dependencias.
    echo Verificar que el archivo requirements.txt existe en: %SCRIPT_DIR%
    pause
    exit /b 1
)
color 0a
echo [OK] Dependencias instaladas

:: Crear archivo .env si no existe
color 0a
if not exist .env (
    echo Creando archivo .env desde .env.example...
    copy .env.example .env >nul
    echo [OK] Archivo .env creado
    echo IMPORTANTE: Por favor, edite el archivo .env y configure sus credenciales
) else (
    echo [OK] Archivo .env ya existe
)

:: Crear directorio para PDFs
color 0a
echo Creando directorio para PDFs...
if not exist downloads mkdir downloads
echo [OK] Directorio para PDFs creado

color 0a
echo.
echo Configuracion completada exitosamente!
echo Para ejecutar el script principal:
echo 1. Edite el archivo .env y configure sus credenciales
echo 2. Use run.cmd para ejecutar el script
echo.
pause 