@echo off
setlocal enabledelayedexpansion

:: Obtener el directorio desde donde se ejecutó el script
set "SCRIPT_DIR=%~dp0"
cd /d "%SCRIPT_DIR%"

:: Configurar color verde
color 0a

:: Verificar si el entorno virtual existe
if not exist venv (
    echo ERROR: El entorno virtual no existe.
    echo Por favor, ejecute primero auto_setup.cmd
    pause
    exit /b 1
)

:: Activar entorno virtual
echo Activando entorno virtual...
call venv\Scripts\activate
if %errorlevel% neq 0 (
    echo ERROR: No se pudo activar el entorno virtual.
    pause
    exit /b 1
)
echo [OK] Entorno virtual activado

:: Crear archivo de log para main.py
echo Iniciando ejecucion de main.py...
echo Iniciando ejecucion de main.py... > setup_log.txt

:: Ejecutar el script principal
echo Ejecutando el script principal desde %SCRIPT_DIR%...
python "%SCRIPT_DIR%main.py" >> setup_log.txt 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Hubo un problema al ejecutar el script principal.
    echo Para mas detalles, revise el archivo setup_log.txt
    pause
    exit /b 1
)

echo Script completado exitosamente!
echo Los detalles de la ejecucion se encuentran en setup_log.txt
pause 