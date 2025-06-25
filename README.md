# Automatización de Búsqueda de COA Mindray

Este proyecto automatiza el proceso de:
1. Inicio de sesión en el portal de Mindray
2. Búsqueda de COAs basada en datos de Excel
3. Descarga de reportes PDF
4. Extracción de fechas de expiración
5. Actualización del archivo Excel

## Guía de Instalación y Ejecución Paso a Paso

### 1. Instalación de Python y google Chrome
1. Descargar Python 3.11 o posterior desde la página oficial:
   - Visitar https://www.python.org/downloads/
   - Descargar "Python 3.11.x" o posterior para Windows
   - **IMPORTANTE**: Durante la instalación, marcar la casilla "Add Python to PATH"
   - Hacer clic en "Install Now"

### 2. Preparación del Proyecto
1. Crear el archivo de configuración:
   - Abrir `.env` con el Bloc de notas
   - Completar los siguientes datos:
     ```
     WEBSITE_USERNAME=tu_usuario_mindray
     WEBSITE_PASSWORD=tu_contraseña_mindray
     EXCEL_FILE=SEGUIMIENTO DE INGRESO A ALMACÉN Y ORDENES DE TRABAJO.xlsx
     EXCEL_SHEET=SIGNIA Y PRODIS
     ```
   - Guardar y cerrar el archivo

### 3. Ejecución de auto_setup.cmd
1. Hacer clic derecho en el archivo `auto_setup.cmd`
2. Seleccionar "Ejecutar como administrador"
3. Si aparece una advertencia de seguridad:
   - Presionar "Más información"
   - Hacer clic en "Ejecutar de todas formas"
4. **IMPORTANTE**: No cerrar la ventana de cmd durante la instalación
5. Esperar hasta que aparezca el mensaje "Instalación completada"
6. Presionar cualquier tecla para cerrar la ventana

### 4. Ejecución del Programa (run.cmd)
1. Hacer doble clic en `run.cmd`
2. Se abrirán dos ventanas:
   - Una ventana de Chrome (controlada por el programa)
   - Una ventana de consola negra
3. **IMPORTANTE**: 
   - No cerrar ninguna de las ventanas
   - No interactuar con el navegador Chrome
   - No abrir ni modificar el archivo Excel durante la ejecución
4. El programa terminará cuando:
   - La ventana de Chrome se cierre automáticamente
   - En la consola negra aparezca "Presione cualquier tecla para continuar..."
5. Una vez finalizado, presionar cualquier tecla para cerrar la consola

### 5. Verificación de Resultados
1. Los PDFs descargados se encontrarán en la carpeta `downloads/pdfs/`
2. El archivo Excel habrá sido actualizado con las nuevas fechas
3. Se puede revisar el archivo `setup_log.txt` para ver el registro de la ejecución

## Solución de Problemas Comunes

### Error al ejecutar auto_setup.cmd
1. Abrir PowerShell como administrador
2. Ejecutar el comando:
   ```powershell
   Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
   ```
3. Escribir "S" y presionar Enter
4. Intentar ejecutar auto_setup.cmd nuevamente

### El programa se cierra inmediatamente
1. Verificar que Python esté instalado correctamente
2. Verificar que el archivo `.env` existe y está configurado
3. Asegurarse de que el archivo Excel esté cerrado
4. Ejecutar run.cmd como administrador

### Error de "Chrome no encontrado"
1. Instalar Google Chrome desde https://www.google.com/chrome/
2. Reiniciar el equipo
3. Intentar ejecutar run.cmd nuevamente

## Notas Importantes
- No cerrar las ventanas del programa hasta que termine la ejecución
- Mantener una conexión a internet estable durante todo el proceso
- Hacer una copia de seguridad del archivo Excel antes de ejecutar el programa
- El tiempo de ejecución dependerá de la cantidad de registros a procesar
- Si el programa se interrumpe, se puede volver a ejecutar run.cmd para continuar

## Requisitos del Sistema

- Windows 10 o superior
- Permisos de administrador
- Conexión a internet
- Google Chrome instalado

## Estructura del Excel

El archivo Excel debe estar configurado así:
- Nombre: "SEGUIMIENTO_ALMACEN.xlsx"
- Hoja: "SIGNIA Y PRODIS"
- Columnas necesarias:
  - F: Código de material (se omiten los primeros 3 caracteres)
  - H: Número de lote
  - Expiry Date: Se actualizará con la fecha extraída

Ejemplo:
```
Columna F          | Columna H    | Expiry Date
MR 105-007379-00  | 2024070111   | 2025-08-13
```

## Mantenimiento

- Los PDFs se guardan en `downloads/pdfs/`
- El log de ejecución se guarda en `setup_log.txt`
- Hacer backup del Excel antes de ejecutar
- No modificar el Excel durante la ejecución

## Seguridad

- No compartir el archivo .env
- No subir credenciales a repositorios
- Mantener el Excel respaldado
- Usar contraseñas seguras

## Notas Importantes

- El script procesa solo la hoja "SIGNIA Y PRODIS"
- Se omiten los primeros 3 caracteres del código en columna F
- Los PDFs se guardan con nombre basado en el código y lote
- El proceso puede tomar tiempo según la cantidad de registros 