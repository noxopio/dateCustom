@echo off
REM install.bat - Instalador de dateCustom
REM Copia todos los archivos a la carpeta del usuario y crea el acceso directo
setlocal enabledelayedexpansion

echo.
echo ========================================
echo   Instalador de dateCustom
echo ========================================
echo.

REM Detectar la carpeta de origen (donde está este install.bat)
set "SOURCE=%~dp0"
set "SOURCE=%SOURCE:~0,-1%"

REM Carpeta de destino en la raíz del usuario
set "DEST=%USERPROFILE%\dateCustom"

echo [1/3] Creando carpeta de destino...
if not exist "%DEST%" (
    mkdir "%DEST%"
    echo      Carpeta creada: %DEST%
) else (
    echo      Carpeta ya existe: %DEST%
)

echo.
echo [2/3] Copiando archivos...
xcopy "%SOURCE%\*.*" "%DEST%\" /Y /Q /EXCLUDE:%SOURCE%\install.bat
if errorlevel 1 (
    echo      Error al copiar archivos.
    pause
    exit /b 1
)
echo      Archivos copiados correctamente.

echo.
echo [3/3] Creando acceso directo en el Escritorio...
powershell -NoProfile -ExecutionPolicy Bypass -File "%DEST%\create_shortcut.ps1"
if errorlevel 1 (
    echo      Error al crear acceso directo.
    pause
    exit /b 1
)

echo.
echo ========================================
echo   Instalacion completada
echo ========================================
echo.
echo - Archivos instalados en: %DEST%
echo - Acceso directo creado en el Escritorio
echo.
echo Para usar el script:
echo   1) Doble clic en "setDate - Ejecutar rapido" en el Escritorio
echo   2) O ejecuta: %DEST%\setDate.bat archivo.xlsx 202511191200
echo.
pause
