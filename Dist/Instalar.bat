@echo off
setlocal

:: Verificar permisos de administrador
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo Error: Este script debe ejecutarse como Administrador.
    echo Por favor, haga clic derecho en este archivo y seleccione "Ejecutar como administrador".
    pause
    exit /b 1
)

:: Localizar RegAsm de 64 bits
set "REGASM=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe"

if not exist "%REGASM%" (
    echo Error: No se encontro RegAsm de 64 bits en:
    echo %REGASM%
    echo Asegurese de tener instalado .NET Framework 4.x.
    pause
    exit /b 1
)

:: Ruta de la DLL (se asume en la misma carpeta que este script)
set "DLL_PATH=%~dp0ConversorPDF.dll"

if not exist "%DLL_PATH%" (
    echo Error: No se encontro la DLL en:
    echo %DLL_PATH%
    pause
    exit /b 1
)

echo Registrando ConversorPDF.dll...
"%REGASM%" "%DLL_PATH%" /codebase /tlb

if %errorLevel% equ 0 (
    echo.
    echo ==========================================
    echo    REGISTRO COMPLETADO EXITOSAMENTE
    echo ==========================================
    echo.
) else (
    echo.
    echo ==========================================
    echo    OCURRIO UN ERROR DURANTE EL REGISTRO
    echo ==========================================
    echo.
)

pause
