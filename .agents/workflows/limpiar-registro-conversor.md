---
description: Limpiar copias antiguas o fantasma de ConversorPDF en el registro de Windows
---

# Limpieza del Registro de ConversorPDF

Este workflow está diseñado para ejecutarse en entornos de desarrollo donde el instalador MSI pudo haber quedado en un estado inconsistente (instalaciones fantasma o múltiples versiones en paralelo). 

Esto ocurre típicamente cuando se hacen pruebas iniciales con `UpgradeCode` variables, o combinando instalaciones `perUser` con `perMachine`. Cuando esto sucede, el Panel de Control de Windows (Programas y Características) muestra fechas o versiones incorrectas porque agrupa las entradas antiguas.

---

## Pasos

### 1. Ejecutar desinstalación forzada por WMI

Este comando busca en la base de datos de Windows Installer (`Win32_Product`) cualquier producto cuyo nombre coincida con "Conversor" y fuerza su desinstalación. 

> [!WARNING]
> La consulta `Win32_Product` puede ser lenta (tarda algunos minutos en ejecutarse por completo). Por favor, tener paciencia mientras el script finaliza.

// turbo
2. Ejecuta el comando en PowerShell:
```powershell
Get-WmiObject -Class Win32_Product | Where-Object { $_.Name -match "Conversor" } | ForEach-Object { 
    Write-Host "Desinstalando $($_.Name) versión $($_.Version)..."
    $_.Uninstall() 
}
Write-Host "Limpieza completada."
```

### 3. Verificación adicional

Si la entrada no desaparece del Panel de Control, puede ser un remanente en el registro del usuario.

// turbo
4. Ejecuta este comando secundario para limpiar claves huérfanas en el registro del usuario actual (opcional pero recomendado):
```powershell
Get-ChildItem -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall" -Recurse -ErrorAction SilentlyContinue | Get-ItemProperty | Where-Object { $_.DisplayName -match "Conversor" } | ForEach-Object {
    Write-Host "Borrando clave de registro huérfana de $($_.DisplayName)..."
    Remove-Item -Path $_.PSPath -Recurse -Force
}
```

### 5. Limpieza de instalaciones globales (Opcional pero crítico)

Si la versión remanente era `1.1.2.0` o inferior, es probable que haya sido instalada **para todos los usuarios** (nivel Máquina `perMachine`). Estos registros no se pueden borrar sin privilegios.

Si recibes un error `1603` o acceso denegado en los pasos anteriores:
1. Abre **PowerShell como Administrador**
2. Ejecuta el siguiente comando para limpiar las rutas del sistema central:

```powershell
Get-ChildItem -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall" -Recurse -ErrorAction SilentlyContinue | Get-ItemProperty | Where-Object { $_.DisplayName -match "Conversor" } | ForEach-Object {
    Write-Host "Borrando clave global de $($_.DisplayName)..."
    Remove-Item -Path $_.PSPath -Recurse -Force
}
```

### 6. Reinstalar la versión oficial

Una vez que el entorno esté completamente limpio y sin rastros de versiones anteriores en "Programas y características", puedes usar el script de build principal o tu último instalador estable para volver a instalar el software:

```powershell
powershell -ExecutionPolicy Bypass -File ".\scripts\build_installer.ps1"
```

El Panel de Control y Excel reflejarán ahora exclusivamente esta versión limpia y única.
