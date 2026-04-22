---
description: Build and release the ConversorPDF installer (WiX 4, Release|x64)
---

# Build & Release — ConversorPDF Installer

Este workflow cubre todos los pasos para generar un instalador profesional, estable y con auto-activación en Excel.

---

## Pre-requisitos

- SDK de .NET instalado y accesible desde PowerShell.
- WiX Toolset 4 instalado (`dotnet tool list -g` debe mostrar `wix`).
- El proyecto compila en modo `Release|x64` sin errores de tipo `CS`.

---

## Pasos

### 1. Determinar la nueva versión

Decidir el número de versión en formato `Major.Minor.Patch.Build` (ej: `1.2.45.10`).

### 2. Ejecutar el script maestro

El script maestro automatiza el flujo completo: limpia artefactos anteriores, sincroniza la nueva versión en todos los archivos del proyecto, ejecuta **todos los tests unitarios** para prevenir regresiones, y finalmente compila el XLL y empaqueta el MSI.

Abre una terminal en la **raíz del proyecto** y ejecuta:

// turbo
```powershell
powershell -ExecutionPolicy Bypass -File ".\scripts\build_installer.ps1" -NewVersion "LA.NUEVA.VERSION"
```

> **Notas técnicas WiX 4:**
> - No colocar archivos `.wxs` adicionales en la raíz de `installer/`, ya que WiX 4 hace globbing automático y genera conflictos de símbolos.
> - El archivo de destino en `AppData` debe llamarse `ConversorPDF.xll` para consistencia con el Registro de Windows.
> - La Custom Action `RegisterExcelXLL` debe tener la condición `NOT REMOVE` para ejecutarse en instalaciones nuevas **y** en actualizaciones (MajorUpgrade).

---

### ⚠️ Lección Aprendida 1 — Microsoft.Bcl.HashCode.dll (PackExcelAddIn ERROR)

**Síntoma:**
```
PackExcelAddIn error: Reference with Path: Microsoft.Bcl.HashCode.dll and Name: NOT FOUND
```

**Causa raíz:**
`Microsoft.Bcl.HashCode.dll` es un polyfill de .NET Standard incluido como dependencia transitiva de `PdfPig` (via `System.Reflection.Metadata`). En **.NET Framework 4.8**, `HashCode` está integrado en el runtime y **este archivo DLL nunca se genera** en `bin\x64\Release\`.

**Solución permanente ya aplicada:**
El archivo `ConversorPDF.XLL-AddIn64.dna` **NO debe** contener:
```xml
<Reference Path="Microsoft.Bcl.HashCode.dll" Pack="true" />
```

**Regla general para el futuro:**
Antes de agregar cualquier `<Reference>` al `.dna`, verificar que el archivo existe en `bin\x64\Release\`:
```powershell
Test-Path "ConversorPDF.XLL\bin\x64\Release\NombreAssembly.dll"
```
Si devuelve `False`, **no agregar** esa referencia al `.dna`.

### 6. Verificación del artefacto

Confirmar que se generó:

```
installer\InstaladorConversorPDFv<Version>.msi
```

Verificar metadatos del MSI con `msiinfo` o el explorador de propiedades de Windows.

> **Advertencias WiX esperadas (no críticas, ignorar):**
> - **ICE61** — `AllowSameVersionUpgrades="yes"`: comportamiento intencional.
> - **ICE91** — XLL en `AddInsFolder` de usuario: correcto para complementos Excel por usuario.

### 7. Smoke test manual

1. Instalar el MSI en una máquina limpia (o VM).
2. Abrir Excel y confirmar que el complemento ConversorPDF aparece activo.
3. Ejecutar una conversión de prueba (Individual, Masivo, Cumplidos).
4. Desinstalar y verificar que no quedan residuos en `AppData` ni en el Registro.
