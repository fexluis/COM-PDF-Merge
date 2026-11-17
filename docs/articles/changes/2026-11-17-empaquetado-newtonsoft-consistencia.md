# Empaquetado del add-in: Newtonsoft.Json declarado y contrato de consistencia

## Propósito

El escalador de ranuras clona el contexto con `Newtonsoft.Json` (`AgEscaladorRanuraAula.CopiarModelo`).
Al habilitar la planificación de aulas `PendienteDeProgramar`, el flujo llega por primera vez a ese punto y
el add-in instalado fallaba con `FileNotFoundException`:

```
No se puede cargar el archivo o ensamblado 'Newtonsoft.Json, Version=13.0.0.0 ...'
```

La DLL existía en `bin` (compilaba) pero **no estaba en `%APPDATA%\Microsoft\AddIns`, ni declarada en el
`.dna`, ni incluida en el MSI**. El defecto era de empaquetado, no de lógica.

## Regla general

> Toda dependencia que el add-in necesite en tiempo de ejecución debe estar (1) restaurada por NuGet,
> (2) declarada en el `.dna`, (3) instalada por el MSI y (4) presente en la salida de compilación.
> Un script verifica las cuatro y **falla** ante cualquier desajuste.

## Cambios

| Archivo | Cambio |
|---|---|
| `ConversorPDF.XLL/ConversorPDF.XLL-AddIn64.dna` | Se declara `Newtonsoft.Json.dll`. Se retiran `System.Buffers`, `System.Memory` y `System.Numerics.Vectors`: no existen en `bin` ni hay código que las use (referencias fantasma que Excel-DNA intentaba cargar). |
| `ConversorPDF.XLL/ConversorPDF.XLL.csproj` | `PackageReference Newtonsoft.Json 13.0.1` (restauración) + `Reference` con `HintPath=$(NuGetPackageRoot)…` (compilación). Se eliminan las importaciones duplicadas de `obj\*.nuget.g.props/.targets`, que producían MSB4011 e impedían materializar la referencia del paquete. |
| `installer/InstaladorConversorPDF.wxs` | Nuevo `Component cmpNewtonsoft` / `File filNewtonsoft` dentro de `DllComponents` (ya referenciado por `MainFeature`). |
| `build_project.ps1` | Copia explícita de `Newtonsoft.Json.dll` a la carpeta de salida antes de compilar, igual que el resto de dependencias del empaquetador. |
| `scripts/check-addin-packaging.ps1` | **Nuevo contrato ejecutable** (exit 0/2/3): `.dna` ⊆ `bin`, todo `Source` del WiX existe, y todo DLL de `bin` está en el `.dna` o en el instalador (incluye el caso de `x64\SQLite.Interop.dll`). |

## Cómo verificarlo

```powershell
dotnet build ConversorPDF.XLL\ConversorPDF.XLL.csproj -c Release -p:Platform=x64
powershell -ExecutionPolicy Bypass -File scripts/check-addin-packaging.ps1 -Configuration Release
dotnet build installer\InstaladorConversorPDF.wixproj -c Release -p:Platform=x64
```

- El primer comando debe quedar en **0 errores y sin MSB4011**.
- El segundo debe imprimir `EMPACADO_OK` (hoy: 20 DLL en bin, 16 referencias en `.dna`, 21 DLL en WiX).
- El tercero genera el MSI con el componente nuevo.

Para una instalación existente: reinstalar el MSI o copiar `Newtonsoft.Json.dll` y el `.dna` nuevo a
`%APPDATA%\Microsoft\AddIns`. Con Excel abierto el archivo queda bloqueado, así que el despliegue manual se
hace con Excel cerrado.

## Cobertura de pruebas

Sin cambios de lógica: la suite se re-ejecutó completa por lotes tras modificar el `.csproj`
(**1629 ejecutadas, 0 fallos**; baseline 1446) y el contrato de empaquetado se ejecutó contra
`Release|x64` con resultado `EMPACADO_OK`.

## Límites

- No se migró el proyecto a SDK-style: se mantiene el patrón del repositorio (`PackageReference` para
  restaurar + `Reference` explícita para compilar), ahora sin depender de `$(UserProfile)`.
- Los demás `HintPath` del `.csproj` siguen usando `$(UserProfile)`; migrarlos a `$(NuGetPackageRoot)` es
  una mejora mecánica posterior, sin impacto en el fallo corregido.
- El MSI conserva el desajuste preexistente entre versión del producto (`1.6.4.19` en WiX) y de los
  binarios; no se modificó en esta unidad.
