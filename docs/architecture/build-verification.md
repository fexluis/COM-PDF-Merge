# Calibración de Build y Verificación de Pruebas (CLI)

> [!WARNING]
> **Documento inválido para el árbol real (auditoría 2026-10-27).** Los números de esta página
> (51 pruebas, 1 fallo) se midieron en `C:\dev\ConversorPDF`, un `git clone` de HEAD que **no**
> contiene el working tree real (98 `.cs` en XLL y 13 en Tests, frente a 452/138 del árbol).
> El único dato vigente es el comando ganador (`-p:Platform=x64`). La recalibración (R2/R3) debe
> ejecutarse en una copia del árbol real y sobrescribir esta página.

**Fecha**: 2026-10-27  
**Entorno**: Windows x64, .NET SDK 10.0.301, MSBuild 18.6.4, .NET Framework 4.8.10008, Git 2.56.0.3  
**Espacio de trabajo calibrado**: `C:\dev\ConversorPDF` (réplica fuera de OneDrive)

---

## 1. Conclusiones de Calibración

1. **Plataforma del Test Host y Proyectos**:
   - Compilación requerida: `-p:Platform=x64` para ambos proyectos (`ConversorPDF.XLL.csproj` y `ConversorPDF.Tests.csproj`).
   - El test host de `dotnet test` requiere explícitamente `-p:Platform=x64` para ubicar el ensamblado en `bin\x64\Debug\net48\ConversorPDF.Tests.dll`. De omitirse, busca erróneamente en `bin\Debug\...` y falla.
   - `Platform=AnyCPU` no es la configuración predeterminada del proyecto XLL legado (configurado para x64).

2. **Comportamiento de la Suite de Tests**:
   - Total de pruebas descubiertas y ejecutadas: **51**
   - Pruebas superadas (Passed): **50**
   - Fallos preexistentes conocidos (Failed baseline): **1**
     - `ConversorPDF.Tests.UnitTest1.GetUserName_WhenServiceReturnsNull_ShouldReturnFallback`
     - Causa raíz del fallo conocido: Aserción contra `"LUIS.BECHARA"` cuando `Environment.UserName` es `"USUARIO"`.
   - Pruebas de interfaz/UI (`UI/`): Se ejecutan limpiamente en headless sin requerir filtros especiales de exclusión (2/2 Passed).
   - Duración promedio: ~4 segundos.

---

## 2. Comandos Ganadores

### Compilación limpia de la solución
```powershell
dotnet build "ConversorPDF.XLL\ConversorPDF.XLL.csproj" -c Debug -p:Platform=x64 -v:m
dotnet build "ConversorPDF.Tests\ConversorPDF.Tests.csproj" -c Debug -p:Platform=x64 --no-dependencies -v:m
```

### Ejecución de Pruebas Rápidas (Fast / Filtro por SUT)
```powershell
dotnet test "ConversorPDF.Tests\ConversorPDF.Tests.csproj" -c Debug -p:Platform=x64 --no-build --filter "FullyQualifiedName~<Clase>"
```

### Ejecución Completa (Full Suite con reporte TRX)
```powershell
dotnet test "ConversorPDF.Tests\ConversorPDF.Tests.csproj" -c Debug -p:Platform=x64 --no-build --logger "trx;LogFileName=test.trx" --results-directory "<EvidencePath>"
```

---

## 3. Registro de Ejecuciones Reproducibles

### Corrida 1 (2026-10-27T10:28:54)
- **Comando**: `dotnet test "C:\dev\ConversorPDF\ConversorPDF.Tests\ConversorPDF.Tests.csproj" -c Debug -p:Platform=x64 --no-build`
- **Resultado**: Con error: 1, Superado: 50, Omitido: 0, Total: 51. Duración: 4 s.
- **Fallo**: `ConversorPDF.Tests.UnitTest1.GetUserName_WhenServiceReturnsNull_ShouldReturnFallback`.

### Corrida 2 (2026-10-27T10:29:20)
- **Comando**: `dotnet test "C:\dev\ConversorPDF\ConversorPDF.Tests\ConversorPDF.Tests.csproj" -c Debug -p:Platform=x64 --no-build`
- **Resultado**: Con error: 1, Superado: 50, Omitido: 0, Total: 51. Duración: 4 s.
- **Fallo**: `ConversorPDF.Tests.UnitTest1.GetUserName_WhenServiceReturnsNull_ShouldReturnFallback`.
