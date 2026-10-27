# Calibración de Build y Verificación de Pruebas (CLI)

**Calibrado**: 2026-10-27 (R2/R3) · **Workspace**: `C:\dev\ConversorPDF` (copia del commit `6ae4995`: 452 `.cs` en XLL, 138 en Tests)
**Entorno**: Windows x64 · .NET SDK 10.0.301 · MSBuild 18.6.4 · .NET Framework 4.8.10008 · Git 2.56.0.3

> [!IMPORTANT]
> La calibración previa (51 pruebas / 1 fallo) se hizo sobre un `git clone` de HEAD con el árbol incompleto
> (98/452 y 13/138 `.cs`) y quedó **invalidada** por la auditoría del 2026-10-27. Los números de esta página
> son los primeros medidos sobre el árbol real.

---

## 1. Comandos ganadores

```powershell
# Compilación (obligatorio -p:Platform=x64)
dotnet build "ConversorPDF.XLL\ConversorPDF.XLL.csproj" -c Debug -p:Platform=x64 -v:m
dotnet build "ConversorPDF.Tests\ConversorPDF.Tests.csproj" -c Debug -p:Platform=x64 --no-dependencies -v:m

# Ciclo TDD (una clase)
powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "FullyQualifiedName~<Clase>"

# Gate completo
powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Full
```

Salidas: `_evidence/<yyyyMMdd-HHmmss>/{build.log,test.log,test.trx,summary.md}`.

---

## 2. Resultados medidos (árbol real)

| Métrica | Valor |
|---|---|
| Compilación XLL | OK · 0 errores · ~6 s |
| Compilación tests | OK · 0 errores · ~4 s |
| Suite completa | **939 pruebas · 939 passed · 0 failed · 0 skipped** |
| Duración suite | ~58 s |
| Fallos preexistentes | **ninguno** |
| Baseline | `tests/baseline/test_baseline.txt` → `TOTAL: 939` |
| HEAD de referencia | `6ae4995` |

---

## 3. Notas de calibración

1. **`-p:Platform=x64` es obligatorio**: sin él, el test host busca en `bin\Debug\...` y no encuentra el ensamblado.
2. **El gate es serial**: dos `verify.ps1` concurrentes chocan en `bin/x64/Debug/*.xll` (`ExcelDnaBuild DNA-577853131`).
3. **Tamaño de suite verificable**: `Full` falla con exit `3` (`SUITE_SHRUNK`) si se ejecutan menos pruebas que `TOTAL`.
4. **Warnings tolerados** (preexistentes, no bloquean): `MSB4011` ×2 por doble import de `obj\*.nuget.g.props/targets`, y `CS0162` en `MigrationV76_SanearDimensionesAgAula.cs:90`.
5. **De 51 a 939 pruebas**: la diferencia es exactamente el árbol real (Academic incluido). Cualquier número menor indica un workspace incompleto: ejecutar `scripts/preflight.ps1`.

---

## 4. Reproducir desde cero

```powershell
scripts\preflight.ps1 -Workspace .        # debe devolver 0
powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Build
powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Full
```
