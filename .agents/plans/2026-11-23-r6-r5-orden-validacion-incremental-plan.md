# Plan: Orden R6→R5, validación incremental y ripple por carril (Ronda 531)

**Generado**: 2026-11-23 · **Estado**: closed

## Execution envelope
- branch: reparacion-automatica-r5
- workspace: C:\Users\USUARIO\OneDrive - Escuela Superior de Administración Publica\Antigravity\ConversorPDFv1.6.3.72
- preflight: powershell -ExecutionPolicy Bypass -File scripts/preflight.ps1 -Workspace . -Plan ".agents/plans/2026-11-23-r6-r5-orden-validacion-incremental-plan.md" -AllowOneDrive
- baseline: tests/baseline/test_baseline.txt
- verify: dotnet test .\ConversorPDF.Tests\ConversorPDF.Tests.csproj
- build_gate: serializado
- max_parallel_writers: 1
- no_push: true
- stop_conditions:
  - Acción destructiva no reversible sin respaldo
  - Violación de seguridad o datos
  - Plan en estado irrecuperable
- numbering:
  - registry: .agents/plans/REGISTRY.md
  - registry_row: R-203
  - ronda_base: 531
  - fase_base: 2363
  - waves: 1
  - tasks: 5

---

## Mapeo de Olas y Rondas

| Wave | Ronda | Tasks |
|---|---|---|
| Wave 531 | Ronda 531 | T1..T5 (Fase 2363–2367) |

---

## Tasks

### T1: Orden R6→R5 y aceptación por peso
- **id**: T1
- **fase**: 2363
- **depends_on**: []
- **location**: ConversorPDF.XLL/Academic/Application/Services/Genetica/
- **files_lock**:
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgGeneradorHorarios.cs`
- **descripción**: Reparación de carriles multipasada con aceptación por peso de reglas (R6=850 > R5=800) y hash de estado; helpers `PesoDuras`, `ConteoDuras`, `MejoraReparacion` testeables.

### T2: Validación incremental “sin duras nuevas”
- **id**: T2
- **fase**: 2364
- **depends_on**: [T1]
- **location**: ConversorPDF.XLL/Academic/Application/Services/Genetica/
- **files_lock**:
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgReparadorR5.cs`
- **descripción**: `ClavesDuras(ind)` con pares R2/R4/R6 y `Validar(..., preexistentes)` que sólo rechaza violaciones nuevas; guardas de vacantes contra `actual`.

### T3: Diagnóstico de duras no-R1/R7
- **id**: T3
- **fase**: 2365
- **depends_on**: [T2]
- **location**: ConversorPDF.XLL/Academic/Application/Services/Genetica/
- **files_lock**:
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgDiagnosticoInfactibilidad.cs`
- **descripción**: Emitir mensajes de R5/R6 (cohorte, tramo, carril) en `Describir` y `ResumenCorto`, truncados.

### T4: Inventario por carril
- **id**: T4
- **fase**: 2366
- **depends_on**: [T3]
- **location**: ConversorPDF.XLL/Academic/Application/Services/Genetica/
- **files_lock**:
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgInventarioCarril.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgInventarioRanuras.cs`
  - `ConversorPDF.XLL/ConversorPDF.XLL.csproj`
- **descripción**: Inventario `Carriles` por `aula|DiasPatron` con fechas válidas/ocupadas/libres y helpers `Iso`/`FechasDeGen`.

### T5: Operador ripple y telemetría
- **id**: T5
- **fase**: 2367
- **depends_on**: [T4]
- **location**: ConversorPDF.XLL/Academic/ + ConversorPDF.Tests/Academic/
- **files_lock**:
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgReparadorR5.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgResultadoReparacionR5.cs`
  - `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgResultadoCorrida.cs`
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgReparadorR5Tests.cs`
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgInventarioRanurasTests.cs`
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgGeneradorHorariosTests.cs`
- **descripción**: `GenerarDesplazamientosRipple` (cadena del carril), telemetría `RippleIntentos/RippleSinCola`, props en `AgResultadoCorrida`, persistencia JSON en ambos bloques y tests del caso #135.

---

## Validación

- Fixture 132 intacto: `Reparar_Fixture132Sintetico...` mueve 150 a 2027-04-02.
- Fixture 135: sin fechas JUE libres suficientes, ripple mueve la cadena y HARD 1→0.
- R6 preexistente: `Validar(..., preexistentes)` acepta candidato sin nuevas duras.
- Sin cola: rollback con `ripple:sin-cola`.
- `AgInventarioRanuras.Construir` expone `Carriles` con `FechasValidas/Ocupadas/Libres`.
- `PesoDuras`/`MejoraReparacion`: acepta R6→R5 (850→800) y rechaza empeoramiento.

## Verificación

```powershell
powershell -ExecutionPolicy Bypass -File scripts/preflight.ps1 -Workspace . -Plan .agents/plans/2026-11-23-r6-r5-orden-validacion-incremental-plan.md -AllowOneDrive
powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "FullyQualifiedName~AgReparadorR5Tests|FullyQualifiedName~AgInventarioRanurasTests|FullyQualifiedName~AgGeneradorHorariosTests"
powershell -ExecutionPolicy Bypass -File scripts/audit-tasks.ps1 -Plan .agents/plans/2026-11-23-r6-r5-orden-validacion-incremental-plan.md -Tasks T1,T2,T3,T4
powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Full
```

- `T1–T4`: **AUDIT_PASS** (0 hallazgos).
- `T5`: 25 bloqueantes + 2 warnings **preexistentes** en `AgFormGeneradorHorarios.cs` (líneas 238–7609), fuera del diff de la ronda (solo +4 líneas JSON en 3284/3473). 0 introducidos; deuda heredada para ronda futura.
- `verify Full`: **2016/2016 PASS**, 0 regresiones. Baseline actualizado a 2016.

## Evidencia

- `_evidence/20261123-194506/summary.md` (Full PASS, 2016/2016)
- `_evidence/20261123-194439/test.trx` (Fast PASS, 21/21)
- `_evidence/20261123-194335-780-audit/audit.md` (T1–T4 PASS)
