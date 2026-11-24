# Plan: Reparación R6 dedicada y orden Sellar→plan (Ronda 532)

**Generado**: 2026-11-23 · **Estado**: closed

## Execution envelope
- branch: reparacion-automatica-r5
- workspace: C:\Users\USUARIO\OneDrive - Escuela Superior de Administración Publica\Antigravity\ConversorPDFv1.6.3.72
- preflight: powershell -ExecutionPolicy Bypass -File scripts/preflight.ps1 -Workspace . -Plan ".agents/plans/2026-11-23-r6-reparacion-dedicada-plan.md" -AllowOneDrive
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
  - registry_row: R-204
  - ronda_base: 532
  - fase_base: 2368
  - waves: 1
  - tasks: 3

---

## Mapeo de Olas y Rondas

| Wave | Ronda | Tasks |
|---|---|---|
| Wave 532 | Ronda 532 | T1..T3 (Fase 2368–2370) |

---

## Tasks

### T1: El plan de carril sobrevive a Sellar
- **id**: T1
- **fase**: 2368
- **depends_on**: []
- **location**: ConversorPDF.XLL/Academic/Application/Services/Genetica/
- **files_lock**:
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgGeneradorHorarios.cs`
- **descripción**: Retirar `AgCoherenciaGrupo.Sellar` posterior a `RealinearOleadas` (el sellado previo ya ocurrió) para que el plan del carril no se borre; mantener la evaluación y el criterio por peso.

### T2: Reparador R6 dedicado
- **id**: T2
- **fase**: 2369
- **depends_on**: [T1]
- **location**: ConversorPDF.XLL/Academic/Application/Services/Genetica/
- **files_lock**:
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgReparadorR6.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgResultadoReparacionR6.cs`
  - `ConversorPDF.XLL/ConversorPDF.XLL.csproj`
- **descripción**: Detectar pares R6 por cohorte/carril, elegir la víctima corta no fija y reproyectar su bloque tras el fin del solapado usando `AgInventarioCarril`; aceptar sólo si baja el HARD; sin `Sellar`.

### T3: Hook, telemetría y pruebas
- **id**: T3
- **fase**: 2370
- **depends_on**: [T2]
- **location**: ConversorPDF.XLL/Academic/ + ConversorPDF.Tests/Academic/
- **files_lock**:
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgGeneradorHorarios.cs`
  - `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgResultadoCorrida.cs`
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgReparadorR6Tests.cs`
- **descripción**: Invocar `AgReparadorR6` antes del hook R5, exponer `ReparacionR6Intentos/HardPre/HardPost` y cubrir con tests el fixture de carril largo + dos cortos y el caso de asignaturas fijas.

---

## Validación

- Fixture R6: 152 largo (feb–may) + 151/153 cortos dentro del mismo carril → `AgReparadorR6` mueve los cortos después de 152 y HARD 2→0.
- Fijos: si la víctima corta es `EsFijo`, no se mueve.
- Regresión: fixture 132 y 135 de `AgReparadorR5Tests` siguen en verde.
- `Sellar` no se ejecuta después de `RealinearOleadas`.

## Verificación

```powershell
powershell -ExecutionPolicy Bypass -File scripts/preflight.ps1 -Workspace . -Plan .agents/plans/2026-11-23-r6-reparacion-dedicada-plan.md -AllowOneDrive
powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "FullyQualifiedName~AgReparadorR6Tests|FullyQualifiedName~AgReparadorR5Tests|FullyQualifiedName~AgGeneradorHorariosTests"
powershell -ExecutionPolicy Bypass -File scripts/audit-tasks.ps1 -Plan .agents/plans/2026-11-23-r6-reparacion-dedicada-plan.md -Tasks T1,T2
powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Full
```

- `T1–T3`: **AUDIT_PASS** (0 hallazgos).
- `verify Fast` (R6+R5+Generador): 19/19 PASS.
- `verify Full`: **2018/2018 PASS**, 0 regresiones. Baseline actualizado a 2018.

## Evidencia

- `_evidence/20261123-202210/summary.md` (Full PASS)
- `_evidence/20261123-202140/test.trx` (Fast PASS)
- `_evidence/20261123-202159-659-audit/audit.md` (T1–T2 PASS)
- `_evidence/20261123-202442-160-audit/audit.md` (T3 PASS)
