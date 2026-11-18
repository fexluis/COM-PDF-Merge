# Handoff: Optimización de Rendimiento AG — Corrida #105

> **Fecha**: 2026-11-17  
> **Plan**: `.agents/plans/2026-11-17-rendimiento-app-corrida105-plan.md`  
> **Registro**: `R-112` (`REGISTRY.md`) · `Wave 1..6` → `Ronda 433–438` · `T1..T10` → `Fase 1942–1951`

---

## 1. Entorno de Ejecución
- **Workspace**: `C:\Users\USUARIO\OneDrive - Escuela Superior de Administración Publica\Antigravity\ConversorPDFv1.6.3.72`
- **Rama**: `rendimiento-app` (creada a partir de `315b9ec` y publicada en origin)
- **Preflight Command**:
  ```powershell
  powershell -ExecutionPolicy Bypass -File scripts/preflight.ps1 -Workspace "C:\Users\USUARIO\OneDrive - Escuela Superior de Administración Publica\Antigravity\ConversorPDFv1.6.3.72" -Plan ".agents/plans/2026-11-17-rendimiento-app-corrida105-plan.md" -AllowOneDrive
  ```
- **Gate de Verificación**:
  ```powershell
  dotnet build .\ConversorPDF.Tests\ConversorPDF.Tests.csproj -c Debug
  dotnet test .\ConversorPDF.Tests\ConversorPDF.Tests.csproj -c Debug --no-build --logger "trx;LogFileName=gate-alternativo.trx"
  ```
  *(Nota ACL: `scripts/verify.ps1` no disponible por denegación ACL en el host; se usa el gate alternativo con TRX autorizado).*
- **Baseline de Pruebas**: 1446 declarado en `tests/baseline/test_baseline.txt`; R-109 reporta Full 1648. El `build-verifier` medirá el workspace antes de T1 y actualizará el baseline sin reducirlo.

---

## 2. Comando de Ejecución
Para iniciar la ejecución multiagente en paralelo:
```text
/swarm-parallel-task .agents/plans/2026-11-17-rendimiento-app-corrida105-plan.md
```

---

## 3. Compuerta de Entrada (Gate Obligatorio)
Antes de ejecutar la Wave 1:
1. Auditar `.agents/plans/2026-11-17-rendimiento-app-corrida105-plan.md` contra las 9 reglas duras de `plan-reviewer` (verificado: **APROBADO**).
2. Si un revisor emitiera un dictamen `BLOCKED` o reporte de hallazgos críticos, detener inmediatamente la ejecución y subsanar el plan.

---

## 4. Orden de Ejecución por Olas

1. **Fase A (Preparación)**:
   - Rama `rendimiento-app` creada y vinculada con origin.
   - Plan, handoff y registro `R-112` formalizados.
   - Ejecución de preflight con bandera `-AllowOneDrive` (`PREFLIGHT_OK`).
2. **Wave 1 (Ronda 433)**: `T1` (Semilla determinista y flujos RNG).
3. **Wave 2 (Ronda 434)**: `T2` (Telemetría y `DuracionMs`) y `T3` (Arnés replay baseline sembrado). Leases disjuntos.
4. **Wave 3 (Ronda 435)**: `T4` (Subetapas reparación/sellado) y `T5` (Cachés de higiene). Leases disjuntos.
5. **Wave 4 (Ronda 436)**: `T6` (Quick wins reparación) y `T7` (Proyección ligera búsqueda). Leases disjuntos.
6. **Wave 5 (Ronda 437)**: `T8` (Índice individuo y reparación sucia) y `T9` (Huella y caché aptitud). Leases disjuntos.
7. **Wave 6 (Ronda 438)**: `T10` (Benchmark, equivalencia y certificación).
8. **Fase C (Cierre y Sellado)**: `swarm-verify` sella `R-112`, actualiza contadores globales y registra evidencias.

---

## 5. Criterios de Aceptación y Presupuesto

### 5.1. Equivalencia Funcional y Calidad
- **Determinismo**: Misma semilla en baseline y versión optimizada produce idéntica trayectoria de aptitud en las fases de higiene/refactor; diferencias explicadas si una optimización altera orden de exploración.
- **Violaciones**: ≤ 30 violaciones blandas históricas, `HARD=0` garantizado en campeón.
- **Volúmenes canónicos**: 410 genes, 2.228 sesiones, 8.912 horas, 199 vinculaciones, 336 vacantes.
- **Materialización**: Hash canónico del horario materializado idéntico entre baseline y versión optimizada.
- **Desglose de vacantes**: Clasificación por `AgReglaViolacion` y `MotivoVacante`, verificando vacantes en `AsignaturasVacantesConsentidas`.

### 5.2. Presupuesto de Tiempo
- **Actual**: 58 transiciones de estado estacionario × 3,90 s = 226 s + ~14 s iniciales = 240,158 s.
- **Meta intermedia**: ≤ 1,00 s por transición de estado estacionario (total ≤ 72 s con 14 s iniciales, o ~65 s con inicio optimizado).
- **Meta objetivo**: ~0,65 s por transición de estado estacionario + ~7 s de inicialización = **~44,7 s (≤ 45 s)**.
- **Resolución**: Tiempos de transición p50/p95 y `DuracionMs` medidos con `Stopwatch`, no con timestamps de log de 1 s.

---

## 6. Prohibiciones y Restricciones Estrictas
- ❌ **No push de implementación**: Mantener `no_push: true`. Únicamente se autorizó el push inicial de la rama `rendimiento-app`.
- ❌ **No alterar orden de Random en T1–T4**: Conservar exactamente la secuencia de números aleatorios para garantizar la paridad del baseline.
- ❌ **No tocar L0 ni modificar `.csproj`**: Prohibido agregar archivos nuevos a `ConversorPDF.XLL.csproj`. Toda nueva clase auxiliar (`AgPerfilGeneracion`, `AgProyeccionIndividuo`, etc.) se declara como clase interna anidada en archivos existentes.
- ❌ **No migraciones ni columnas nuevas**: La telemetría es en memoria y en líneas `[AG-PERF]`; no alterar esquemas SQLite.
- ❌ **No reducir `estancamientoMax`**: Mantener el límite de estancamiento para preservar las mejoras tardías observadas en Gen 40.
- ❌ **No asumir `HARD=0` poblacional**: El indicador `HARD=0` corresponde al mejor individuo; la población presenta violaciones que influyen en el tiempo de serialización y sellado.

---

## 7. Reporte Obligatorio por Tarea
Cada tarea debe documentar en su log y evidencia:
- Archivos editados y creados.
- Salida del comando de verificación y archivo `.trx`.
- Métricas comparativas: `TiempoTotal`, `TiempoCruceNeto`, `TiempoMutacionNeto`, `TiempoSelladoTotal`, `TiempoEvaluacion`, `EvaluacionesAptitud`, `SesionesProyectadas`, `LlamadasDisponibleGrupo`, `RegistrarJson`, `Duplicados`.
- Dictamen de avance `GO / NO-GO` de la ola.
