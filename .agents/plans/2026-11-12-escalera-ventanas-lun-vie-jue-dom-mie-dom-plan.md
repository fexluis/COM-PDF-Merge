---
numbering:
  registry: .agents/plans/REGISTRY.md
  registry_row: R-058
  ronda_base: 345
  fase_base: 1707
  waves: 6
  tasks: 8
---

# Plan: escalado automático de ventanas LUN-VIE / JUE-DOM / MIE-DOM (consolidado)
**Generado**: 2026-11-12 · **Autor**: Antigravity/Swarm-Planner · **Estado**: closed

## Execution envelope
- branch: refactor/academic-modularization
- workspace: c:\Users\USUARIO\OneDrive - Escuela Superior de Administración Publica\Antigravity\ConversorPDFv1.6.3.72
- preflight: powershell -ExecutionPolicy Bypass -File scripts/preflight.ps1 -Workspace "c:\Users\USUARIO\OneDrive - Escuela Superior de Administración Publica\Antigravity\ConversorPDFv1.6.3.72" -Plan .agents/plans/2026-11-12-escalera-ventanas-lun-vie-jue-dom-mie-dom-plan.md -AllowOneDrive
- workspace_guard:
  - min_xll_cs: 300
  - min_tests_cs: 100
  - required_slice: Academic
- baseline: tests/baseline/test_baseline.txt (TOTAL: 1172)
- verify: powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Full
- build_gate: serializado (un único build/tests a la vez; Excel-DNA bloquea bin/)
- max_parallel_writers: 2
- agents: plan-reviewer, tdd-red-writer, implementer, integrator, task-reviewer, build-verifier
- hot_paths_L0: (ninguno tocado)
  - ConversorPDF.XLL/ConversorPDF.XLL.csproj
  - ConversorPDF.XLL/Core/Helpers/App.cs
  - ConversorPDF.XLL/**/Migrations/MigrationV*.cs
  - ConversorPDF.XLL/**/Persistence/SchemaMigrator.cs
  - ConversorPDF.XLL/Ribbon/RibbonController.cs
  - docs/toc.yml
- stop_conditions:
  - Acción destructiva no reversible sin respaldo
  - git clean / git stash / cambio de rama con archivos sin seguimiento
  - Intento de push, merge o publicación a remoto
  - Plan irrecuperable tras más de 3 rondas de fix fallidas
- ledger_owner: swarm-parallel-task
- no_push: true
- numbering:
  - registry: .agents/plans/REGISTRY.md
  - registry_row: R-058
  - ronda_base: 345          # Wave 1 = Ronda 345
  - fase_base: 1707          # T1 = Fase 1707
  - waves: 6                 # Ronda 345–350
  - tasks: 8                 # Fase 1707–1714

---

## 1. Grafo de dependencias

```text
T1 (stub Rotular) ─┐
T2 (RED parser) ───┴─ T4 (impl parser)
                        │
T3 (RED dominio+integración+persistencia) ── T5 (impl escalera+Rotular) ── T6 (escalador: solo Siguiente)
                                                                              │
                                                                        T7 (UI: recarga + leyenda + modal)
                                                                              │
                                                                        T8 (docs/cierre)
```

Dependencias declaradas: `T2: []` · `T3: [T1]` · `T4: [T2]` · `T5: [T3, T4]` · `T6: [T5]` · `T7: [T4, T6]` · `T8: [T5, T6, T7]`.
T1 y T2 comparten ola (lease disjunto: producción vs tests). T3 y T4 comparten ola (tests vs parser). Cada ola ≤ 2 escritores.

---

## 2. Olas (Parallel execution groups)

| Wave | Ronda | Tasks | Puede iniciar cuando | Gate de la ola |
|---|---|---|---|---|
| 1 | 345 | T1, T2 | Inmediato (leases disjuntos) | `-Mode Build` (stub compila) + RED parser registrado |
| 2 | 346 | T3, T4 | T1 y T2 completados | Build + `AgDiasVentanaTests` verde (parser) + RED del resto registrado |
| 3 | 347 | T5 | T3 y T4 completados | Dominio + integración verdes (escalera, planner, validador, mapper, escalador nominal) |
| 4 | 348 | T6 | T5 completado | Escalador + persistencia verdes (`ag_aula`, `ag_arbol_nodo`, recarga, segunda corrida) |
| 5 | 349 | T7 | T4 y T6 completados | Build + flujo escalar→recargar→cambiar→guardar→recargar |
| 6 | 350 | T8 | T5, T6 y T7 completados | `-Mode Full` = 0; TOTAL ≥ baseline (1172) |

Nota de gate: entre la ola 2 y la 5 la suite completa queda roja a propósito (los RED de T3 siguen abiertos). `-Mode Full` **solo** se ejecuta en la ola 6; un `Full` intermedio con menos pruebas que el baseline devolvería `SUITE_SHRUNK` (3) y no es un fallo del plan.

---

## 3. Tareas (contrato de 11 campos)

### T1: Stub de `Rotular` para habilitar RED compilable
- **id**: T1
- **fase**: 1707
- **depends_on**: []
- **location**: `ConversorPDF.XLL/Academic/Domain/Models/Genetica/`
- **files_lock**: `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgEscaleraVentanas.cs`
- **description**: añadir **solo** el miembro `internal static string Rotular(string diasVentana, double unidad)` con `throw new NotImplementedException();`. Sin cambios de escalera, sin `Normalizar`, sin `csproj` (la clase ya existe y está registrada). Objetivo: que el proyecto de tests compile cuando T3 invoque `Rotular` (Fase 1702).
- **validation**: `-Mode Build` sin errores; el método existe con la firma exacta; `git diff` del archivo limitado a ese miembro.
- **verification**: `powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Build`
- **evidence**: `_evidence/20261112-153005/summary.md`
- **status**: done
- **log**: Stub de Rotular añadido en AgEscaleraVentanas.cs; compila con 0 errores CS en XLL y Tests. Auditado por task-reviewer (OK).
- **files edited/created**: `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgEscaleraVentanas.cs`

### T2: RED del parser de ventanas
- **id**: T2
- **fase**: 1708
- **depends_on**: []
- **location**: `ConversorPDF.Tests/Academic/Domain/`
- **files_lock**: `ConversorPDF.Tests/Academic/Domain/AgDiasVentanaTests.cs`
- **description**: casos nuevos, sin tocar producción: `AIso/Dias("LUN-VIE")="1,2,3,4,5"`, `JUE-DOM="4,5,6,7"`, `MIE-DOM="3,4,5,6,7"`; rangos genéricos `LUN-MAR=[1,2]`, `SAB-DOM=[6,7]`; sin envolvente `VIE-LUN=[]`; tokens de franja rechazados `Dias("6M")=[]`, `Dias("6T")=[]` (R1); no regresión de las 5 etiquetas vigentes ni de `"UNKNOWN"`.
- **validation**: compila con el stub de T1 y **falla por aserción** (RED real, no por compilación). Validación limitada al parser: no incluye escalera ni integración.
- **verification**: `powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "FullyQualifiedName~AgDiasVentana"`
- **evidence**: `_evidence/20261112-152930/summary.md`
- **status**: red
- **log**: RED_CERTIFIED por tdd-red-writer. 27 tests (16 pasados, 11 fallando por aserción esperada). Auditado por task-reviewer (OK).
- **files edited/created**: `ConversorPDF.Tests/Academic/Domain/AgDiasVentanaTests.cs`

### T3: RED de dominio, integración y persistencia
- **id**: T3
- **fase**: 1709
- **depends_on**: [T1]
- **location**: `ConversorPDF.Tests/Academic/`
- **files_lock**:
  - `ConversorPDF.Tests/Academic/Domain/AgEscaleraVentanasTests.cs`
  - `ConversorPDF.Tests/Academic/Infrastructure/Persistence/AgEscaladorVentanaAulaTests.cs`
  - `ConversorPDF.Tests/Academic/Infrastructure/Persistence/AgEscaladorVentanaAulaPersistenciaTests.cs`
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgPlanificadorOleadasTests.cs`
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgPlanificadorOleadasCaracterizacionTests.cs`
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgValidadorVentanaTests.cs`
  - `ConversorPDF.Tests/Academic/Application/Mappers/AgNodoArbolKeyMapperTests.cs`
- **description**: tres bloques falsables.
  **(a) Escalera y rótulo:** `Siguiente("LUN-JUE")=="LUN-VIE"`, `Siguiente("LUN-VIE")==""`, `Siguiente("MAR-VIE")=="LUN-VIE"`, `Siguiente("VIE-DOM")=="JUE-DOM"`, `Siguiente("JUE-DOM")=="MIE-DOM"`, `Siguiente("MIE-DOM")==""`; `Escalera("VIE-DOM")=[JUE-DOM, MIE-DOM]` y nunca 3; cadena `VIE-SAB→JUE-SAB→4,5,6,7→""` intacta; `PuedeEscalar("4,5,6,7")==false`; idempotencia; `Rotular("MIE-DOM",4)=="3,4,5,6M,6T,7"`, `Rotular("MIE-DOM",8)=="3,4,5,6,7"`, `Rotular("LUN-VIE",4)=="1,2,3,4,5"`.
  **(b) Integración nominal (Fase 1704):** fixture con **W semanas completas, sin festivos y bloques activos para todas las ranuras**; `CapacidadAntes=4W` y `CapacidadDespues=5W`. `Construir` en `LUN-VIE` con tres asignaturas de demanda `2W`, `2W`, `W` y patrones sembrados `1,2`; `3,4`; `5` → `5W` sesiones planificadas, `Infactibles=0`. Casos adicionales: festivo en el día añadido (W=1 → sin ganancia), bloque del día añadido ausente (sin ganancia), `5W+1` → `I5_SIN_CAPACIDAD`; `VIE-DOM=4W`, `JUE-DOM=5W`, `MIE-DOM=6W`; `VentanaSugerida=="JUE-DOM"` para `VIE-DOM`.
  **(c) Persistencia (Fase 1705):** SQLite temporal con el patrón de `MigrationV76Tests` (`SQLiteConnection.CreateFile` + `Func<IDbConnection>`). `ag_aula` y `ag_arbol_nodo` con **dos árboles**; `AplicarAsync(decisiones, idArbol, ct)`; verificar por consulta y por contexto recargado: `LUN-JUE`/`MAR-VIE` con demanda `5W` → `ag_aula.dias_ventana="LUN-VIE"`, nodo del árbol activo `=LUN-VIE`, **otros árboles intactos**, `Evaluar(contextoRecargado)` vacío. `VIE-DOM` con demanda `6W` → 1.ª aplicación `JUE-DOM`, recarga, 2.ª aplicación `MIE-DOM`, recarga, 3.ª evaluación vacía. **No** se exige segundo escalado con `5W`.
  **(d) Caso 1703:** primer paso sin ganancia y segundo con ganancia → **ninguna propuesta** (se fija `Siguiente` como único candidato).
- **validation**: todo falla por aserción; compila con el stub (T1). No se modifica producción.
- **verification**: `powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "FullyQualifiedName~AgEscaleraVentanas|FullyQualifiedName~AgEscaladorVentanaAula|FullyQualifiedName~AgPlanificadorOleadas|FullyQualifiedName~AgValidadorVentana|FullyQualifiedName~AgNodoArbolKeyMapper"`
- **evidence**: _evidence/20261112-153415/summary.md
- **status**: red
- **log**: RED_CERTIFIED por tdd-red-writer. Suite compila con 0 errores CS. 14 fallos esperados por aserción/stub en tests nuevos de dominio, persistencia e integración. Auditado por task-reviewer (OK).
- **files edited/created**:
  - ConversorPDF.Tests/Academic/Domain/AgEscaleraVentanasTests.cs
  - ConversorPDF.Tests/Academic/Infrastructure/Persistence/AgEscaladorVentanaAulaTests.cs
  - ConversorPDF.Tests/Academic/Infrastructure/Persistence/AgEscaladorVentanaAulaPersistenciaTests.cs
  - ConversorPDF.Tests/Academic/Application/Services/Genetica/AgPlanificadorOleadasTests.cs
  - ConversorPDF.Tests/Academic/Application/Services/Genetica/AgPlanificadorOleadasCaracterizacionTests.cs
  - ConversorPDF.Tests/Academic/Application/Services/Genetica/AgValidadorVentanaTests.cs
  - ConversorPDF.Tests/Academic/Application/Mappers/AgNodoArbolKeyMapperTests.cs

### T4: Parser único de ventanas
- **id**: T4
- **fase**: 1710
- **depends_on**: [T2]
- **location**: `ConversorPDF.XLL/Academic/Domain/Models/Genetica/`
- **files_lock**: `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgDiasVentana.cs`
- **description**: añadir a `Opciones` `LUN-VIE→"1,2,3,4,5"`, `JUE-DOM→"4,5,6,7"`, `MIE-DOM→"3,4,5,6,7"`. Resolver rangos `AAA-BBB` solo si ambos extremos son válidos y `inicio<=fin` (sin envolvente). Mantener ISO, abreviaturas, orden ascendente y `Distinct`. Rechazar `6M/6T` devolviendo `[]` (R1). Sin métodos públicos nuevos (clase pública con superficie potencial de UDF).
- **validation**: **solo** pruebas del parser verdes; `Dias("VIE-LUN")=[]`, `Dias("6M")=[]`; las 5 etiquetas vigentes sin regresión.
- **verification**: `powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "FullyQualifiedName~AgDiasVentana"`
- **evidence**: _evidence/20261112-154403/summary.md
- **status**: done
- **log**: GREEN. Opciones incluye LUN-VIE, JUE-DOM, MIE-DOM. Rangos AAA-BBB sin envolvente y rechazo R1 de franjas 6M/6T. 27/27 tests pasando. Auditado por task-reviewer (OK).
- **files edited/created**: ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgDiasVentana.cs

### T5: Escalera ampliada y `Rotular` real
- **id**: T5
- **fase**: 1711
- **depends_on**: [T3, T4]
- **location**: `ConversorPDF.XLL/Academic/Domain/Models/Genetica/`
- **files_lock**: `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgEscaleraVentanas.cs`
- **description**: reemplazar el stub de `Rotular` por:
```csharp
internal static string Rotular(string diasVentana, double unidad)
{
    return string.Join(",", AgRanuras
        .PorVentana(diasVentana, unidad: unidad)
        .Select(r => r.Dia == 6 && r.Horas < 8
            ? (string.Equals(r.Franja, AgFranjaJornada.MANANA, StringComparison.OrdinalIgnoreCase) ? "6M" : "6T")
            : r.Dia.ToString()));
}
```
  Reconstruir `Escalones` (etiqueta → siguientes, `[0]` = paso inmediato):
  `LUN-JUE→[LUN-VIE]`, `MAR-VIE→[LUN-VIE]`, `LUN-VIE→[]`, `VIE-SAB→[JUE-SAB,"4,5,6,7"]` (**sin cambio**), `JUE-SAB→["4,5,6,7"]`, `"4,5,6,7"→[]`, `VIE-DOM→[JUE-DOM,MIE-DOM]`, `JUE-DOM→[MIE-DOM]`, `MIE-DOM→[]`, `"3,4,5,6,7"→[]`; alias ISO de legado: `"1,2,3,4"→[LUN-VIE]`, `"2,3,4,5"→[LUN-VIE]`, `"1,2,3,4,5"→[]`, `"5,6"→["4,5,6"]`, `"4,5,6"→["4,5,6,7"]`, `"5,6,7"→[JUE-DOM]`. `Normalizar`: coincidencia exacta de **Key** primero; si no, `AIso` (deja de mapear alias ISO→etiqueta, R2). `MaxPasos=2` intacto; firmas públicas sin cambios.
- **validation**: dominio e integración verdes (escalera, `Rotular`, planner, validador, mapper, escalador nominal); `Siguiente("4,5,6,7")==""`; `Siguiente("MIE-DOM")==""`.
- **verification**: `powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "FullyQualifiedName~AgEscaleraVentanas|FullyQualifiedName~AgPlanificadorOleadas|FullyQualifiedName~AgValidadorVentana|FullyQualifiedName~AgNodoArbolKeyMapper"`
- **evidence**: _evidence/20261112-154543/summary.md
- **status**: done
- **log**: GREEN. Escalera ampliada y Rotular real implementados en AgEscaleraVentanas.cs. Dominio e integración verdes (AgEscaleraVentanasTests 21/21, KeyMapper/ValidadorVentana 35/35, AgPlanificadorOleadas 33/33, Caracterizacion 8/8). Auditado por task-reviewer (OK).
- **files edited/created**: ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgEscaleraVentanas.cs

### T6: Escalador — evaluar exclusivamente `Siguiente`
- **id**: T6
- **fase**: 1712
- **depends_on**: [T5]
- **location**: `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/`
- **files_lock**: `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgEscaladorVentanaAula.cs`
- **description**: en `Evaluar` (`:65-82`) eliminar el `foreach` y el `break`; evaluar solo el paso inmediato:
```csharp
var paso = AgEscaleraVentanas.Siguiente(aula.DiasVentana);
if (string.IsNullOrWhiteSpace(paso)) continue;

int capacidad = Capacidad(paso, aula.Jornada, unidad, contexto, ini, fin);
if (capacidad <= capacidadAntes) continue;

// Conservar la creación de Decision existente; Motivo con Desde, Hacia y capacidades.
```
  `AplicarAsync` no se modifica (ya escribe `ag_aula` y `ag_arbol_nodo` con alcance `idArbol`).
- **validation**: escalador + persistencia verdes: `LUN-JUE→LUN-VIE`, `MAR-VIE→LUN-VIE`, `VIE-DOM→JUE-DOM`, `JUE-DOM→MIE-DOM`; primer paso sin ganancia → 0 propuestas; `5W` no dispara segunda escalada; `MIE-DOM`/`LUN-VIE` sin propuesta.
- **verification**: `powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "FullyQualifiedName~AgEscaladorVentanaAula"`
- **evidence**: _evidence/20261112-155126/summary.md
- **status**: done
- **log**: GREEN. Evaluar evalúa exclusivamente AgEscaleraVentanas.Siguiente. Caso 1703 verificado. Pruebas unitarias (12/12) y de persistencia (3/3) en PASS. Auditado por task-reviewer (OK).
- **files edited/created**:
  - ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgEscaladorVentanaAula.cs
  - ConversorPDF.Tests/Academic/Infrastructure/Persistence/AgEscaladorVentanaAulaTests.cs

### T7: UI — recarga de claves y leyenda no clicable
- **id**: T7
- **fase**: 1713
- **depends_on**: [T4, T6]
- **location**: `ConversorPDF.XLL/Academic/UI/Forms/`
- **files_lock**: `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **description**: tres cambios quirúrgicos, con anclas verificadas.
  **(a) Modal genérico:** sustituir el texto fijo de `:1398-1401` por una cadena construida desde las `Decision` (`Desde → Hacia`) y `AgEscaleraVentanas.Rotular(Hacia, unidad)`; sin casos hardcodeados.
  **(b) Recarga antes de habilitar edición (Fase 1706):** tras `AplicarAsync` (`:1411`) y la actualización en memoria (`:1414-1419`), y **antes** de la re-validación (`:1421-1430`): obtener los DTO persistidos con `_arbolService.CargarPerfilAsync(idArbolActivo, ct)` (mismo patrón de `:4841`), localizar los nodos de nivel 6 por `AulaRef` en el árbol activo y reaplicar `AgNodoArbolKeyMapper.AplicarDimensiones(key, dto)`; luego `treeContextoAcademico.Invalidate()`. No reejecutar `CargarPerfilArbolSeleccionadoAsync` completo (rebaraja combos y contexto); helper estrecho con guarda de reentrada `_restaurandoPerfil`. Alcance: solo `idArbolActivo`; otros árboles intactos.
  **(c) Leyenda (Fase 1706, ubicación concreta):** **no** dibujar nada dentro de `[DxLunJue, DxSep2)`: `NodeMouseClick` (`:316-353`) ignora `e.Y` y `ProcesarClickDimensionNivel6` (`:4167-4186`) traduce cualquier `dx` de ese rango a una ventana. Solución: `treeContextoAcademico.ShowNodeToolTips = true` y, para nivel 6 con `dias_ventana` fuera de las 5 etiquetas, `node.ToolTipText = $"Ventana actual: {dv} ({AgDiasVentana.AIso(dv)})"` (detalle seleccionado, no clicable, sin layout). Defensa en profundidad: añadir guarda `dy` en el handler (`dy < 0 || dy > e.Node.Bounds.Height → return`) antes de delegar en `ProcesarClickDimensionNivel6`. Si se exigiera leyenda visible, ubicarla fuera de todo rango `x` (cabecera `lblColDias` o franja de estado) y solo con la guarda `dy`.
- **validation**: `-Mode Build` sin warnings nuevos; flujo en test de integración (repositorio + mapper, sin WinForms): escalar → recargar DTO → cambiar `Jornada` → guardar → recargar → `DiasVentana` conserva `LUN-VIE`/`MIE-DOM` y los demás árboles no cambian. Smoke manual diferido a la ejecución.
- **verification**: powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "FullyQualifiedName~AgNodoArbolKeyMapper|FullyQualifiedName~AgEscaladorVentanaAula"
- **evidence**: _evidence/20261112-155750/summary.md
- **status**: done
- **log**: GREEN. Modal genérico con Rotular dinámico según unidad del aula. Recarga de dimensiones del perfil activo implementada sin rebarajar combos. ToolTip seguro y guarda vertical dy en click de nivel 6 integrados. Compilación y tests en PASS (23/23). Auditado por task-reviewer.
- **files edited/created**:
  - ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs

### T8: Documentación y cierre
- **id**: T8
- **fase**: 1714
- **depends_on**: [T5, T6, T7]
- **location**: `docs/articles/changes/`
- **files_lock**: `docs/articles/changes/2026-11-12-escalera-ventanas-lun-vie-jue-dom-mie-dom.md`
- **description**: artículo con: tabla de capacidad **nominal** (4 h/8 h) y su diferencia con la efectiva (`4W`, `5W`, `6W`); rutas de escalera; R1 (sábado como rótulo), R2 (ISO `4,5,6,7` terminal), R3 (sin migración/L0), R4 (un escalón por corrida); códigos `I5_INFECTIBLE`/`I5_CAPACIDAD_CARRIL_PATRON`/`R8`; límites y coordinación con el plan de ranuras 4→8 h. En Fase C: fila R-057/R-058, contadores 350/1714, `verify`, `HEAD`, `implementation_plan.md`/`walkthrough.md`.
- **validation**: documentación enlazada; `-Mode Full` = 0 y `TOTAL` ejecutado ≥ baseline (1172).
- **verification**: `powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Full`
- **evidence**: _evidence/20261112-155902/summary.md
- **status**: done
- **log**: GREEN. Artículo técnico redactado en docs/articles/changes/. Gate verify -Mode Full superado con éxito: 1223 tests PASS (0 fallos, 0 errores), superando el baseline oficial de 1172 tests. Registro REGISTRY.md sellado en fila R-058 como closed.
- **files edited/created**:
  - docs/articles/changes/2026-11-12-escalera-ventanas-lun-vie-jue-dom-mie-dom.md
  - tests/baseline/test_baseline.txt
  - .agents/plans/REGISTRY.md

---

## 4. Contrato de fixtures (Fase 1704)

```text
Precondiciones de toda medición de capacidad:
  - W semanas completas ancladas a fechaInicio (sin semanas parciales).
  - Ningún festivo, salvo el caso que lo declare explícitamente.
  - Catálogo con bloque activo de la unidad (4 h u 8 h) para TODAS las ranuras.
  - Unidad = 4 h (pregrado). Para 8 h, el sábado aporta 1 ranura y se declara.

Capacidad nominal por semana (4 h):
  LUN-JUE 4 · LUN-VIE 5 · MAR-VIE 4 · VIE-DOM 4 · JUE-DOM 5 · MIE-DOM 6
Capacidad efectiva del período:
  nominal × W − ocurrencias festivas − ranuras sin bloque en catálogo

Escalador:
  CapacidadAntes = 4W, CapacidadDespues = 5W  (nunca "4 → 5")

Planificador en LUN-VIE (patrones 1,2 / 3,4 / 5; demandas 2W, 2W, W):
  Sesiones planificadas = 5W · Infactibles = 0
  Suficiencia total NO implica reparto factible: sin patrones sembrados
  el reparto entre carriles puede fragmentar y dejar infactibles.
```

La tabla de arriba es **nominal**; cualquier desviación por festivo, catálogo o patrón debe reflejarse en el caso de prueba, no en la tabla.

---

## 5. Estrategia de pruebas y gates

1. `-Mode Build` por ola (stub, parser, escalera, escalador, UI).
2. Fast por componente:
   - `FullyQualifiedName~AgDiasVentana`
   - `FullyQualifiedName~AgEscaleraVentanas`
   - `FullyQualifiedName~AgEscaladorVentanaAula`
   - `FullyQualifiedName~AgPlanificadorOleadas`
   - `FullyQualifiedName~AgValidadorVentana`
   - `FullyQualifiedName~AgNodoArbolKeyMapper`
3. Persistencia: `AgEscaladorVentanaAulaPersistenciaTests` sobre SQLite temporal (patrón `MigrationV76Tests`), sin tocar la BD del usuario.
4. `-Mode Full` **solo** en la ola 6; `SUITE_SHRUNK` (exit 3) = fallo, no éxito.
5. Evidencia por tarea en `_evidence/<stamp>/` (no se commitea).

---

## 6. Riesgos y mitigaciones

| Riesgo | Mitigación |
|---|---|
| `Normalizar` deja de mapear alias ISO→etiqueta y cambia sugerencias de aulas persistidas en ISO | Claves ISO de legado en `Escalones`; `Siguiente("4,5,6,7")` sigue en `""`; caso de prueba explícito |
| `Evaluar` solo con `Siguiente` ya no "salta" un paso sin ganancia (cambio de comportamiento también para `VIE-SAB`) | Es la semántica pedida en 1703; caso dedicado y nota en el artículo |
| La recarga completa del perfil rebasa el alcance y rebaraja `_contextoActual` | Helper estrecho (DTO + `AplicarDimensiones` por `AulaRef`) en lugar de `CargarPerfilArbolSeleccionadoAsync`; prueba de flujo a nivel repositorio+mapper |
| Leyenda dentro del rango `x` del grupo B cambia la ventana al pulsarla | Tooltip no clicable + guarda `dy` en el handler; prohibido dibujar en `[DxLunJue, DxSep2)` sin guarda |
| Sábado contado dos veces si se aceptan `6M/6T` persistidos | Parser los rechaza (`Dias("6M")=[]`); solo rótulo de salida |
| Conflicto con el plan de ranuras 4→8 h (`AgDiasVentana`, `AgPlanificadorOleadas`, `AgValidadorVentana`) | Ejecución estrictamente secuencial; este plan primero y el otro rebasa sobre las etiquetas nuevas |
| Árbol sucio + `R-047 executing` invalidan preflight y reserva | Fase A: commit previo (e523b51), preflight 0 con `-AllowOneDrive`, remedir baseline (1172), escribir plan/handoff y filas R-057/R-058 |
| `Full` intermedio devuelve 3 por RED abiertos | `Full` solo en la ola 6; gates intermedios por filtro |

---

## 7. Fuera de alcance y precondiciones

**No se hace:** migración de BD, cambios de `csproj`/`App.cs`/Ribbon, radios nuevos, tokens `6M/6T` persistidos, cambios en R4/R5/R10, cambios en `AgPlanificadorOleadas.cs`/`AgValidadorVentana.cs`/`AgRanuras.cs`/`AgBandasDias.cs`/`AgParticionCarriles.cs`, build release, smoke test interactivo.

