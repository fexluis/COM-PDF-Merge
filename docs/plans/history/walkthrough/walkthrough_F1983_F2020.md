<!-- historico congelado - generado por .agents/tools/split-history.ps1 - commit 9b3dfd4 - no editar -->
# walkthrough - Fases 1983-2020 (Rondas 450..456)

> Indice: [../../../walkthrough.md](../../../walkthrough.md) - Escision: 2026-11-18

---
### Verificación de Principios (Ronda 449)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1978 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEstadoCorrida.cs`, `AgCausaFallo.cs`, `AgFormGeneradorHorarios.cs` (código en disco; pendiente verificación) |
| F1979 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgGeneradorHorarios.cs`, `AgSugerenciaCobertura.cs` (código en disco; pendiente verificación) |

---

## Ronda 450 — Revisión contra disco y robustecimiento de conteos (Fases 1983–1986)

### Resumen de la Ronda
- **Fases 1983–1986:** Descuadre con signo en violaciones duras. Predicado de vacantes alineado a `AgIndividuo`. Limpieza de enum inalcanzable. Anclaje de parámetros genéticos en `AgFormGeneradorHorarios.cs`. Tests unitarios en `AgSugerenciaCoberturaTests.cs`.

### Verificación de Principios (Ronda 450)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1983-F1986 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgGeneradorHorarios.cs`, `AgFormGeneradorHorarios.cs`, `AgSugerenciaCoberturaTests.cs` (código en disco; pendiente verificación) |

---

## Ronda 451 — Trazabilidad de IdCorrida en telemetría de parada (Fase 1987)

### Resumen de la Ronda
- **Fase 1987:** Asignación de `IdCorrida = contexto.IdCorrida` al instanciar `AgResultadoCorrida` en `AgGeneradorHorarios.cs`.

### Verificación de Principios (Ronda 451)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1987 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgGeneradorHorarios.cs` (código en disco; pendiente verificación) |

---

## Ronda 452 — Revisión contra disco, robustecimiento de release y auditor de cierre (Fases 1988–1993)

### Resumen de la Ronda
- **Fases 1988–1990:** Alcance de variables en `build_installer.ps1`, validación estricta de versión/ancla y descriptor `SinFactibilidad` en `AgCorridaSelector.cs`.
- **Fases 1992–1993:** Implementación mecánica en `scripts/audit-closure.ps1` y sinceramiento de estados a `executing` en `REGISTRY.md`.

### Verificación de Principios (Ronda 452)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1988-F1990 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `scripts/build_installer.ps1`, `AgCorridaSelector.cs` (código en disco; pendiente verificación) |
| F1992-F1993 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `scripts/audit-closure.ps1`, `REGISTRY.md`, `walkthrough.md` (código en disco; pendiente verificación) |

---

## Ronda 453 — Diagnóstico forense campeón infactible, diálogo modal y clasificación vacantes (Fases 1994–1997)

### Resumen de la Ronda
- **Fase 1994:** Helper `AgDiagnosticoInfactibilidad.cs` y log `[AG-CAUSAS-DETALLE]`.
- **Fase 1995:** Resumen legible en diálogo modal y log visible de gate P08.
- **Fase 1996:** Atribución estricta de vacantes (sin centinela, log patrón y selección modal).
- **Fase 1997:** Tests unitarios y sincronización de registro.

### Verificación de Principios (Ronda 453)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1994 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgDiagnosticoInfactibilidad.cs`, `AgGeneradorHorarios.cs`, `ConversorPDF.XLL.csproj` (código en disco; pendiente verificación) |
| F1995 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.cs` (código en disco; pendiente verificación) |
| F1996 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFabricaIndividuos.cs`, `AgCoherenciaGrupo.cs` (código en disco; pendiente verificación) |
| F1997 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgDiagnosticoInfactibilidadTests.cs`, `AgMotivoVacanteTests.cs`, `REGISTRY.md` (código en disco; pendiente verificación) |

---

## Ronda 454 — Correcciones de compilación, tests de diagnóstico y blindaje de ejecución (Fases 1998–2009)

### Resumen de la Ronda
- **Fase 1998:** Corrección de `CS0103` usando `grp.Key` en lugar de `clave` en `AgCoherenciaGrupo.cs:493-494`.
- **Fase 1999:** Asignación de `HardViolations = 2/1` en fixtures de `AgDiagnosticoInfactibilidadTests.cs` para evitar early return de lista vacía en pruebas unitarias.
- **Fase 2000:** Filtrado exclusivo de violaciones `R1` y `R7` y descarte de `IdTercero <= 0` en pares de conflicto dentro de `AgDiagnosticoInfactibilidad.cs`.
- **Fase 2001:** Preservación de motivos declarados `SIN_PLAN_OLEADA` y `SIN_BLOQUE_NIVEL` antes de la comprobación estructural en `AgDiagnosticoInfactibilidad.cs`.
- **Fase 2002:** Aislamiento de llamadas de diagnóstico en `try/catch` defensivo con warning limitado en `AgGeneradorHorarios.cs`; elevación de DTOs de diagnóstico a `internal sealed class`.
- **Fase 2003:** Reincorporación de mensaje explícito de "No se materializarán sesiones en staging" y acotamiento de nombres de asignaturas (`Take(3)`) en `AgDiagnosticoInfactibilidad.cs` y `AgFormGeneradorHorarios.cs`.
- **Fase 2004:** Verificación y conservación de UTF-8 con BOM en `build_project.ps1`, `scripts/build_installer.ps1` y `scripts/audit-closure.ps1`.
- **Fase 2005:** Soporte para `EstancamientoMax = 0` usando `Math.Max(0, perfil.EstancamientoMax)` en `AgFormGeneradorHorarios.cs`.
- **Fase 2006:** Reemplazo de constante fija `EstancamientoMax = 20` por `AgParametrosPorDefecto.EstancamientoMax` en `AgArbolService.cs`.
- **Fase 2007:** Validación de `contexto.EsDocenteHabilitado(...)` en el cálculo de `poolHabilitado` en `AgDiagnosticoInfactibilidad.cs`.
- **Fase 2008:** Sustitución de `Write-Error` por `Write-Host` en rojo antes de `exit 1` en `scripts/audit-closure.ps1` para evitar terminación abrupta con `$ErrorActionPreference = 'Stop'`.
- **Fase 2009:** Resolución de `CS0136` por shadowing de `grp` renombrando parámetro lambda a `mg =>` en `AgCoherenciaGrupo.cs:485-486`.

### Verificación de Principios (Ronda 454)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1998, F2009 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgCoherenciaGrupo.cs:485-486, 493-494` (código en disco; pendiente verificación) |
| F1999 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgDiagnosticoInfactibilidadTests.cs:47, 65, 104, 121, 140, 160` (código en disco; pendiente verificación) |
| F2000, F2001, F2003, F2007 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgDiagnosticoInfactibilidad.cs:72-75, 150, 189-200, 205, 238-248, 329` (código en disco; pendiente verificación) |
| F2002 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgGeneradorHorarios.cs:337-350`, `AgDiagnosticoInfactibilidad.cs:17-31` (código en disco; pendiente verificación) |
| F2004 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `build_project.ps1`, `scripts/build_installer.ps1`, `scripts/audit-closure.ps1` (código en disco; pendiente verificación) |
| F2005 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.cs:2911` (código en disco; pendiente verificación) |
| F2006 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgArbolService.cs:83` (código en disco; pendiente verificación) |
| F2008 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `scripts/audit-closure.ps1:65-67` (código en disco; pendiente verificación) |

---

## Ronda 455 — Mutación a vacante explícita sin alterno, movimiento inverso y normalización de IdTercero=0 (Fases 2012–2015)

### Resumen de la Ronda
- **Fase 2012:** En `AgOperadorMutacion.cs:43-98`, ante mutación sin alterno disponible en pool, asignación explícita de `EsVacante = 1`, `IdTercero = 0` y diagnóstico de motivo con guarda de `Fijado` y sellado grupal incondicional.
- **Fase 2013:** En `AgAsignadorTercero.cs:195-230`, método estático `DiagnosticarMotivoVacante(AgGen, AgContexto)` para tipificar vacantes sin instanciación.
- **Fase 2014:** En `AgCoherenciaGrupo.cs:461,478`, normalización de vacantes a `IdTercero = 0`; en `AgValidadorReglas.cs:178-199`, relajación de `R9` para validar `IdTercero = 0` o centinela como vacante sin penalidad HARD.
- **Fase 2015:** En `AgOperadorMutacion.cs:78-95`, implementación de movimiento inverso: grupo vacante reintenta recuperar docente habilitado si surge disponibilidad en el pool.

### Verificación de Principios (Ronda 455)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2012, F2015 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgOperadorMutacion.cs:43-115`, `AgOperadorMutacionTests.cs:75-180` |
| F2013 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgAsignadorTercero.cs:195-235` |
| F2014 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgCoherenciaGrupo.cs:461,478`, `AgValidadorReglas.cs:178-199`, `AgValidadorReglasTests.cs:247-275`, `AgCoherenciaGrupoTests.cs:254,292,557-575` |

---

## Ronda 456 — Blindaje antidegeneración, preservación de causas no asignables y determinismo RNG (Fases 2016–2020)

### Resumen de la Ronda
- **Fase 2016:** En `AgOperadorMutacion.cs`, guarda de docente vigente: sólo liberar a vacante si la asignación actual es inviable y no hay alterno en pool.
- **Fase 2017:** En `AgOperadorMutacion.cs`, guarda `MotivosNoAsignables` (`SIN_PLAN_OLEADA`, `SIN_BLOQUE_NIVEL`) que excluye grupos estructuralmente incompletos de la mutación.
- **Fase 2018:** En `AgContexto.cs`, `AgValidadorReglas.cs`, `AgReglaViolacion.cs`, `AgPesosFitness.cs`, `AgGeneradorHorarios.cs` y `AgCompuertaPromocion.cs`: tope estructural `MaxVacantesForzadas`, regla dura `R15_VacanteExcedida`, telemetría de vacantes forzadas en `[AG-CAUSAS]` y validación estricta de vacantes consentidas por asignatura en compuerta de promoción.
- **Fase 2019:** En `AgCoherenciaGrupo.cs:93,199`, normalización de vacantes canónicas a `IdTercero = 0`; en `AgReglaViolacion.cs` y `AgValidadorReglas.cs`, precisión semántica `R9_VacanteConDocenteReal = 26`.
- **Fase 2020:** En `AgOperadorMutacion.cs`, sorteo RNG con cortocircuito dentro de la condición de mutación elegible para preservar la secuencia determinista.

### Verificación de Principios (Ronda 456)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2016, F2017, F2020 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgOperadorMutacion.cs:15-135`, `AgOperadorMutacionTests.cs:105-240` |
| F2018 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgContexto.cs:90-95`, `AgValidadorReglas.cs:705-725`, `AgValidadorReglasTests.cs:275-300`, `AgReglaViolacion.cs:40`, `AgPesosFitness.cs:66`, `AgGeneradorHorarios.cs:320-335`, `AgCompuertaPromocion.cs:255-275, 355` |
| F2019 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgCoherenciaGrupo.cs:93,199`, `AgReglaViolacion.cs:32`, `AgValidadorReglas.cs:186` |

---

