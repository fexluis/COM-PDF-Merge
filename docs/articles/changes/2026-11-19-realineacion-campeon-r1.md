# Hotfix del Campeón: Forward Shifting por Carril (Plan R1)

**Fecha:** 2026-11-19  
**Rondas asociadas:** Ronda 482 (Fases 2129–2135), Ronda 483 (Fases 2136–2149)  
**Módulos afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica`

## 1. Contexto y Diagnóstico
Durante la corrida #120 del Algoritmo Genético de Horarios Académicos, el individuo campeón final finalizó la evolución acumulando violaciones duras:
`HARD=26 (R5×16, R6×5, R13×5)`.

El análisis causal confirmó que:
- Las violaciones R6 (`R6_SecuenciaAsignaturas`) y R13 (`R13_OleadasDispersas`) se originan por solapes temporales de oleadas sucesivas asignadas al mismo carril efectivo (`IdAula` y `DiasPatron`) dentro de una cohorte.
- El espacio en la ventana semestral es suficiente para alojar las asignaturas de forma estrictamente secuencial sin solape.
- La dinámica de mutaciones y cruces durante las generaciones no garantiza el orden cronológico no solapado al cierre.

## 2. Solución Aplicada (Plan R1)
Se introdujo una fase determinista de reparación post-evolución aplicada **exclusivamente al campeón**, previo al gate P08 y a la persistencia en base de datos:

1. **Helper Canónico de Carril:**
   `AgPlanificadorOleadas.ClaveCarril(AgGen g)` unifica la definición de carril efectivo compartida con `AgValidadorReglas`.
2. **Forward Shifting en Cascada:**
   `AgPlanificadorOleadas.RealinearOleadas(AgIndividuo, AgContexto)` agrupa los genes por `(Cohorte, Carril)` y desplaza las oleadas solapadas hacia el siguiente slice semanal disponible posterior a `prevFin`.
3. **Preservación Estricta de Invariantes I1/I2:**
   - Cada gen mantiene `HorasGen = FechasGen.Count * Unidad`.
   - Se erradicó el reparto residual arbitrario de horas que desalineaba fechas.
   - Si una asignatura no puede secuenciarse o no cabe en ventana, el carril completo no se muta (cálculo en dos fases).
4. **Hook con Rollback:**
   En `AgGeneradorHorarios.cs`, se clona profundamente el campeón antes de reparar. Si `hardPost >= hardPre`, se revierte al clon original.
5. **Observabilidad:**
   Telemetría estructurada `[AG-REPARACION]` con detalle de carriles evaluados, reparados, asignaturas desplazadas, fallos y evolución de violaciones duras.

## 3. Pruebas de Regresión y Cobertura
- Pruebas unitarias de casos de frontera en `AgPlanificadorOleadasTests`.
- Validación de erradicación de `R6_SecuenciaAsignaturas` y `R13_OleadasDispersas` en `AgValidadorReglasOleadasTests`.
- Prueba de integración end-to-end en `AgGeneradorHorariosTests`.
- Prueba de determinismo en `AgGeneradorHorariosReplayTests`.
- Prueba de invariantes I1/I2 y coherencia de fechas en `AgCoherenciaGrupoTests`.
