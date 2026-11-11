# Ajuste del Validador de Reglas AG para Soporte de Oleadas Temporales (Fases 1568–1574)

## 1. Visión General
En el motor genético académico de **ConversorPDF**, la programación de asignaturas por cohortes avanzadas frecuentemente utiliza el modelo de **oleadas consecutivas**: asignaturas modulares que comparten el mismo día y bloque horario dentro de la misma aula física pero en rangos de fechas disyuntos (por ejemplo, Asignatura A de febrero a abril, y Asignatura B de abril a mayo).

Antes de esta versión, el validador `AgValidadorReglas` evaluaba colisiones únicamente contra la plantilla semanal (`DiaSemana`, `IdBloqueHorario`), generando falsos positivos de colisión dura (R1, R2, R3, R4, R14) en corridas válidas (como el caso real #48 con 101 HARD espurias pese a 0 solapes reales).

Esta ronda implementa la validación temporal basada en ventanas reales de fechas estimadas (`FechaInicioEstimada`, `FechaFinEstimada`) con semántica conservadora.

---

## 2. Componentes y Cambios

### Helper `AgVentanaSolape` (Fase 1568)
- Proporciona el método universal `SeSolapan(DateTime? iniA, DateTime? finA, DateTime? iniB, DateTime? finB)`.
- **Semántica conservadora**: si cualquiera de los extremos es nulo, asume solape (`true`), garantizando compatibilidad retrospectiva con asignaciones sin fechas explícitas.
- Intervalos cerrados `[iniA, finA]` y `[iniB, finB]`: solapan si `iniA <= finB && iniB <= finA`.
- Proporciona `AgruparPorSolape(IEnumerable<T>)` para formar componentes conexas de colisión conservando mensajes informativos.

### Reglas R1, R2, R3, R4 y R14 (Fase 1569)
- **R1 (Docente)**: Dos genes con el mismo docente (`IdTercero`) en el mismo día y bloque sólo emiten colisión si sus ventanas temporales se solapan (`AgVentanaSolape.SeSolapan`).
- **R2 (Aula)**: Dos genes en la misma aula física (`IdAula`) en el mismo día y bloque sólo colisionan si se solapan en fechas.
- **R3 (Grupo/Cohorte)**: Dos genes de la misma cohorte (`Programa + Semestre + Grupo`) en el mismo bloque sólo colisionan si se solapan en fechas.
- **R4 (Dos Asignaturas Mismo Día y Franja)**: Permite múltiples asignaturas en la misma franja y día siempre que no coexistan temporalmente en las mismas fechas.
- **R14 (Bloque Solapado por Intervalo de Horas)**: Bloques contiguos o solapados en horas sólo emiten infracción si sus rangos de fechas se intersecan.

### R5 Concurrencia de Asignaturas por Día Real (Fase 1570)
- Sustituye la agrupación por semana ISO (que fragmentaba asignaturas en los bordes de semana y podía emitir violaciones artificiales) por una evaluación continua día a día a lo largo del intervalo `[min(FechaInicio), max(FechaFin)]` de la cohorte.
- Cuenta la cantidad de asignaturas activas en cada día real contra `maxCarriles` de la ventana. Emite una única violación consolidada por tramo contiguo de sobrecarga.

### R4 Estado de Aula Permitido (Fase 1571)
- Admite tanto `AgEstadoAula.Disponible` como `AgEstadoAula.Aprobado` como estados válidos para programar.

### Cobertura de Asignaturas Requeridas (Fase 1572)
- Acota la exigencia de asignaturas requeridas al conjunto de asignaturas vinculadas a las aulas que participan en la corrida (`contexto.AulaAsignaturas`), previniendo penalizaciones por asignaturas de aulas excluidas del alcance.

---

## 3. Pruebas y Verificación (Fase 1573 / 1574)
- Suite completa de regresión para oleadas: `AgValidadorReglasOleadasTests.cs` (6 escenarios clave).
- Casos unitarios para `AgVentanaSolapeTests.cs` (5 tests).
- 1123 pruebas automáticas pasando al 100% (`scripts/verify.ps1 -Mode Fast`).
## 4. Verificación sobre datos reales (corrida #48)

Evidencia decisiva obtenida ejecutando el **validador nuevo** contra los 30 genes reales del individuo #33
(copia del vault, solo lectura), reconstruyendo el contexto desde `ag_aula`, `ag_bloque_horario`,
`asignaturas` y el pool docente:

| Comprobación sobre el vault real | Resultado |
|---|---|
| Pares de genes con solape **real** de fechas — misma aula + bloque | **0** |
| Pares con solape real — misma cohorte + bloque | **0** |
| Pares con solape real — mismo docente + bloque | **0** |
| Pares con bloques distintos, horas solapadas y fechas cruzadas | **0** |
| `AgValidadorReglas.Evaluar` con el código corregido | **HARD = 0 · SOFT = 0 · 0 filas** |

Estructura de oleadas verificada (la misma franja semanal usada por asignaturas en semanas distintas):

| Cohortes / franja | Oleada 1 | Oleada 2 | Oleada 3 |
|---|---|---|---|
| IX-A · Jue/Vie 08:00–12:00 (bloques 24 y 29) | 148 · PROYECTO DE FUTURO IV · 02/02→15/03 | 150 · PROCESOS ECONÓMICOS TERRITORIALES · 16/03→03/05 | 152 · GERENCIA PÚBLICA INTEGRAL · 04/05→28/06 |
| IX-A · Sáb 08:00–12:00 y 13:00–17:00 (bloques 5 y 6) | 149 · PROYECTOS DEL DESARROLLO · 02/02→05/04 | 151 · GESTIÓN PARA EL DESARROLLO · 06/04→19/04 | 153 · ELECTIVA V · 20/04→07/06 |
| X-A · Vie 08:00–12:00 (bloque 29) | 154 · SEMINARIO ELECTIVO II · 02/02→05/04 | 156 · TRABAJO DE GRADO · 06/04→19/04 | — |

> Nota: `ag_gen.dias_patron` **sí** está poblado (`"4,5"` y `"6"`); por eso R6/R13 agrupan correctamente por carril.
> Un arnés que omita esa columna colapsa todos los genes en el carril `DEF` y produce falsos positivos de R6/R13
> (lección para futuras verificaciones).

## 5. Corrección del fixture de pruebas (dentro de Fase 1573/1574)

- `AgValidadorReglasOleadasTests` declaraba `Franja = "M"` / `"T"`; el contrato canónico es `MANANA` / `TARDE`
  (`AgFranjaJornada`), lo que hacía fallar `Oleadas_IXA_Sabado_DosBloques_SinViolaciones` con
  `R11_DuracionSesionInvalida`. Corregido a `MANANA`/`TARDE`.
- Las aserciones de “sin violaciones” ahora incluyen el listado de reglas/mensajes detectados para diagnóstico.


