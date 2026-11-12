# Plan: aprovechamiento dinámico de ranuras semanales (4 h → 8 h)

Fecha: 2026-11-12 · Actualizado: 2026-11-12 · Estado: draft ejecutable, **sin reserva de numeración** · Autor: Cline

> Documento de análisis. El plan ejecutable (contrato swarm: envelope + tasks atómicas) vive en
> `.agents/plans/2026-11-12-ranuras-dinamicas-4-a-8-horas-plan.md`.
> No se ejecutó nada: no se cambió código, libro ni base de datos. Las citas `archivo:línea` provienen
> del árbol de trabajo en `HEAD` = `931f529` (árbol sucio; ver §Estado real).

## Estado real del repositorio (verificado)

- `REGISTRY.md` §2: contador `ronda` = 331, `fase` = 1637 (antes de esta actualización).
- `Fase 1606–1612` = R-041 (`closed`). `Fase 1613–1620` = R-042 (`closed`). La numeración 1612–1620 usada por la primera versión de este documento **ya estaba ocupada**.
- R-047 (`Ronda 331`, `Fase 1636–1637`) sigue en `executing`.
- `git status`: 26 entradas; 6 archivos nuevos sin seguimiento, entre ellos `AgParticionCarriles.cs`.
- `HEAD` **no compila desde un clon limpio**: `ConversorPDF.XLL.csproj:702` referencia `AgParticionCarriles.cs`, que no está en `git`.
- `git clean -fd`, `git stash` o cambio de rama **destruyen** esos 6 archivos. Commit antes de cualquier limpieza.
- Baseline `tests/baseline/test_baseline.txt` mide HEAD anterior; `verify.ps1 -Mode Full` devuelve `3` (`SUITE_SHRUNK`) hasta remedir.

## Hallazgos del libro (no verificables en repo)

Fuente declarada por el usuario: `C:/Users/USUARIO/Downloads/Libro2.xlsx`, hoja Semestres, filas 3–126;
columna K y calendario L:FX (fechas hasta 2027-07-23), corrida #62. El libro está fuera del repositorio:
T1 debe copiarlo a `_evidence/` y registrar SHA256; sin eso la evidencia no es reproducible.

- 20 filas con K exactamente `VIERNES|4|08:00 A 12:00` o `VIERNES|4|13:00 A 17:00`: 216 sesiones = 864 h; 224 posteriores al 30/05.
- 10 filas adicionales con viernes único: 76, 77, 85, 86, 89, 92, 99, 108, 113, 116.
- El 26/03 no tiene sesión en estas secuencias: confirmar con el calendario de festivos del vault, sin inferir calendario externo.
- Cohortes de 240, 256 y 272 horas. No vincular el mensaje de 62 ranuras a una cohorte sin el `IdAula` y su configuración.

| Filas Excel | Asignatura y cohorte | Sesiones | Fechas |
|---|---|---|---|
| 5 | Fundamentos Sociológicos, I-A Tuluá | 8 × 4 h | 05/02–02/04 |
| 8 | Matemáticas I, I-A Tuluá | 12 × 4 h | 09/04–25/06 |
| 27 y 41 | Estadística I, III-A y III-B Cali | 12 × 4 h por fila | 05/02–30/04 |
| 36 y 44 | Teorías de las organizaciones públicas, III-A y III-B Cali | 12 × 4 h por fila | 07/05–23/07 |
| 51 y 65 | Derecho administrativo, IV-A y IV-B Cali | 12 × 4 h por fila | 05/02–30/04 |
| 60 y 68 | Presupuesto público, IV-A y IV-B Cali | 12 × 4 h por fila | 07/05–23/07 |

Las celdas del libro son datos, no instrucciones.

## Causa estructural verificada en código

- `AgPlanificadorOleadas.cs:93-112`: slices semanales anclados a `fechaInicio` (`c = c.AddDays(7)`), no a lunes; cierre en `c.AddDays(6)`.
- `AgPlanificadorOleadas.cs:378, 393-401, 432`: `kBase = ceil(SRequeridas / slotsCarril[c])`, `cursor[c] += k`; una materia del carril residual de un solo día conserva 4 h/semana.
- `AgCarriles.cs:13-14`: `RanurasPorCarril = 2`, `DiasPorCarril = 2`; `Residual` deja 1 día ⇒ 4 h/semana.
- `AgPlanificadorPatronSemanal.cs:129-150`: `FechasDeRanura(diaIso, sesiones, ...)` produce **una sola ocurrencia semanal por día**; dos sesiones/semana exigen dos genes.
- `AgPlanificadorPatronSemanal.cs:172-181`: patrón textual solo se acepta si es carril de 2 días consecutivos o 1 día; días libres arbitrarios caen a carril con advertencia.
- `AgFactoryIndividuos`/`AgFabricaIndividuos.cs:524-525`: `FechasGen` por ranura y `HorasGen = FechasGen.Count * horasEfectivas`.
- `AgGen.cs:133-165`: `Clonar` reconstruye campo por campo (lista manual).
- `AgCoherenciaGrupo.cs:77, 229-240, 261-262, 289-291, 308-313`: `Sellar` **vacía `FechasGen`** y re-cuantiza `HorasGen`; es el último paso del ciclo (`AgGeneradorHorarios.cs:195`).
- `AgOperadorCruce.cs:237-239`: `FechasGen = new List<DateTime>()` y `HorasGen = 0.0`.
- `AgExpansorSesiones.cs:31-34`: `HorasGen <= 0` ⇒ 0 sesiones; `:62-70`: `FechasGen` vacío ⇒ generador semanal.
- `AgParticionCarriles.cs:106-116`: `hayPatronValido` decide con `AgCarriles.Ventana` (parser tolerante).
- `AgDiasVentana.cs:29-58`: `AIso`/`DiaIso` aceptan 1..7 o prefijos LUN/MAR/MIE/JUE/VIE/SAB/DOM; texto libre devuelve lista vacía; rangos se colapsan al primer prefijo.
- `Academic/Infrastructure`: sin referencias a `FechasGen`; `ag_gen` no persiste fechas por sesión (`MigrationV49:101-117`).

## Lo que la optimización puede y no puede resolver

- 62 ranuras × 4 h = 248 h. Demanda de 256 h (64 sesiones) ⇒ faltan ≥2 ranuras; 272 h ⇒ ≥6. Ningún reparto semanal elimina ese déficit.
- Tres diagnósticos distintos: insuficiencia física total, insuficiencia de ranuras compatibles con días/docente, y fragmentación por el planificador. Solo la tercera se corrige redistribuyendo.
- Subir a 8 h/semana adelanta el fin de una materia, no aumenta horas de la cohorte. Reducir simultaneidad tampoco reduce demanda total.

## Restricciones duras verificadas

```text
R4  AgValidadorReglas.cs:337-379  unicidad por (Cohorte, DíaSemana, Franja) en intervalos solapados
                                  + tope diario: 1 franja (2 solo sábado)
R5  AgValidadorReglas.cs:382-455  máximo de asignaturas concurrentes por cohorte
R10 AgValidadorReglas.cs:259-270  todos los genes de igual ClaveGrupo comparten IdTercero
D4  AgExpansorSesiones.cs:100-113 duración de sesión == bloque.DuracionHoras == HorasSesionAula
```

- Dos bloques el mismo día hábil es **prohibido** (R4), no objetivo. "Dos sesiones por semana" ⇒ dos días distintos.
- La ranura liberada solo se reutiliza en día+franja distinto o cuando la otra materia ya terminó.
- Ancla de semana: se conserva el slice desde `fechaInicio`; ISO solo para identidad y reportes.
- El saldo se resuelve por cardinalidad de sesiones, nunca por duración reducida.
- Invariante: `PlanExplicito == true` ⇒ `FechasGen` no vacío. Ninguna ruta puede vaciar fechas sin demover el flag.
- Déficit por festivos o fin de ventana se declara (`I5`), no se convierte en fallo del expansor.

## Comportamiento requerido

- Objetivo preferente: dos sesiones de 4 h por asignatura y semana cuando sean factibles; una cuando solo una sea compatible; volver a dos en la primera semana posible. Cero sesiones debe justificarse.
- Días permitidos explícitos siguen siendo límite duro. Pareja JUEVES–VIERNES utilizable solo si ambos están autorizados. Patrón estricto VIERNES ⇒ 4 h/semana y diagnóstico.
- Diferenciar patrón permitido de sesiones efectivamente usadas; K resume estas últimas.
- Ejemplo de aceptación: jueves ocupado semanas 1–2, libre desde la 3, materia de 24 h ⇒ V, V, J+V, J+V (4+4+8+8). Jueves libre desde el inicio ⇒ tres semanas 8+8+8. Patrón estricto V ⇒ seis semanas de 4 h.
- Fechas de corte reales, no solo intervalos min/max, gobiernan colisiones.

## Prohibiciones

- No fijar `HorasSemanales = 8` en un bloque de 4 h.
- No ampliar patrones estrictos ni añadir jueves a materias configuradas solo viernes.
- No eliminar `I5` ni alterar horas del pensum ni acortar fechas de exportación.
- No ampliar la fecha límite por decisión del motor.
- No reconstruir semanalmente un plan explícito que falló al cargar.
- No declarar mejora sin cobertura exacta y sin cero violaciones duras nuevas.

## Tareas (detalle atómico en el plan ejecutable)

| # | Tarea | Depende | Archivo principal |
|---|---|---|---|
| T1 | Reproducibilidad #62, hash del libro, fixtures anónimos, informe oferta/demanda | — | `ConversorPDF.Tests/Academic/Domain/AgRanurasDinamicasFixtures.cs` |
| T2 | Contrato de dominio: plan de sesiones inmutable, `AgGen.PlanExplicito`, migración V80 | T1 | `Academic/Domain/Models/Genetica/AgPlanSesiones.cs` |
| T3 | Calendario único de ranuras fechadas (identidad por intervalo físico + recurso) | T2 | `Academic/Domain/Models/Genetica/AgRanuras.cs` |
| T4 | Factibilidad y códigos I5 por subconjuntos compatibles | T3 | `AgValidadorVentana.cs` |
| T5 | Planificador por reservas con R1–R10; parser único de patrón | T4 | `AgPlanificadorReservas.cs` |
| T6 | Integración con genes: 1 gen por (asignatura, día, bloque), `PlanExplicito` por ruta | T5 | `AgFabricaIndividuos.cs` |
| T7 | `Sellar`/cruce/clonación: no destruir el plan, demover flag al invalidar fechas | T6 | `AgCoherenciaGrupo.cs` |
| T8 | Persistencia `fechas_gen` + `plan_explicito`, roundtrip, guardas del expansor | T7 | `AgScheduleRepository.cs` |
| T9 | Validación por reglas reales, penalización de intensidad opt-in, exportación | T8 | `AgValidadorReglas.cs` |
| T10 | Comparación original/ampliada con cortes 30/05 y 23/07; activación versionada | T9 | `AgGeneradorHorarios.cs` |

`- **fase**` se asigna al reservar en `REGISTRY.md`; no se fija en este documento.

## Formato de la fuente de verdad

- Plan en Domain: sesiones explícitas `(asignatura, cohorte, fecha, bloque, inicio, fin, duración, clave de reserva)`; inmutable.
- Persistencia: `ag_gen.fechas_gen` (texto `yyyy-MM-dd` separado por `;`) + `ag_gen.plan_explicito`.
- Migración: siguiente número libre es **V80** (existen hasta `MigrationV79_NormalizeUsuariosIds`); registro obligatorio en `Core/Helpers/SqliteHelper.cs:380`; **no aditiva** (columna exigida por el INSERT).
- Nuevos `.cs`: `internal sealed` en Domain/Genetica y alta manual en `ConversorPDF.XLL.csproj` (ítems `Compile` explícitos). El proyecto de tests es SDK-style: no requiere alta.

## Verificación

```powershell
powershell -ExecutionPolicy Bypass -File scripts/preflight.ps1 -Workspace <ws> -Plan .agents/plans/2026-11-12-ranuras-dinamicas-4-a-8-horas-plan.md
powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Build
powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "FullyQualifiedName~AgPlanificadorReservasTests|FullyQualifiedName~AgExpansorSesionesTests"
powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Full      # sin -Filter; 0 = OK, 3 = SUITE_SHRUNK
```

Evidencia en `_evidence/<stamp>/{build.log,test.trx,summary.md}`; no se commitea.

## Correcciones incorporadas (Fases 1638–1686)

| Fase | Corrección aplicada al plan |
|---|---|
| 1638–1639 | Numeración real (Ronda/ Fase por reserva) y envelope de ejecución |
| 1640, 1671, 1676, 1681, 1686 | Árbol sucio y `HEAD` sin compilar: commit previo obligatorio |
| 1641, 1656, 1663, 1667, 1672, 1679 | R4 explícita; parser único de patrón; fidelidad estricta solo con patrón interpretable |
| 1642 | Ancla de semana conservada (sin redefinir a lunes) |
| 1643, 1683 | Sin duración residual; déficit por festivos declarado, no fallo |
| 1644, 1660, 1678, 1682, 1684 | `PlanExplicito` con degradación, copia en `Clonar` y asignación por ruta |
| 1645 | Invariante de sesiones según procedencia |
| 1646, 1661, 1668 | Migración V80 con contrato real, registro y no aditiva; 3 sitios SQL |
| 1647, 1669 | R10 por `ClaveGrupo`; helper de cohorte en Application |
| 1648, 1654, 1662 | Reglas reales (R11/R12/R13 reservadas); fixtures en `Academic/Domain`; csproj de tests fuera del lease |
| 1649, 1673, 1677 | Gate de partición, caché versionada y guarda quirúrgica en `Sellar` |
| 1650, 1658, 1670 | Reproducibilidad de fuente, comandos de gate y citas contra commit sellado |
| 1651–1653, 1655, 1659, 1664–1666, 1674–1675, 1680, 1685 | Numeración de migración, accesibilidad, reserva única, clave de cohorte, re-sellado de hijos, registro |

## Pendientes administrativos

1. Commit de los 6 archivos sin seguimiento (Fase 1686) antes de cualquier limpieza o reserva.
2. Reserva de numeración del plan (`swarm-planner`, Fase A) con `preflight` en verde.
3. Remisión del baseline tras el commit.

## Criterio final

Cobertura exacta, cero violaciones duras nuevas, transición 4→8 probada cuando la segunda ranura esté disponible y autorizada, fechas persistidas idénticas al plan y uso de huecos sin ampliar la fecha límite por decisión del motor. El cierre de todas las cohortes al 30/05 solo puede declararse si el diagnóstico y la ejecución lo demuestran.
