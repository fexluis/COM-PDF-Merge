# Indice de Historial - walkthrough.md

Escision: 2026-11-18 - commit `9b3dfd4` - trozo objetivo: 1000 lineas

Contenido historico congelado en `docs/plans/history/walkthrough/`. Orden append-only: no se reordena ni se edita.

| Archivo | Fases | Rondas | Lineas |
|---|---|---|---|
| [walkthrough_F0620_F0710.md](docs/plans/history/walkthrough/walkthrough_F0620_F0710.md) | 620-710 | 102..117 | 1239 |
| [walkthrough_F0707_F0799.md](docs/plans/history/walkthrough/walkthrough_F0707_F0799.md) | 707-799 | 145, 146, 147, 148, 149, 150 | 1233 |
| [walkthrough_F0800_F0885.md](docs/plans/history/walkthrough/walkthrough_F0800_F0885.md) | 800-885 | 151..169 | 1246 |
| [walkthrough_F0883_F0990.md](docs/plans/history/walkthrough/walkthrough_F0883_F0990.md) | 883-990 | 170..191 | 1248 |
| [walkthrough_F0990_F1107.md](docs/plans/history/walkthrough/walkthrough_F0990_F1107.md) | 990-1107 | 192..205 | 1246 |
| [walkthrough_F1103_F1347.md](docs/plans/history/walkthrough/walkthrough_F1103_F1347.md) | 1103-1347 | 232 | 1232 |
| [walkthrough_F1342_F1509.md](docs/plans/history/walkthrough/walkthrough_F1342_F1509.md) | 1342-1509 | 264..286 | 1243 |
| [walkthrough_F1510_F1629.md](docs/plans/history/walkthrough/walkthrough_F1510_F1629.md) | 1510-1629 | 287..328 | 1250 |
| [walkthrough_F1613_F1819.md](docs/plans/history/walkthrough/walkthrough_F1613_F1819.md) | 1613-1819 | 329..398 | 590 |
| [walkthrough_F1820_F1982.md](docs/plans/history/walkthrough/walkthrough_F1820_F1982.md) | 1820-1982 | 399..449 | 1247 |
| [walkthrough_F1983_F2020.md](docs/plans/history/walkthrough/walkthrough_F1983_F2020.md) | 1983-2020 | 450..456 | 137 |
| [`walkthrough.md`](walkthrough.md) (vivo) | posteriores al corte | ultimas 5 | 128 |

## Busqueda por fase
Toda `Fase N` se localiza en el primer archivo cuyo rango la contenga; ante solape de rangos, usar la columna Rondas.

## Citas historicas
Las citas `walkthrough.md:NNNN` anteriores al corte apuntan al archivo monolitico; su ubicacion actual esta en [LINE_MAP.md](docs/plans/history/LINE_MAP.md).

---

## Ronda 457 — Activación de tope estructural, exclusión de vacantes consentidas, compuerta estricta y calendario verificable (Fases 2021–2025)

### Resumen de la Ronda
- **Fase 2021:** Cálculo de déficit estructural `CalcularDeficitEstructural(contexto)` en `AgValidadorVentana.cs` y cableado a `contexto.MaxVacantesForzadas` durante pre-vuelo P07 e inicialización de optimización.
- **Fase 2022:** Exclusión de `AsignaturasVacantesConsentidas` en el contador de vacantes forzadas para la regla `R15_VacanteExcedida` en `AgValidadorReglas.cs` y en la telemetría `[AG-CAUSAS]` de `AgGeneradorHorarios.cs`.
- **Fase 2023:** En `AgCompuertaPromocion.cs`, cálculo incondicional de asignaturas no consentidas, rechazo ante vacantes no consentidas si `!permitirVacantes`, y registro/trazabilidad explícita en `corrida.VacantesConsentidas` con warning de log si se autoriza por justificación.
- **Fase 2024:** En `AgReglaViolacion.cs`, eliminación de alias redundante 26 (`R9_VacanteSinCentinela`), consolidando `R9_VacanteConDocenteReal = 26` en modelo de dominio, evaluador de fitness y tests.
- **Fase 2025:** En `AgOperadorMutacion.cs`, verificación de calendario mediante `AgCalendarioConflictos.Crear(genes, contexto).IsSuccess` para evitar enmascarar `R14_BloqueHorarioNoAsignado`, y comparación case-insensitive en `EsNoAsignable`.

### Verificación de Principios (Ronda 457)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2021, F2022 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgValidadorVentana.cs:25-35, 585-625`, `AgValidadorReglas.cs:710-725`, `AgGeneradorHorarios.cs:103-110, 325-340`, `AgValidadorReglasTests.cs:280-325` |
| F2023 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgCompuertaPromocion.cs:255-295, 345-360` |
| F2024 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgReglaViolacion.cs:30-35`, `AgPesosFitness.cs:56-62`, `AgValidadorReglasTests.cs:240-280` |
| F2025 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgOperadorMutacion.cs:15-25, 60-120`, `AgOperadorMutacionTests.cs:240-315` |

---

## Ronda 458 — Precisión de déficit estructural, desacoplamiento de factibilidad, optimización de mutación, compuerta no mutable y granularidad de consentimiento (Fases 2026–2030)

### Resumen de la Ronda
- **Fase 2026:** Corrección de `CalcularDeficitEstructural` en `AgValidadorVentana.cs`: agrupamiento por `IdAsignatura`, exclusión de consentidas, conteo de aulas distintas (`Distinct().Count()`) y retorno `-1` ante contexto incompleto.
- **Fase 2027:** Eliminación de efectos secundarios en `ValidarFactibilidad` (`AgValidadorVentana.cs`), garantizando pureza de solo lectura sin mutar `MaxVacantesForzadas`.
- **Fase 2028:** En `AgOperadorMutacion.cs`, optimización de ruta crítica ejecutando `AgCalendarioConflictos.Crear` únicamente en rama `else` cuando no se encuentra candidato alterno.
- **Fase 2029:** En `AgCompuertaPromocion.cs`, eliminación de mutación autogenerada sobre `corrida.VacantesConsentidas`; requerimiento de autorización explícita sin alterar entidades persistidas.
- **Fase 2030:** En `AgContexto.cs`, incorporación de `ClavesVacantesConsentidas` y helper `EstaConsentida(idAsignatura, idSede)` para soportar granularidad a nivel cohorte/sede; cableado en `AgValidadorReglas.cs` y `AgGeneradorHorarios.cs`.

### Verificación de Principios (Ronda 458)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2026, F2027 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgValidadorVentana.cs:25-35, 590-630`, `AgGeneradorHorarios.cs:103-110`, `AgValidadorVentanaTests.cs:585-635` |
| F2028 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgOperadorMutacion.cs:75-120`, `AgOperadorMutacionTests.cs:240-315` |
| F2029 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgCompuertaPromocion.cs:240-295, 330-365` |
| F2030 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgContexto.cs:130-165`, `AgValidadorReglas.cs:710-725`, `AgGeneradorHorarios.cs:330-335`, `AgValidadorReglasTests.cs:320-350` |

---

## Ronda 459 — Concurrencia temporal en déficit, proveedor de autorización de vacantes, unificación de consentimiento, productor de claves y cobertura de pruebas (Fases 2031–2035)

### Resumen de la Ronda
- **Fase 2031:** En `AgValidadorVentana.cs`, `CalcularDeficitEstructural` calcula la concurrencia temporal considerando los carriles/franjas activos (`Math.Ceiling(aulasDistintas / carriles)`), corrigiendo el cap inerte (92 vs real ≈8).
- **Fase 2032:** En `AgPromotorVinculaciones.cs`, parseo de materias autorizadas con `ParseAutorizacionVacantes` y transmisión a `AgCompuertaPromocion.ValidarPromocion`. En `AgFormGeneradorHorarios.cs`, enriquecimiento de `justificacionVacante` con las asignaturas no consentidas y provisión de candidatos autorizables a `EsPromovible`.
- **Fase 2033:** Unificación de predicados de consentimiento con `PermiteVacantes` y `EstaConsentida(idAsignatura, idSede)` en `AgCompuertaPromocion.cs` y `AgValidadorVentana.cs`, eliminando rechazos espurios en vacantes consentidas por sede.
- **Fase 2034:** En `AgFormGeneradorHorarios.cs`, población de `ClavesVacantesConsentidas` en pre-vuelo y soporte bidireccional en `ParametrosJson["vacantesConsentidas"]` (int y `asig:sede`). Retirado el chequeo redundante en `EstaConsentida`.
- **Fase 2035:** Cobertura de pruebas en `AgValidadorVentanaTests.cs` (concurrencia disjunta vs aulas totales) y en `AgCompuertaPromocionTests.cs` (autorización explícita, autorización incompleta, claves de sede y parseo).

### Verificación de Principios (Ronda 459)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2031 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgValidadorVentana.cs:592-625`, `AgValidadorVentanaTests.cs:588-645` |
| F2032 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPromotorVinculaciones.cs:135-145, 1030-1065`, `AgFormGeneradorHorarios.cs:3360, 4460, 4705, 4770, 4875, 4960, 4975` |
| F2033 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgCompuertaPromocion.cs:258-275, 342-364`, `AgContexto.cs:165`, `AgValidadorVentana.cs:605-612` |
| F2034 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgContexto.cs:140-160`, `AgFormGeneradorHorarios.cs:1695-1710, 3345-3355, 4440-4455`, `AgCompuertaPromocion.cs:390-405` |
| F2035 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgValidadorVentanaTests.cs:620-650`, `AgCompuertaPromocionTests.cs:235-315` |

---

## Ronda 460 — Concurrencia por cohortes en déficit estructural, diálogo de autorización, parser fail-closed de justificación, preservación de claves en corrida y cobertura de pruebas (Fases 2036–2040)

### Resumen de la Ronda
- **Fase 2036:** En `AgValidadorVentana.cs`, `CalcularDeficitEstructural` adopta la cota defendible por cohortes (aula, semestre, grupo, sede) sin divisor arbitrario, corrigiendo el estrangulamiento de liberaciones (`MaxForzadas >= 8`).
- **Fase 2037:** En `AgFormGeneradorHorarios.cs`, integración de `FormAutorizacionVacantes.cs` para autorización explícita de asignaturas vacantes previo a la promoción. `btnPromoverProduccion.Enabled` y `EsPromovible` se vinculan a `_vacantesAutorizadas`, eliminando autoautorizaciones automáticas.
- **Fase 2038:** En `AgPromotorVinculaciones.cs`, `ParseAutorizacionVacantes` pasa a modo fail-closed, requiriendo token explícito (`asignaturas:`, `asig=`, `materias:`) para evitar autorizar números fortuitos de texto libre.
- **Fase 2039:** En `AgCompuertaPromocion.cs`, `VacantesConsentidasEnCorrida` devuelve tupla `(HashSet<int> Asignaturas, HashSet<string> Claves)`, previniendo que la recarga de corridas degrade `asig:sede` a wildcards por asignatura.
- **Fase 2040:** Pruebas unitarias actualizadas en `AgValidadorVentanaTests.cs` (concurrencia por cohorte/sede y caso corrida 109) y en `AgCompuertaPromocionTests.cs` (autorización parcial y parser fail-closed).

### Verificación de Principios (Ronda 460)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2036 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgValidadorVentana.cs:592-628`, `AgValidadorVentanaTests.cs:589-665` |
| F2037 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `FormAutorizacionVacantes.cs:1-120`, `AgFormGeneradorHorarios.cs:88, 3371, 3461, 3496, 4478, 4672, 4734, 4830-4860, 4955`, `AgCompuertaPromocionTests.cs:300-325` |
| F2038 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPromotorVinculaciones.cs:1040-1070`, `AgCompuertaPromocion.cs:285`, `AgCompuertaPromocionTests.cs:325-345` |
| F2039 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgCompuertaPromocion.cs:258-265, 350-365, 375-430`, `AgFormGeneradorHorarios.cs:4440-4450` |
| F2040 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgValidadorVentanaTests.cs:620-665`, `AgCompuertaPromocionTests.cs:300-345` |

---

## Ronda 461 — Registro de nuevos archivos en csproj, cohorte académica pura en déficit, aislamiento de permisos entre corridas, compuerta solicitable y granularidad completa en justificación (Fases 2041–2045)

### Resumen de la Ronda
- **Fase 2041:** En `ConversorPDF.XLL.csproj`, inclusión formal de `FormAutorizacionVacantes.cs` (`<SubType>Form</SubType>`) e `ItemVacanteDto.cs` en `<Compile Include=...>` para eliminar riesgo de CS0246 en entornos de compilación explícita.
- **Fase 2042:** En `AgValidadorVentana.CalcularDeficitEstructural`, depuración del selector de cohorte a `(Semestre, Grupo, Sede)` eliminando `IdAula` para no colapsar la concurrencia a aulas físicas totales (92). Agregado soporte a `InformeFactibilidad.DeficitDocentePorAsignatura`.
- **Fase 2043:** En `AgFormGeneradorHorarios.cs`, limpieza sistemática de `_vacantesAutorizadas` y `_vacantesAutorizadasClaves` en `CargarCorridaDesdeComboAsync`, `IniciarOptimizacionAsync` y `RecalcularViolaciones`, con intersección de permisos frente a las vacantes no consentidas del individuo actual para impedir fuga de autorizaciones entre corridas.
- **Fase 2044:** En `AgCompuertaPromocion.cs`, incorporación de `EsPromovibleSolicitable(...)` como evaluación preliminar formal (HARD == 0 y staging válido) para habilitar el botón en la interfaz, dejando que `EsPromovible` gobierne con rigor la compuerta post-autorización.
- **Fase 2045:** En `ItemVacanteDto.cs`, `FormAutorizacionVacantes.cs`, `AgPromotorVinculaciones.cs`, `AgCompuertaPromocion.cs` y `AgFormGeneradorHorarios.cs`, propagación y validación integral de claves `asig:sede` en el diálogo, justificación y parser de promociones. Actualización de suites de pruebas en `AgValidadorVentanaTests.cs` y `AgCompuertaPromocionTests.cs`.

### Verificación de Principios (Ronda 461)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2041 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `ConversorPDF.XLL.csproj:685-689` |
| F2042 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgValidadorVentana.cs:592-625`, `InformeFactibilidad.cs:21,29`, `AgGeneradorHorarios.cs:108`, `AgValidadorVentanaTests.cs:696-745` |
| F2043 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.cs:2427-2428, 4440-4443, 4675-4678, 4855-4860` |
| F2044 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgCompuertaPromocion.cs:338-350`, `AgFormGeneradorHorarios.cs:4445, 4680, 4744`, `AgCompuertaPromocionTests.cs:385-410` |
| F2045 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `ItemVacanteDto.cs:10-12`, `FormAutorizacionVacantes.cs:25, 41-59, 155`, `AgPromotorVinculaciones.cs:137-145, 1045-1075`, `AgCompuertaPromocion.cs:245-295, 360-375`, `AgFormGeneradorHorarios.cs:90, 4850-4870, 4975`, `AgCompuertaPromocionTests.cs:340-385` |

---

## Ronda 462 — Productor del déficit en pre-vuelo, soporte de vacantes sin sede (asig:0), botón consistente en finally de promoción y compuerta de promoción de solo lectura (Fases 2046–2049)

### Resumen de la Ronda
- **Fase 2046:** En `AgValidadorVentana.cs`, cálculo de `informe.DeficitDocentePorAsignatura` en `ValidarFactibilidad` evaluando la concurrencia máxima por franja/jornada frente a docentes habilitados del pool. Emisión de `WARN [AG-TOPE] fallback usado: {deficitTotal}` en `CalcularDeficitEstructural` cuando se recurre a la cota superior.
- **Fase 2047:** En `ItemVacanteDto.cs`, simplificación de `Clave` a `IdAsignatura.ToString()` si `IdSede == 0`. En `FormAutorizacionVacantes.cs`, enrutamiento de vacantes con sede nula a `Autorizadas`. En `AgPromotorVinculaciones.cs`, parseo de `asig:0` sin descarte hacia `asigs`.
- **Fase 2048:** En `AgFormGeneradorHorarios.cs`, sincronización de la evaluación del botón en el bloque `finally` de promoción mediante `EsPromovibleSolicitable(...)`, permitiendo reabrir el diálogo de autorización con vacantes no consentidas.
- **Fase 2049:** En `AgCompuertaPromocion.cs`, retiro de mutación sobre `contexto.ClavesVacantesConsentidas` en `ValidarPromocion`. Implementación del predicado local `ConsentidaLocal` de estricta lectura inmutable.

### Verificación de Principios (Ronda 462)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2046 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgValidadorVentana.cs:484-517, 672`, `AgValidadorVentanaTests.cs:746-785` |
| F2047 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `ItemVacanteDto.cs:12`, `FormAutorizacionVacantes.cs:159-166`, `AgPromotorVinculaciones.cs:1072-1076`, `AgCompuertaPromocionTests.cs:385-420` |
| F2048 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.cs:4975` |
| F2049 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgCompuertaPromocion.cs:259-281`, `AgCompuertaPromocionTests.cs:421-445` |

---

## Ronda 463 — Verificación de Wave 462 sin defectos bloqueantes y análisis de precisión del déficit (Fases 2050–2051)

### Resumen de la Ronda
- **Verificación Wave 462:** Sin defectos bloqueantes en compilación, runtime ni lógica en Fases 2046–2049.
- **Fase 2050 (Opcional):** Análisis de granularidad por patrón vs jornada. Se mantiene cota superior de jornada sin regresión.
- **Fase 2051 (Opcional):** Análisis de momento de cálculo de déficit en informe P07 vs consentimiento. Sin impacto funcional en la compuerta ni en el generador.

### Verificación de Principios (Ronda 463)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2050 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgValidadorVentana.cs:484-517` (opcional sin cambios de código) |
| F2051 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgValidadorVentana.cs:484-517`, `AgFormGeneradorHorarios.cs:1690-1712` (opcional sin cambios de código) |

---

## Ronda 464 — Consentimiento sobre demanda real de la corrida y separación de vacantes estructurales (R16) de liberadas (R15) (Fases 2052–2056)

### Resumen de la Ronda
- **Fase 2052:** Creación de `AgConsentimientoVacantes.cs` y `VacanteEstructuralInfo.cs`. El consentimiento se evalúa sobre la demanda real `ctx.AulaAsignaturas` y `AulasPorId`, evitando incluir materias del catálogo ajenas al programa de la corrida.
- **Fase 2053:** `AgContexto.cs` incorpora `PoolExiste` y `PoolHabilitado(idAsignatura, idPrograma, semestre)`. `AgFabricaIndividuos.cs` discrimina `SIN_CANDIDATOS` (sin registros) de `NO_CUMPLE_REQUISITOS` (docente presente pero no habilitado para ese programa/semestre). Telemetría `PoolEfectivo` en la carga del contexto.
- **Fase 2054:** Adición de `R16_VacanteEstructuralNoConsentida = 35` en `AgReglaViolacion.cs` y penalización en `AgPesosFitness.cs`. En `AgValidadorReglas.cs`, `R16` evalúa vacantes estructurales no consentidas (1 por asignatura) y `R15` evalúa vacantes liberadas frente a `MaxVacantesForzadas`.
- **Fase 2055:** Diálogo informado en `AgFormGeneradorHorarios.cs` agrupando por asignatura demandada con cohorte/sede y motivo. Persistencia dual (`int` y `"asig:sede"`) en `parametrosDict` y `dict["vacantesConsentidas"]`.
- **Fase 2056:** Registro de `AgConsentimientoVacantes.cs` y `VacanteEstructuralInfo.cs` en `ConversorPDF.XLL.csproj`. Pruebas en `AgConsentimientoVacantesTests.cs`, `AgFabricaIndividuosTests.cs` y `AgValidadorReglasTests.cs` incluyendo regresión de Corrida #110.

### Verificación de Principios (Ronda 464)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2052 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgConsentimientoVacantes.cs:1-55`, `VacanteEstructuralInfo.cs:1-17`, `AgFormGeneradorHorarios.cs:1684-1725, 2673-2715` |
| F2053 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgContexto.cs:336-360`, `AgFabricaIndividuos.cs:941-965`, `AgFormGeneradorHorarios.cs:1406-1420` |
| F2054 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgReglaViolacion.cs:41`, `AgPesosFitness.cs:67`, `AgValidadorReglas.cs:710-755` |
| F2055 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.cs:1684-1725, 2673-2715, 2925-2930, 3350-3362` |
| F2056 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `ConversorPDF.XLL.csproj:697-698`, `AgConsentimientoVacantesTests.cs:1-100`, `AgFabricaIndividuosTests.cs:493-514`, `AgValidadorReglasTests.cs:1525-1685` |

---

## Ronda 465 — Corrección de regresión centinela, alineación de semilla cohorte y granularidad sin wildcard prematuro (Fases 2057–2061)

### Resumen de la Ronda
- **Fase 2057:** En `AgFabricaIndividuos.cs`, corrección de `hayActivosReales` y `hayInactivosReales` excluyendo `IdTerceroVacante`. Un pool con solo el centinela devuelve `SIN_CANDIDATOS` (evita falsa clasificación `NO_CUMPLE_REQUISITOS`, resolviendo la regresión de F1996).
- **Fase 2058:** Creación de la sobrecarga `ClasificarMotivo(c, a, semestre, idPrograma)` en `AgFabricaIndividuos.cs` y su consumo en `AgConsentimientoVacantes.cs`, alineando la semilla de clasificación con la cohorte del aula en asignaturas MultiSemestre.
- **Fase 2059:** En `AgConsentimientoVacantes.cs`, se emite `asigs.Add(g.Key)` únicamente si todas las cohortes demandadas de la asignatura carecen de docente habilitado. Si solo una cohorte carece de docente, se restringe al par granular `claves.Add($"{asig}:{sede}")`.
- **Fase 2060:** En `AgValidadorReglas.cs`, incorporación de `OrigenVacante.NoDocente` para motivos `SIN_PLAN_OLEADA` y `SIN_BLOQUE_NIVEL`, excluyéndolos de los contadores `R15` (liberadas) y `R16` (estructurales docentes).
- **Fase 2061:** Seguimiento git mediante `git add` de los 3 archivos nuevos (`AgConsentimientoVacantes.cs`, `VacanteEstructuralInfo.cs`, `AgConsentimientoVacantesTests.cs`), sincronización en `REGISTRY.md` (fila R-139, §2 `ronda=465`, `fase=2061`) y pruebas de validación añadidas en las 3 suites de tests.

### Verificación de Principios (Ronda 465)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2057 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFabricaIndividuos.cs:940-970`, `AgFabricaIndividuosTests.cs:516-555` |
| F2058 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFabricaIndividuos.cs:342-344, 942-972`, `AgConsentimientoVacantes.cs:55-75`, `AgFabricaIndividuosTests.cs:556-580`, `AgConsentimientoVacantesTests.cs:145-180` |
| F2059 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgConsentimientoVacantes.cs:25-85`, `AgConsentimientoVacantesTests.cs:111-144` |
| F2060 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgValidadorReglas.cs:715-785`, `AgValidadorReglasTests.cs:1690-1775` |
| F2061 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `git status (archivos seguidos en index)`, `REGISTRY.md:27-28, 176` |

---

## Ronda 466 — Aislamiento de motivos no docentes en fallos de patrón, filtro de aulas generables y reutilización de consentimiento informado (Fases 2062–2065)

### Resumen de la Ronda
- **Fase 2062:** En `AgFabricaIndividuos.cs:715`, asignación de `AgMotivoVacante.SIN_PLAN_OLEADA` cuando `resPatron.IsFailure`, eliminando la falsa clasificación como `SIN_DISPONIBILIDAD` que incrementaba `liberadas` en `R15`.
- **Fase 2063:** En `AgConsentimientoVacantes.cs:32`, adición del filtro `ctx.ManifiestoCobertura == null || ctx.PuedeGenerarAula(aa.IdAula)`, descartando cohortes en aulas no disponibles o bloqueadas.
- **Fase 2064:** En `AgFormGeneradorHorarios.cs:2675-2715`, verificación de idempotencia en Optimización mediante `SetEquals`; se reutiliza el consentimiento aprobado en pre-vuelo P07 evitando el doble diálogo modal idéntico.
- **Fase 2065:** Trazabilidad formal en `REGISTRY.md` (fila R-140, §2 `ronda=466`, `fase=2065`) y validación mediante tests unitarios añadidos en las suites correspondientes.

### Verificación de Principios (Ronda 466)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2062 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFabricaIndividuos.cs:715`, `AgFabricaIndividuosTests.cs:583-605`, `AgValidadorReglasTests.cs:1780-1810` |
| F2063 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgConsentimientoVacantes.cs:32`, `AgConsentimientoVacantesTests.cs:186-215` |
| F2064 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.cs:2675-2715` |
| F2065 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `REGISTRY.md:27-28, 177`, suites de pruebas en disco |

---

## Ronda 467 — Consentimiento informado idempotente y unificado, erradicación de doble diálogo P07/optimización y cobertura de pruebas (Fases 2066–2068)

### Resumen de la Ronda
- **Fase 2066:** Extracción de método unificado e idempotente `AsegurarConsentimientoVacantes(bool interactivo)` en `AgFormGeneradorHorarios.cs:1446-1496`, consumido en `EjecutarDiagnosticoPreVueloAsync` e `IniciarOptimizacionAsync`.
- **Fase 2067:** Cobertura de pruebas unitarias en `AgConsentimientoVacantesTests.cs`: inclusión con manifiesto nulo, exclusión con manifiesto construido, idempotencia de `SetEquals` con mismo universo y disparador ante divergencia de claves.
- **Fase 2068:** Cierre formal en `REGISTRY.md` (fila R-141, §2 `ronda=467`, `fase=2068`), sincronización de bitácoras vivas e indexado en control de versiones.

### Verificación de Principios (Ronda 467)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2066 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.cs:1446-1496, 1741, 2700` |
| F2067 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgConsentimientoVacantesTests.cs:186-327` |
| F2068 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `REGISTRY.md:27-28, 178`, `implementation_plan.md`, `walkthrough.md` |

---

## Ronda 468 — Erradicación definitiva del doble diálogo de consentimiento, unificación de logs y helper puro EsVigente (Fases 2069–2072)

### Resumen de la Ronda
- **Fase 2069:** Eliminación de la llamada a `AsegurarConsentimientoVacantes` en `EjecutarDiagnosticoPreVueloAsync` (P07). El flujo garantiza que la consulta modal ocurra en un único punto canónico en `IniciarOptimizacionAsync` tras construir el manifiesto de cobertura.
- **Fase 2070:** Normalización de logs con la etiqueta `[CONSENTIMIENTO]` en `AgFormGeneradorHorarios.cs`.
- **Fase 2071:** Implementación de `AgConsentimientoVacantes.EsVigente` y pruebas unitarias dedicadas en `AgConsentimientoVacantesTests.cs`.
- **Fase 2072:** Registro en `REGISTRY.md` (fila R-142, §2 `ronda=468`, `fase=2072`), verificación de rotación e indexación en git stage.

### Verificación de Principios (Ronda 468)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2069 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.cs:1730-1745, 2690-2705` |
| F2070 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.cs:1446-1496` |
| F2071 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgConsentimientoVacantes.cs:86-95`, `AgConsentimientoVacantesTests.cs:282-340` |
| F2072 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `REGISTRY.md:27-28, 179`, `implementation_plan.md`, `walkthrough.md` |

---

## Ronda 469 — Precondición de manifiesto antes de consentimiento en corrida y cierre de ciclo (Fases 2073–2074)

### Resumen de la Ronda
- **Fase 2073:** Reordenamiento en `AgFormGeneradorHorarios.cs:2690-2705`: precondición de cobertura evaluada antes de invocar `AsegurarConsentimientoVacantes`.
- **Fase 2074:** Cierre y gate de las olas 464–469, registro en `REGISTRY.md` (fila R-143, §2 `ronda=469`, `fase=2074`), compilación y empaquetado release v1.6.4.31.

### Verificación de Principios (Ronda 469)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2073 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.cs:2690-2705` |
| F2074 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `REGISTRY.md:27-28, 180`, `implementation_plan.md`, `walkthrough.md` |

---

## Ronda 470 — Corrección de advertencia Fecha Límite Excedida en Sesiones Académicas según ventana canónica de período (Fases 2075–2078)

### Resumen de la Ronda
- **Fase 2075:** Creación de `ICierrePeriodoResolver` y `CierrePeriodoResolver` en `ConversorPDF.XLL.Academic.Application` junto al DTO `ResultadoCierrePeriodo`. Implementada la precedencia canónica: ventana temporal del período (`TryObtenerVentana`) como fuente primaria y fallback al 15/12 del año resuelto (nombre YYYY-MM, PeriodoFin o hoy.Year) con constantes `DiaCierreSemestre = 15` y `MesCierreSemestre = 12`.
- **Fase 2076:** En `FormSesionesAcademicas.cs`, cableado del campo e inicializador `_cierrePeriodoResolver` y reemplazo de la regla fija de DateTime.Now.Year/15-12 por `_cierrePeriodoResolver.Evaluar(periodoSeleccionado, dtpFinal.Value.Date, DateTime.Today)`, preservando el aviso confirmable y emitiendo traza diagnóstica con el límite y origen evaluados.
- **Fase 2077:** En `CatalogCacheService.cs`, enriquecimiento de la query de `periodos` con `PeriodoInicio` y `PeriodoFin` y fallback `pragma_table_info`. Añadida la invalidación explícita `CatalogCacheService.Invalidate()` en `FormPeriodos.cs` (guardar/eliminar) y `AgFormGeneradorHorarios.cs` (extensión de ventana). Registro de tipos en `ConversorPDF.XLL.csproj`.
- **Fase 2078:** Implementación de la suite de pruebas unitarias `CierrePeriodoResolverTests.cs` en `ConversorPDF.Tests` con xUnit + FluentAssertions, cubriendo la regresión del defecto y todos los escenarios de ventana, fallback y formato de mensaje.

### Verificación de Principios (Ronda 470)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2075 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `ICierrePeriodoResolver.cs`, `CierrePeriodoResolver.cs`, `ResultadoCierrePeriodo.cs` |
| F2076 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `FormSesionesAcademicas.cs:77, 124, 683-692` |
| F2077 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `CatalogCacheService.cs:44-59, 107-120`, `FormPeriodos.cs:189, 226`, `AgFormGeneradorHorarios.cs:792`, `ConversorPDF.XLL.csproj:277-279` |
| F2078 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `CierrePeriodoResolverTests.cs:1-190` |

---

## Ronda 471 — Corrección de mojibake en servicio nuevo y coherencia de caché de períodos (Fases 2079–2082)

### Resumen de la Ronda
- **Fase 2079:** Reescritura limpia en UTF-8 con BOM de `CierrePeriodoResolver.cs`, `ICierrePeriodoResolver.cs`, `ResultadoCierrePeriodo.cs` y `CierrePeriodoResolverTests.cs`, eliminando doble codificación y preservando acentos/caracteres en diálogos y documentación.
- **Fase 2080:** Implementación de `CatalogCacheService.ActualizarPeriodo` y `EliminarPeriodo` para refresco no destructivo in-place del catálogo en memoria, enlazándolo en `FormPeriodos.cs` y `AgFormGeneradorHorarios.cs`.
- **Fase 2081:** Verificación de rango 1..9999 para año en `CierrePeriodoResolver.ResolverAnioFallback`, previniendo excepción con año cero e incorporando test de regresión en `CierrePeriodoResolverTests.cs`.
- **Fase 2082:** Eliminación de la variable no utilizada `fechaLimite` en `FormSesionesAcademicas.cs:681`.

### Verificación de Principios (Ronda 471)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2079 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `CierrePeriodoResolver.cs`, `ICierrePeriodoResolver.cs`, `ResultadoCierrePeriodo.cs`, `CierrePeriodoResolverTests.cs` |
| F2080 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `CatalogCacheService.cs:94-135`, `FormPeriodos.cs:189-198, 233`, `AgFormGeneradorHorarios.cs:792` |
| F2081 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `CierrePeriodoResolver.cs:84-88, 97-100`, `CierrePeriodoResolverTests.cs:190-205` |
| F2082 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `FormSesionesAcademicas.cs:680-682` |

---

## Ronda 472 — Diagnóstico forense del bloqueo P08 y trazabilidad DEBUG del consentimiento, tope estructural y vacantes (Fases 2083–2089)

### Resumen de la Ronda
- **Fase 2083:** Implementación de `AgConsentimientoVacantes.FormatearDetalle` y `OrdenSemestre` con ordenación romana canónica soportada en `AgFormatoAsignatura.SemestreArabigo`. Integración en `AgFormGeneradorHorarios.cs` con el resolver de etiquetas `ResolverEtiquetaSedeConsentimiento` y cobertura con 3 pruebas unitarias en `AgConsentimientoVacantesTests.cs`.
- **Fase 2084:** Trazabilidad DEBUG `[AG-CONSENT]` en `AgConsentimientoVacantes.Calcular` documentando motivos y decisiones por materia, cohortes y alertas de `IdSede=0`, junto con registro de entrada y salida del diálogo en `AgFormGeneradorHorarios.AsegurarConsentimientoVacantes`.
- **Fase 2085:** Instrumentación DEBUG `[AG-TOPE]` en la fase de pre-vuelo de factibilidad (`AgValidadorVentana.cs:485-516`), en el fallback de `CalcularDeficitEstructural` y en el punto de fijación del tope en `AgGeneradorHorarios.cs:106-109`.
- **Fase 2086:** Diagnóstico forense del campeón infactible en `AgGeneradorHorarios.cs` con etiquetas `[AG-HARD]` (reglas duras no resueltas), `[AG-VAC]` (balance de vacantes consentidas vs liberadas), `[AG-SNAPSHOT]` (trazas de genes causantes o liberados acotadas a 50) y `[AG-PARADA]` (motivo de fin, estancamiento e historial de convergencia).
- **Fase 2087:** Instrumentación `[AG-VACACION]` en `AgAsignadorTercero.DiagnosticarMotivoVacante`, `AgCoherenciaGrupo.Sellar` y rama de liberación de `AgOperadorMutacion.cs` para explicar la causa raíz de las asignaciones pasadas a `SIN_DISPONIBILIDAD`.
- **Fase 2088:** Enriquecimiento de `AgDiagnosticoInfactibilidad.Describir` con `ReglasDuras=[...]` y distinción de vacantes consentidas vs liberadas; adición de `Reglas duras: [...]` en `ResumenCorto`; y emisión de `[AG-GATE]` en `AgFormGeneradorHorarios.cs:3266` antes del modal. Cobertura con 2 tests en `AgDiagnosticoInfactibilidadTests.cs`.
- **Fase 2089:** Compilación limpia de `ConversorPDF.XLL` (0 errores) y suite completa de pruebas unitarias (`ConversorPDF.Tests`) con 1801 pruebas superadas (0 fallas). Cierre documental y sincronización de registro.

### Verificación de Principios (Ronda 472)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2083 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgConsentimientoVacantes.cs:140-185`, `AgFormGeneradorHorarios.cs:1467-1510`, `AgConsentimientoVacantesTests.cs:340-395` |
| F2084 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgConsentimientoVacantes.cs:65-135`, `AgFormGeneradorHorarios.cs:1455-1490` |
| F2085 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgValidadorVentana.cs:510-525, 680-695`, `AgGeneradorHorarios.cs:106-118` |
| F2086 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgGeneradorHorarios.cs:380-515` |
| F2087 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgAsignadorTercero.cs:200-245`, `AgCoherenciaGrupo.cs:495-508`, `AgOperadorMutacion.cs:112-128` |
| F2088 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgDiagnosticoInfactibilidad.cs:155-165, 290-300, 330-340`, `AgFormGeneradorHorarios.cs:3260-3270`, `AgDiagnosticoInfactibilidadTests.cs:165-205` |
| F2089 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `REGISTRY.md`, `implementation_plan.md`, `walkthrough.md` |

---

## Ronda 473 — Diagnóstico forense con trazabilidad correcta, pero causa raíz de R15 sin corregir (unidades y consentimiento sede 0) (Fases 2090–2095)

### Resumen de la Ronda
- **Fase 2090:** Corrección de la unidad de conteo en R15 (`AgValidadorReglas.cs:725-745`) de sesiones a cohortes (`cohortesLiberadas.Count`) evitando falsos positivos HARD=1 cuando 1 cohorte sin docente libera múltiples sesiones contra un déficit 0. Centralización del cálculo de déficit estructural en `AgValidadorVentana.CalcularDeficitPorAsignatura` consumido de forma idéntica por pre-vuelo y fallback.
- **Fase 2091:** Prevención de disparo falso e irreparable de R16 cuando `IdSede <= 0`: registro de token compuesto `{asig}:0` en `AgConsentimientoVacantes.Calcular` y normalización con `Math.Max(0, idSede)` en `AgContexto.EstaConsentida`.
- **Fase 2092:** Creación de `AgOrigenVacante` (Domain) y `AgClasificadorVacante` (Application) unificando la taxonomía de vacantes (NoDocente, EstructuralDocente, LiberadaAg) en todo el ciclo: validación R15/R16, telemetría `[AG-VAC]` y diagnóstico de infactibilidad.
- **Fase 2093:** Cálculo modal riguroso del motivo mostrado en el modal de consentimiento (`AgConsentimientoVacantes.FormatearDetalle`) mediante agrupación por frecuencia (`GroupBy` + `OrderByDescending`), sustituyendo el primer motivo arbitrario (`FirstOrDefault()`).
- **Fase 2094:** Limpieza de resolución de sedes en `AgFormGeneradorHorarios.ResolverEtiquetaSedeConsentimiento` suprimiendo el fallback ambiguo a subunidad de contexto y garantizando orden determinista de sedes por id en el detalle del modal.
- **Fase 2095:** Pruebas unitarias de las 5 fases en `ConversorPDF.Tests` y sincronización documental de cierre en `REGISTRY.md` (R-147).

### Verificación de Principios (Ronda 473)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2090 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgValidadorReglas.cs:725-745`, `AgValidadorVentana.cs:505-520, 640-660, 685-715` |
| F2091 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgConsentimientoVacantes.cs:65-75`, `AgContexto.cs:150-160` |
| F2092 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgOrigenVacante.cs`, `AgClasificadorVacante.cs`, `AgValidadorReglas.cs:760-780`, `AgGeneradorHorarios.cs:425-455`, `AgDiagnosticoInfactibilidad.cs:290-305` |
| F2093 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgConsentimientoVacantes.cs:175-185` |
| F2094 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.cs:1495-1510`, `AgConsentimientoVacantes.cs:165-175` |
| F2095 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgValidadorReglasTests.cs:1810-1850`, `AgConsentimientoVacantesTests.cs:395-440`, `AgValidadorVentanaTests.cs:780, 830-865`, `AgMotivoVacanteTests.cs:15-40`, `REGISTRY.md` |

---

## Ronda 474 — R15 en cohortes aplicado, pero el déficit perdió la concurrencia por jornada (regresión F2031/F2036) y persisten fugas de clasificación (Fases 2096–2102)

### Resumen de la Ronda
- **Fase 2096:** Preservación de la concurrencia por jornada en el cálculo de déficit estructural de `AgValidadorVentana.CalcularDeficitPorAsignatura` (`maxConcurrentes` por jornada menos pool de docentes activos), restituyendo el comportamiento sin inflar vacantes permitidas.
- **Fase 2097:** Implementación de `AgClasificadorVacante.ClaveCohorte(AgGen g)` incorporando `aula:{id}|gen:{id}` cuando no existen `Semestre` ni `Grupo`, previniendo subconteo en R15 (`AgValidadorReglas.cs`) y telemetría `[AG-VAC]`.
- **Fase 2098:** Incorporación de `AgOrigenVacante.Indeterminado` para motivos vacíos o desconocidos, con exclusión preventiva en R15/R16. Cobertura en `AgMotivoVacanteTests.cs` y `AgValidadorReglasTests.cs`.
- **Fase 2099:** Completitud de la partición de vacantes en `AgDiagnosticoInfactibilidad.Describir` desglosando `ESTRUCTURALES_NC`, `NO_DOCENTE` e `INDETERMINADAS`, verificando que la suma coincida exactamente con `Vacantes`.
- **Fase 2100:** Eliminación del costo O(G×A) en factibilidad y déficit con `AgValidadorVentana.ConstruirDeficitBase`, evaluando todo el catálogo de aulas/asignaturas en un único recorrido O(N).
- **Fase 2101:** Salvaguarda en `AgConsentimientoVacantes.FormatearDetalle` para evitar etiquetas de sede vacías ante resolvers que retornan cadenas vacías o espacios. Renombrado de prueba en `AgValidadorVentanaTests.cs`.
- **Fase 2102:** Pruebas unitarias de regresión en `ConversorPDF.Tests` y sincronización documental de cierre en `REGISTRY.md` (R-148).

### Verificación de Principios (Ronda 474)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2096 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgValidadorVentana.cs:620-645`, `AgValidadorVentanaTests.cs:745-805` |
| F2097 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgClasificadorVacante.cs:15-25`, `AgValidadorReglas.cs:730-740`, `AgGeneradorHorarios.cs:425-435`, `AgMotivoVacanteTests.cs:180-210` |
| F2098 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgOrigenVacante.cs:10-15`, `AgClasificadorVacante.cs:30-38`, `AgValidadorReglas.cs:720-726`, `AgMotivoVacanteTests.cs:175-185`, `AgValidadorReglasTests.cs:1865-1885` |
| F2099 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgDiagnosticoInfactibilidad.cs:295-305`, `AgDiagnosticoInfactibilidadTests.cs:205-230` |
| F2100 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgValidadorVentana.cs:485-505, 645-685, 700-725` |
| F2101 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgConsentimientoVacantes.cs:165-175`, `AgValidadorVentanaTests.cs:745`, `AgConsentimientoVacantesTests.cs:445-460` |
| F2102 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `REGISTRY.md`, `implementation_plan.md`, `walkthrough.md` |

---

## Ronda 475 — Error de compilación en la suite y R15 aún compara agregados distintos (max por jornada vs total de cohortes) (Fases 2103–2107)

### Resumen de la Ronda
- **Fase 2103:** Corrección de error de compilación `CS1022`/`CS1513` por llave de cierre duplicada en `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgValidadorReglasTests.cs`.
- **Fase 2104:** Ajuste de R15 en `AgValidadorReglas.cs` para sumar la concurrencia máxima por jornada de cada asignatura (`liberadasPorAsignatura.Values.Sum(...)`), igualando la unidad de medida del déficit estructural. Cobertura en `AgValidadorReglasTests.cs`.
- **Fase 2105:** Lazy allocation de `sedesPorAsig` en `AgValidadorVentana.cs` para evitar O(N) fuera de modo verboso y normalización estricta `Trim().ToUpperInvariant()` de `Jornada` para neutralizar diferencias de case.
- **Fase 2106:** Clarificación de trazas en `AgGeneradorHorarios.cs` (`[AG-VAC]`) distinguiendo `estructuralesAsignaturas` de `estructuralesSesiones`.
- **Fase 2107:** Validación mediante build y ejecución exitosa de los 1824 tests en `ConversorPDF.Tests`. Sincronización en `REGISTRY.md` (R-149).

### Verificación de Principios (Ronda 475)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2103 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgValidadorReglasTests.cs:1880-1887` |
| F2104 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgValidadorReglas.cs:715-765`, `AgValidadorReglasTests.cs:1890-1915` |
| F2105 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgValidadorVentana.cs:488-510, 670-675` |
| F2106 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgGeneradorHorarios.cs:410-440` |
| F2107 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `REGISTRY.md`, `implementation_plan.md`, `walkthrough.md`, `ConversorPDF.Tests` (1824 tests PASS) |

---

## Ronda 476 — La infactibilidad de capacidad se publica como `Error` estructural y veta escalón/reclasificación (Fases 2108–2113)

### Resumen de la Ronda
- **Fase 2108:** Se instrumentó `AgEscaladorRanuraAula.cs` registrando `evalPaso` y los parámetros diagnósticos de omisión en `ESCALADOR_OMITIR` (`pasoCap`, `pasoPend`, `pasoError`, `antesError`, `candidatos`), erradicando el valor de reserva engañoso de `capDespues`.
- **Fase 2109:** En `AgPlanificadorOleadas.cs` y `AgEvaluadorDemandaVentana.cs`, se delimitó `EsInfactibilidadDeCapacidad` para tratar `I5_SIN_CAPACIDAD` e `I5_SIN_SEMANAS_HABILES` como resultados evaluables (no errores duros), evitando el bloqueo ciego de `PasoMejora` y permitiendo evaluar escalones y candidatas de redistribución.
- **Fase 2110:** En `AgValidadorVentana.cs`, se vinculó el consejo de escalado a la reducción verificada de `DemandaPendiente`, advirtiendo explícitamente cuándo los patrones rígidos impiden aprovechar los días incorporados (Lunes).
- **Fase 2111:** Se añadieron pruebas de regresión en `AgEscaladoTresCasosMatrizTests.cs` y `AgValidadorVentanaTests.cs` garantizando que aulas saturadas (64 sesiones vs 61 ranuras) propongan el escalón correspondiente y no se silencien.
- **Fase 2113:** En `AgFormGeneradorHorarios.cs`, se implementó validación de factibilidad fail-fast previa a la inserción en base de datos (`InsertCorridaAsync`), impidiendo la persistencia de corridas fallidas estériles.
- **Fase 2112:** Gate y trazabilidad de cierre en `REGISTRY.md` para la asignación R-150.

### Verificación de Principios (Ronda 476)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2108 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEscaladorRanuraAula.cs:515-565` |
| F2109 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorOleadas.cs:200-220`, `AgEvaluadorDemandaVentana.cs:73-86, 122-135` |
| F2110 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgValidadorVentana.cs:313-345` |
| F2111 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEscaladoTresCasosMatrizTests.cs`, `AgValidadorVentanaTests.cs` |
| F2113 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.cs:2708-2735` |
| F2112 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `REGISTRY.md`, `implementation_plan.md`, `walkthrough.md` |

---

## Ronda 477 — La ola 476 se cerró como implementada sin cambios en disco y con un cambio huérfano no declarado (Fases 2114–2115)

### Resumen de la Ronda
- **Fase 2114:** Se reabrió la ola y se materializó el código de las Fases 2108–2113 en disco en los 7 archivos de código y pruebas, confirmando marcadores con grep.
- **Fase 2115:** Se normalizó EOL y se atribuyó formalmente el título en `SemestresExcelExporter.cs:252`. Se asignó R-151 y se actualizó `REGISTRY.md`.

### Verificación de Principios (Ronda 477)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2114 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorOleadas.cs`, `AgEvaluadorDemandaVentana.cs`, `AgEscaladorRanuraAula.cs`, `AgValidadorVentana.cs`, `AgFormGeneradorHorarios.cs`, tests |
| F2115 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `SemestresExcelExporter.cs:252`, `REGISTRY.md` |

---

## Ronda 478 — El fail-fast anula el consentimiento explícito y el consejo de ventana sigue prometiendo una solución inexistente (Fases 2116–2119)

### Resumen de la Ronda
- **Fase 2116:** Se respetó el override explícito del usuario (`continuarPeseAInfactibilidad`) en `AgFormGeneradorHorarios.cs`, ejecutando el gate de factibilidad previa solo si no hubo consentimiento explícito de continuar y reubicándolo tras el guard de manifiesto.
- **Fase 2117:** Se refinó el consejo de ventana en `AgValidadorVentana.cs` declarando el pendiente residual cuando el escalón mejora pero no elimina el déficit por patrones rígidos.
- **Fase 2118:** Se evitó la invocación de `AgEvaluadorDemandaVentana.Evaluar` en `AgValidadorVentana.cs` cuando `capEscalon <= capacidadTotalRanuras`, eliminando la sobrecarga computacional y proyecciones a 2 años.
- **Fase 2119:** Se declararon cotas conservadoras sin plan en `AgEvaluadorDemandaVentana.cs` para infactibilidades de capacidad, registrando el contexto en el log de advertencia.

### Verificación de Principios (Ronda 478)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2116 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.cs:2729-2750` |
| F2117 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgValidadorVentana.cs:338-355`, `AgValidadorVentanaTests.cs:930-975` |
| F2118 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgValidadorVentana.cs:328-335` |
| F2119 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEvaluadorDemandaVentana.cs:139-155` |

---

## Ronda 479 — El fail-fast sigue sin salida cuando el escalador no propone, y el diagnóstico atribuye una causa falsa al escalón sin ganancia (Fases 2120–2123)

### Resumen de la Ronda
- **Fase 2120:** Se introdujo opción interactiva Yes/No en `AgFormGeneradorHorarios.cs` cuando el escalador retorna `SinCambios`, restituyendo la capacidad de continuar voluntariamente con sesiones pendientes.
- **Fase 2121:** En `AgValidadorVentana.cs`, se distinguió si la falta de ganancia del escalón proviene de ausencia de bloques o de festivos que neutralizan los días añadidos.
- **Fase 2122:** Se rotuló "En riesgo" en `FormAutorizarReclasificacion.cs` para evitar presentar demandas completas como déficits netos, precisándolo en `AgEvaluadorDemandaVentana.cs`.
- **Fase 2123:** Se incluyeron pruebas de regresión en `AgValidadorVentanaTests.cs` para el mensaje de reducción residual y la lógica pura de override.

### Verificación de Principios (Ronda 479)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2120 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.cs:2740-2765` (`contInf`) |
| F2121 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgValidadorVentana.cs:329-342` (`ranurasDia <= ranurasBase`) |
| F2122 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `FormAutorizarReclasificacion.cs:143-145`, `AgEvaluadorDemandaVentana.cs:141-145` |
| F2123 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgValidadorVentanaTests.cs:995-1070` |

---

## Ronda 480 — El test del override no compila (enum privado) y el test residual no alcanza la rama que dice cubrir (Fases 2124–2126)

### Resumen de la Ronda
- **Fase 2124:** Se extrajo `AgResultadoEscaladoUI` a `ConversorPDF.XLL/Academic/Domain/Enums/AgResultadoEscaladoUI.cs` y se registró en `ConversorPDF.XLL.csproj`, eliminando la declaración privada en `AgFormGeneradorHorarios.cs`.
- **Fase 2125:** Se calibró el fixture de `AgValidadorVentanaTests.cs` (69 sesiones con materias rígidas y una automática) y se sustituyó el assert por comprobaciones exactas de la rama residual.
- **Fase 2126:** Se computaron `diasNuevos` y `ocurrenciasNuevas` en `AgValidadorVentana.cs`, discriminando causas sin inferencias espurias.

### Verificación de Principios (Ronda 480)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2124 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgResultadoEscaladoUI.cs`, `AgFormGeneradorHorarios.cs:2145-2165`, `ConversorPDF.XLL.csproj:782` |
| F2125 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgValidadorVentanaTests.cs:998-1065` |
| F2126 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgValidadorVentana.cs:329-348` |

---

## Ronda 481 — La teoría del override es tautológica y el mensaje del escalón cuenta ranuras como días (Fases 2127–2128)

### Resumen de la Ronda
- **Fase 2127:** Se creó el predicado canónico `AgPreVueloEscalado.ContinuarPeseAInfactibilidad` en `ConversorPDF.XLL/Academic/Domain/Enums/AgPreVueloEscalado.cs`, registrado en `ConversorPDF.XLL.csproj`, aplicado en `AgFormGeneradorHorarios.cs` y testeado con `AgValidadorVentanaTests.cs`, eliminando la aserción tautológica sin ejecución de código.
- **Fase 2128:** Se computaron los días añadidos sobre días semanales distintos (`ranurasNuevas.Select(r => r.Dia).Distinct().Count()`) en `AgValidadorVentana.cs` y se suprimió la rama muerta de `ocurrenciasNuevas > 0` cuando `capEscalon <= capacidadTotalRanuras`.

### Verificación de Principios (Ronda 481)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2127 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPreVueloEscalado.cs`, `AgFormGeneradorHorarios.cs:2715-2725`, `AgValidadorVentanaTests.cs:1065-1073`, `ConversorPDF.XLL.csproj:783` |
| F2128 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgValidadorVentana.cs:330-342` |

---

## Ronda 482 — Hotfix del campeón con Forward Shifting por carril para eliminación de solapes R6/R13 (Fases 2129–2135)

### Resumen de la Ronda
- **Fase 2129:** Se extrajo `ConstruirSemanasValidas` y se implementó `RealinearOleadas` con la clase anidada `ResultadoRealineacion` en `AgPlanificadorOleadas.cs`, efectuando un forward shifting determinista en dos fases (calcular -> validar -> aplicar) agrupado por cohorte, aula y patrón semanal, preservando horas totales e invariantes I1/I2.
- **Fase 2130:** Se insertó el hook post-evolución en `AgGeneradorHorarios.cs` con clon profundo previo, sellado y evaluación, con reversión al original si no hay mejora neta en violaciones HARD (`hardPost >= hardPre`), log estructurado `[AG-REPARACION]` y flag estático de activación `RealineacionOleadasActiva = true`.
- **Fase 2131:** Se crearon 6 pruebas unitarias en `AgPlanificadorOleadasTests.cs` cubriendo desplazamiento por solape, rechazo seguro por ventana insuficiente, neutralidad sin solapes, recálculo de `k` con festivos intermedios, independencia entre carriles distintos e idempotencia estricta.
- **Fase 2132:** Se añadieron pruebas de regresión en `AgValidadorReglasOleadasTests.cs` verificando que un individuo campeón con múltiples solapes en carril queda completamente limpio de `R6_SecuenciaCurricular` y `R13_SolapeOleadas`, y que el caso infactible mantiene el HARD sin degradación.
- **Fase 2133:** Se incorporó prueba de integración end-to-end en `AgGeneradorHorariosTests.cs` asegurando convergencia con `HardViolations == 0` y fechas secuenciales entregadas en el mejor individuo del resultado.
- **Fase 2134:** Se añadieron pruebas de reproducibilidad y control por flag en `AgGeneradorHorariosReplayTests.cs` garantizando paridad exacta de hash estructural bajo la misma semilla y conmutación de estado del flag estático.
- **Fase 2135:** Se implementó prueba de coherencia de grupo en `AgCoherenciaGrupoTests.cs` comprobando que `Sellar` post-reparación mantiene `FechasGen` dentro del rango y concordancia con el día de semana sin festivos ni duplicados.

### Verificación de Principios (Ronda 482)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2129 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorOleadas.cs:281-283, 655-1070` |
| F2130 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgGeneradorHorarios.cs:28-34, 315-365` |
| F2131 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorOleadasTests.cs:740-955` |
| F2132 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgValidadorReglasOleadasTests.cs:340-410` |
| F2133 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgGeneradorHorariosTests.cs:110-180` |
| F2134 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgGeneradorHorariosReplayTests.cs:148-200` |
| F2135 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgCoherenciaGrupoTests.cs:575-635` |

---

## Ronda 483 — Corrección de compilación en pruebas, paridad de clave carril, preservación I1/I2 y franjas, robustez de cursor y nota documental (Fases 2136–2149)

### Resumen de la Ronda
- **Fase 2136:** Se corrigieron los nombres de enum en `AgValidadorReglasOleadasTests.cs:366-380` utilizando `AgReglaViolacion.R6_SecuenciaAsignaturas` y `AgReglaViolacion.R13_OleadasDispersas`.
- **Fase 2137:** Se normalizó la configuración de festivos en `AgPlanificadorOleadasTests.cs:853-858` y `AgCoherenciaGrupoTests.cs:585-595` mediante `ctx.Festivos.Add(...)` y `ctx.ConstruirIndices()`.
- **Fase 2138:** Se preservó la discriminación por `(Dia, Franja)` en `AgPlanificadorOleadas.cs:940-970`, eliminando el colapso con `Distinct()` ciego cuando existen múltiples ranuras o bloques en el mismo día.
- **Fase 2139:** Se suprimió el ajuste residual de horas sueltas que rompía I2 en `AgPlanificadorOleadas.cs:980-1010`, implementando validación estricta de `FechasGen.Count * Unidad == HorasGen` por gen.
- **Fase 2140:** Se unificó la clave de carril en `AgPlanificadorOleadas.cs:698-705` y `AgValidadorReglas.cs:332-336, 427-431` mediante `ClaveCarril(g)` y agrupación por `AgCalendarioConflictos.Cohorte(g)`.
- **Fase 2141:** Se resolvieron ranuras desde los genes (`DiaSemana`, `IdBloqueHorario` / `Franja`) con fallback canónico en `AgPlanificadorOleadas.cs:815-840`.
- **Fase 2142:** Se registró fielmente `hardPost` antes de la restauración del clon en el log de rollback de `AgGeneradorHorarios.cs:342-348`.
- **Fase 2143:** Se restringió `RealineacionOleadasActiva` a `internal` en `AgGeneradorHorarios.cs:32` y se implementó prueba conductual en `AgGeneradorHorariosReplayTests.cs:181-215`.
- **Fase 2144:** Se pobló `contexto.AulaAsignaturas` con llamada a `ConstruirIndices()` y se afirmó la secuencia sin condiciones en `AgGeneradorHorariosTests.cs:146-180`.
- **Fase 2145:** Se sembraron fechas en la oleada inicial y se validaron invariantes para genes reparados en `AgCoherenciaGrupoTests.cs:600-630`.
- **Fase 2146:** Se revirtió `AgValidadorVentanaTests.cs` a HEAD restaurando el escenario original.
- **Fase 2147:** Se poblaron `ctx.Aulas` y `ctx.Bloques` en `CrearContextoBase()` de `AgValidadorReglasOleadasTests.cs:36-47`.
- **Fase 2148:** Se generó la nota documental `docs/articles/changes/2026-11-19-realineacion-campeon-r1.md` y se ajustó la trazabilidad en `REGISTRY.md`.
- **Fase 2149:** Se implementó `TieneDiaUtilPosterior` en `AgPlanificadorOleadas.cs:880-915` para descartar falsos positivos de `NoCabe` cuando el slice contiene días hábiles tras `prevFin`.

### Verificación de Principios (Ronda 483)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2136 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgValidadorReglasOleadasTests.cs:366-380` |
| F2137 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorOleadasTests.cs:853-858`, `AgCoherenciaGrupoTests.cs:585-595` |
| F2138 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorOleadas.cs:940-970` |
| F2139 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorOleadas.cs:980-1010` |
| F2140 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorOleadas.cs:698-705`, `AgValidadorReglas.cs:332-336, 427-431` |
| F2141 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorOleadas.cs:815-840` |
| F2142 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgGeneradorHorarios.cs:342-348` |
| F2143 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgGeneradorHorarios.cs:32`, `AgGeneradorHorariosReplayTests.cs:181-215` |
| F2144 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgGeneradorHorariosTests.cs:146-180` |
| F2145 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgCoherenciaGrupoTests.cs:600-630` |
| F2146 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgValidadorVentanaTests.cs` (HEAD restaurado) |
| F2147 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgValidadorReglasOleadasTests.cs:36-47` |
| F2148 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `2026-11-19-realineacion-campeon-r1.md`, `REGISTRY.md:193-194` |
| F2149 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorOleadas.cs:880-915` |

---

## Ronda 484 — Balance de llaves en AgPlanificadorOleadas, validación de franjas y fechas duplicadas, paridad de clave carril en AgCarriles, aislamiento paralelo xUnit y trazabilidad documental (Fases 2150–2158)

### Resumen de la Ronda
- **Fase 2150:** Se eliminó la llave de cierre redundante en `AgPlanificadorOleadas.cs` tras el método `TieneDiaUtilPosterior`, restaurando la paridad estructural de llaves en la clase.
- **Fase 2151:** Se deduplicaron las fechas en la asignación por gen en `AgPlanificadorOleadas.cs`, incorporando la guarda `FranjaAmbigua` cuando el gen no resuelve franja y existen múltiples sesiones en el mismo día, junto con validación de unicidad de fechas.
- **Fase 2152:** Se normalizaron los días efectivos en `AsignaturaCarril.Dias` mediante `ranuras.Select(r => r.Dia).Distinct().OrderBy(d => d).ToList()` en `AgPlanificadorOleadas.cs`, alineando el patrón semanal con las ranuras materializadas.
- **Fase 2153:** Se aisló la ejecución concurrente en xUnit definiendo la colección `[CollectionDefinition("AgGeneradorHorarios", DisableParallelization = true)]` en `AgGeneradorHorariosReplayTests.cs` y anotando `AgGeneradorHorariosTests.cs`.
- **Fase 2154:** Se renombró formalmente el test de integración en `AgGeneradorHorariosTests.cs` a `EjecutarOptimizacionAsync_CampeonFinal_EntregaFechasSecuencialesEInvariantesI1I2`, asegurando validación de invariantes I1/I2 y secuencia de fechas sin asunciones espurias.
- **Fase 2155:** Se parametrizó `CrearContextoBase(bool conColeccionesIndexadas = false)` en `AgValidadorReglasOleadasTests.cs`, poblando e indexando colecciones base únicamente para pruebas del planificador de oleadas.
- **Fase 2156:** Se centralizó el helper canónico `ClaveCarril(AgGen g)` en el modelo de dominio `AgCarriles.cs` y se unificó su uso en `AgPlanificadorOleadas.cs` y `AgValidadorReglas.cs`.
- **Fase 2157:** Se sincronizó la numeración en `.agents/plans/REGISTRY.md` con contadores `ronda = 484` y `fase = 2158`, insertando la asignación `R-158`.
- **Fase 2158:** Se enriqueció el mensaje de rechazo de carril por no múltiplo con `I2_DemandaNoMultiplo({actual.TotalHoras:0.#}h/{actual.Unidad:0.#}h)` en `AgPlanificadorOleadas.cs`.

### Verificación de Principios (Ronda 484)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2150 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorOleadas.cs:1100-1105` |
| F2151 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorOleadas.cs:970-1040` |
| F2152 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorOleadas.cs:935-945` |
| F2153 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgGeneradorHorariosReplayTests.cs:18-22`, `AgGeneradorHorariosTests.cs:18-20` |
| F2154 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgGeneradorHorariosTests.cs:115-125` |
| F2155 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgValidadorReglasOleadasTests.cs:20-55, 350-355` |
| F2156 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgCarriles.cs:18-24`, `AgPlanificadorOleadas.cs:700-705`, `AgValidadorReglas.cs:334, 429` |
| F2157 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `REGISTRY.md:25-29, 194-196` |
| F2158 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorOleadas.cs:860-870` |

---

## Ronda 485 — Corrección de aserción del flag, aserciones I1/I2 en integración, validación estricta de día en gen único, eliminación de wrapper redundante y gate focalizado (Fases 2159–2163)

### Resumen de la Ronda
- **Fase 2159:** Se corrigió la prueba en `AgGeneradorHorariosReplayTests.cs` renombrándola a `Replay_FlagEstatico_EsRestauradoTrasPrueba`, probando la alternancia y restauración fiel del flag estático.
- **Fase 2160:** Se incorporaron las aserciones de invariantes I1/I2 en `AgGeneradorHorariosTests.cs`, verificando suma exacta de horas, sesiones proporcionales a la unidad de 4h y unicidad de fechas.
- **Fase 2161:** Se eliminó la excepción de gen único en `AgPlanificadorOleadas.cs`, forzando la concordancia estricta de `DiaSemana` para todas las asignaturas y emitiendo `DiaDesconocido` si `DiaSemana <= 0`.
- **Fase 2162:** Se actualizó la trazabilidad en `.agents/plans/REGISTRY.md` asentando las asignaciones R-158 y R-159 con evidencia del gate PASS (83 pruebas superadas).
- **Fase 2163:** Se suprimió el método puente `AgPlanificadorOleadas.ClaveCarril` migrando la agrupación hacia el helper canónico `AgCarriles.ClaveCarril(g)`.

### Verificación de Principios (Ronda 485)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2159 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgGeneradorHorariosReplayTests.cs:185-203` |
| F2160 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgGeneradorHorariosTests.cs:183-195` |
| F2161 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorOleadas.cs:960-985, 1045-1060` |
| F2162 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `REGISTRY.md:25-29, 194-196` |
| F2163 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorOleadas.cs:750-760` |

---

## Ronda 488 — Catálogo de dos opciones de partición (Fases 2179–2184)

### Resumen de la Ronda
- **Fase 2179:** Se implementaron los tipos de dominio `AgOpcionParticion.cs` y `AgOpcionesParticion.cs` en `ConversorPDF.XLL/Academic/Domain/Models/Genetica/` y se incluyeron en `ConversorPDF.XLL.csproj`. El catálogo resuelve `quedarse` y `escalar` determinísticamente respetando `MarVie`, `LunJue` y `FinDeSemana` con `MaxOpciones = 2`.
- **Fase 2180:** Se adaptó `AgParticionCarriles.Seleccionar` en `AgParticionCarriles.cs:71-125` para evaluar la lista de `opciones` acotada a 2 candidatas, verificando conservación de ranuras y alojamiento de patrones antes de evaluar, seleccionando la mejor partición sin priorizar artificialmente número de bandas.
- **Fase 2181:** Se propagó el parámetro opcional `opciones` en `AgPlanificadorOleadas.cs:30-58, 355-365` y `AgEvaluadorDemandaVentana.cs:103-138`, permitiendo que el evaluador inyecte las opciones resueltas hacia el planificador.
- **Fase 2182:** Se ajustó la evaluación de `antes` y `paso` en `AgEscaladorRanuraAula.cs:488-550` empleando `AgOpcionesParticion.Resolver` para cada ventana respectiva, garantizando que si el origen cabe no se desencadene escalado espurio.
- **Fase 2183:** Se extendió `Decision` con `ParticionPreferente` en `AgEscaladorRanuraAula.cs:53, 603-615`, se integró en `HuellaAccion` (línea 975) y se blindó `AplicarDetalladoAsync` (línea 883) para no borrar `dias_patron` cuando la decisión cuenta con partición preferente.
- **Fase 2184:** Se desarrollaron las pruebas unitarias en `AgOpcionesParticionTests.cs` y `AgParticionCarrilesOpcionesTests.cs`, y las pruebas de integración en `AgEscaladoTresCasosMatrizTests.cs`.

### Verificación de Principios (Ronda 488)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2179 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgOpcionParticion.cs`, `AgOpcionesParticion.cs`, `ConversorPDF.XLL.csproj:732-734` |
| F2180 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgParticionCarriles.cs:71-125, 286-292` |
| F2181 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorOleadas.cs:30-58, 355-365`, `AgEvaluadorDemandaVentana.cs:103-138` |
| F2182 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEscaladorRanuraAula.cs:488-550` |
| F2183 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEscaladorRanuraAula.cs:53, 603-615, 883, 975` |
| F2184 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgOpcionesParticionTests.cs`, `AgParticionCarrilesOpcionesTests.cs`, `AgEscaladoTresCasosMatrizTests.cs:595-655` |

---

## Ronda 489 — Catálogo de dos opciones — defectos de implementación (Fases 2185–2192)

### Resumen de la Ronda
- **Fase 2185:** Se corrigió `AgOpcionesParticion.Resolver` para operar exclusivamente sobre la ventana y ranuras provistas en cada invocación, desacoplando la escalera de ampliación y delegando la resolución de `quedarse` y `escalar` al escalador.
- **Fase 2186:** Se refinó `DetectarOrigen` en `AgOpcionesParticion.cs:40-100` verificando cardinalidad y paridad estricta de días ({2,3}, {4,5} para MAR-VIE; {1,2}, {3,4} para LUN-JUE), clasificando como `Ninguno` patrones terciarios como {3,4,5} o mezclas heterogéneas.
- **Fase 2187:** Se eliminó el retorno prematuro en `AgParticionCarriles.cs:92-125`, permitiendo fallback fluido a la enumeración contigua de `Particiones()` cuando el catálogo no arroja candidatas viables.
- **Fase 2188:** Se implementó el predicado de mejora comparativa frente a la referencia en `AgParticionCarriles.cs:105-120` exigiendo reducción de infactibles o makespan sin degradar patrones alojables.
- **Fase 2189:** Se obtuvo `ParticionPreferente` directamente desde `opcionUsada?.CarrilesTexto` en `AgEscaladorRanuraAula.cs:605-620`, asegurando que configuraciones de fin de semana con `6M` y `6T` se sellen con su formato de franja canónico.
- **Fase 2190:** Se implementó `OpcionConservaPatrones` y se vinculó en `AplicarDetalladoAsync` de `AgEscaladorRanuraAula.cs:880-900`, garantizando que `UPDATE ag_aula_asignatura SET dias_patron=''` se invoque únicamente si la partición adoptada rompe los patrones del aula.
- **Fase 2191:** Se actualizó `Aplica` en `AgEscaladorRanuraAula.cs:55` integrando la guarda `(!string.IsNullOrEmpty(ParticionPreferente) && ParticionMejora)` para habilitar remediación intra-ventana.
- **Fase 2192:** Se agregaron pruebas de contrato y casos límite en `AgOpcionesParticionTests.cs` y `AgParticionCarrilesOpcionesTests.cs` validando tolerancia a jornadas de 8h, descarte de patrones anómalos y neutralidad de opciones sin mejora.

### Verificación de Principios (Ronda 489)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2185 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgOpcionesParticion.cs:24-45`, `AgEscaladorRanuraAula.cs:498-560` |
| F2186 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgOpcionesParticion.cs:46-105` |
| F2187 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgParticionCarriles.cs:92-125` |
| F2188 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgParticionCarriles.cs:105-120` |
| F2189 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEscaladorRanuraAula.cs:605-620` |
| F2190 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEscaladorRanuraAula.cs:880-900, 970-980` |
| F2191 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEscaladorRanuraAula.cs:53-57` |
| F2192 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgOpcionesParticionTests.cs:130-180`, `AgParticionCarrilesOpcionesTests.cs:50-95` |

---

## Ronda 490 — Catálogo de opciones — cierre de aplicación y sellado (Fases 2193–2199)

### Resumen de la Ronda
- **Fase 2193:** Se corrigió la discrepancia de postcondición en `AgEscaladorRanuraAula.cs:854, 868, 934` utilizando `liberarEsperado = d.ConservaPatronOrigen ? null : d.IdAsignaturasBottleneck` tanto en la evaluación idempotente como en el cálculo del hash esperado y la verificación final de `EstadoCompleto`.
- **Fase 2194:** Se incorporaron los flags `pm` y `cp` a la lista de partes en `AgEscaladorRanuraAula.cs:HuellaAccion`, asegurando que `ParticionMejora` y `ConservaPatronOrigen` formen parte del hash criptográfico de la decisión.
- **Fase 2195:** Se unificó la invocación a `AgEvaluadorDemandaVentana.Evaluar` en `AgValidadorVentana.cs:350, 425` inyectando la opción resuelta por `AgOpcionesParticion.Resolver`, garantizando consistencia diagnóstica con el escalador.
- **Fase 2196:** Se anticipó el cálculo de `hayMejoraParticion` en `AgEscaladorRanuraAula.cs:570-600`, preservando la generación y aplicación de decisiones de reparación de carril cuando la ventana no cambia pero la partición optimiza la asignación.
- **Fase 2197:** Se desacopló la etiqueta de acción en `AgOpcionesParticion.cs` permitiendo clasificar opciones como `"quedarse"` o `"escalar"`, y se expuso `MaxOpcionesDecision = 2`.
- **Fase 2198:** Se normalizó la firma pública de `AgEvaluadorDemandaVentana.Evaluar` eliminando el parámetro redundante `opciones` y manteniendo únicamente `AgOpcionParticion opcion = null`.
- **Fase 2199:** Se implementaron tests unitarios y de persistencia: `AplicarDetalladoAsync_ConConservaPatron_NoRevierteYNoBorraPatrones_YEsIdempotente` en `AgEscaladorRanuraAulaPersistenciaTests.cs`, `ValidadorYEscalador_MismaParticion_ParaPatronesMarVie` en `AgValidadorVentanaTests.cs`, y `Resolver_Destino_MarcaAccionEscalar` en `AgOpcionesParticionTests.cs`.

### Verificación de Principios (Ronda 490)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2193 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEscaladorRanuraAula.cs:854, 868, 934` |
| F2194 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEscaladorRanuraAula.cs:994-997` |
| F2195 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgValidadorVentana.cs:261, 350, 425` |
| F2196 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEscaladorRanuraAula.cs:570-605` |
| F2197 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgOpcionesParticion.cs:14-45, 125-188` |
| F2198 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEvaluadorDemandaVentana.cs:103-138` |
| F2199 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEscaladorRanuraAulaPersistenciaTests.cs:630-710`, `AgValidadorVentanaTests.cs:1080-1140`, `AgOpcionesParticionTests.cs:215-240` |

---

## Ronda 491 — `ParticionMejora` tautológico y aplicación sin efecto (Fases 2200–2205)

### Resumen de la Ronda
- **Fase 2200:** Se incorporó el cómputo de `referencia` canónica sin opción y la comparación estructural `particionDifiere = !ParticionesIguales(evalMismaVentana.Plan?.Particion, referencia.Plan?.Particion)` en `AgEscaladorRanuraAula.cs:575-595`, eliminando la generación de propuestas espurias para aulas factibles cuya partición ya es óptima.
- **Fase 2201:** Se unificó el conjunto de patrones en `AgEscaladorRanuraAula.cs:576` seleccionando `patronesSimulados` ante reclasificaciones autorizadas para la comprobación de `OpcionConservaPatrones`.
- **Fase 2202:** Se implementó la rama `SinCambios` en `AgEscaladorRanuraAula.cs:980-990` asignando `Confirmadas = 0` y la advertencia descriptiva cuando una propuesta de carril no genera escrituras de base de datos.
- **Fase 2203:** Se añadió la propiedad `ParticionDifiereDeReferencia` a `AgEscaladorRanuraAula.Decision` y se actualizó la propiedad `Aplica` para requerir simultáneamente `ParticionMejora && ParticionDifiereDeReferencia`.
- **Fase 2204:** Se agregó validación de `AccionesValidas` en `AgOpcionesParticion.cs:28-35` y se actualizó `AgParticionCarriles.cs:97` con `MaxOpcionesDecision`.
- **Fase 2205:** Se añadieron pruebas exhaustivas en `AgEscaladoTresCasosMatrizTests.cs` (`AulaConPatronYParticionVigente_NoGeneraDecision`, `ParticionIgualAReferencia_NoAplica`, `ReclasificacionAutorizada_ConservaPatronUsaPatronesSimulados`), en `AgEscaladorRanuraAulaPersistenciaTests.cs` (`Aplicar_SoloParticionSinCambios_RetornaConfirmadasCeroYAdvertencia`) y en `AgOpcionesParticionTests.cs` (`Resolver_AccionInvalida_LanzaArgumentOutOfRangeException`).

### Verificación de Principios (Ronda 491)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2200 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEscaladorRanuraAula.cs:501-505, 575-590` |
| F2201 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEscaladorRanuraAula.cs:576` |
| F2202 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEscaladorRanuraAula.cs:980-995` |
| F2203 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEscaladorRanuraAula.cs:54-58, 638` |
| F2204 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgOpcionesParticion.cs:16, 28-33`, `AgParticionCarriles.cs:97` |
| F2205 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEscaladoTresCasosMatrizTests.cs:650-710`, `AgEscaladorRanuraAulaPersistenciaTests.cs:712-760`, `AgOpcionesParticionTests.cs:238-250` |

---

## Ronda 492 — Propuestas sin mejora y comparaciones de partición (Fases 2206–2211)

### Resumen de la Ronda
- **Fase 2206:** Se reestructuró `particionMejora` en `AgEscaladorRanuraAula.cs:585-595` haciendo que `particionDifiere` sea una condición obligatoria (`&&`) junto con la reducción estricta de infactibles o demanda pendiente.
- **Fase 2207:** Se recalculó `referencia` en `AgEscaladorRanuraAula.cs:565-575` sobre `ctxSimulado` ante reclasificaciones autorizadas, nivelando el estado de patrones.
- **Fase 2208:** Se implementó el atajo `opcionActual == null ? antes : ...` en `AgEscaladorRanuraAula.cs:570`, eliminando 35 evaluaciones redundantes por corrida en aulas sin catálogo.
- **Fase 2209:** Se incorporó el formateador de ranuras `TokenRanura` y la comparación contra `canonicaTexto` en `AgEscaladorRanuraAula.cs:580, 1025-1035`, distinguiendo franjas `6M` y `6T`.
- **Fase 2210:** Se eliminó `MaxOpciones` y `ResolverLista` de `AgOpcionesParticion.cs`, actualizando `AgOpcionesParticionTests.cs:214` a `MaxOpcionesDecision`.
- **Fase 2211:** Se sustituyó `decisiones.Any` por `decisionesAplicables.Any` en `AgEscaladorRanuraAula.cs:990-1000`, asegurando que `SinCambios` reporte `Confirmadas = 0` aunque existan decisiones no aplicables adyacentes.

### Verificación de Principios (Ronda 492)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2206 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEscaladorRanuraAula.cs:585-595` |
| F2207 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEscaladorRanuraAula.cs:565-575` |
| F2208 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEscaladorRanuraAula.cs:570` |
| F2209 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEscaladorRanuraAula.cs:580, 1025-1035` |
| F2210 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgOpcionesParticion.cs:14`, `AgOpcionesParticionTests.cs:214` |
| F2211 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEscaladorRanuraAula.cs:990-1000`, `AgEscaladorRanuraAulaPersistenciaTests.cs:765-815` |

---

## Ronda 494 — Corrección de las propuestas 2213 y comparación entre ventanas (Fases 2215–2216)

### Resumen de la Ronda
- **Fase 2215:** Se sustituyó la comparación contra `canonicas` por el conteo de no alojables sobre `referencia.Plan?.Particion` en `AgEscaladorRanuraAula.cs:600-610`, garantizando que solo se considere mejora si la opción reduce no alojables frente a la partición real del plan de referencia.
- **Fase 2216:** Se incorporó `referenciaComparada` en `AgEscaladorRanuraAula.cs:615-625` evaluando la ventana destino ante escalados, asegurando que la guarda de infactibles compare el plan del escalón con su propia referencia de destino.
- **Pruebas:** Se añadió `Escalon_ComparaInfactiblesConReferenciaDestino` en `AgEscaladoTresCasosMatrizTests.cs:765-785`.

### Verificación de Principios (Ronda 494)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2215 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEscaladorRanuraAula.cs:600-612` |
| F2216 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEscaladorRanuraAula.cs:615-625`, `AgEscaladoTresCasosMatrizTests.cs:765-785` |

---

## Ronda 493 — Catálogo inerte por filtro asimétrico y trazabilidad de rama (Fases 2218–2222)

### Resumen de la Ronda
- **Fase 2218:** Se centralizó la extracción de patrones efectivos en el helper `PatronesEfectivos` en `AgEscaladorRanuraAula.cs:1038-1048` excluyendo `AUTOMATICO`, alineando la semántica con `AgEvaluadorDemandaVentana.cs:124-126` y sustituyendo la consulta en `patronesAula` (líneas 493-497) y `patronesSimulados` (líneas 520-523).
- **Fase 2219:** Se niveló `conservaPatron` en `AgEscaladorRanuraAula.cs:586-590` evaluando con los patrones efectivos y requiriendo que `ContarPatronesNoAlojables(opcionUsada.Carriles, patronesConserva) == 0`, evitando el borrado de `dias_patron` durante `AplicarDetalladoAsync`.
- **Fase 2220:** Se dotó a `AgSeleccionParticion` de las propiedades `Origen` y `OpcionTexto` en `AgParticionCarriles.cs:43-65` y se propagaron en `Seleccionar` (retornando `"catalogo"` o `"enumeracion"`). En `AgPlanificadorOleadas.cs:365-385`, se formateó la línea `[Oleadas] Partición:` con `origen`, `opcion`, `cambiada`, `bandas`, `noAlojables`, `infactibles`, `makespan` y `candidatasEvaluadas`.
- **Fase 2221:** Se implementaron pruebas para patrones mixtos y trazabilidad: `PatronesLegadoYAutomatico_DetectarOrigenUsaSoloEfectivos` en `AgOpcionesParticionTests.cs`, `Seleccionar_ExponeOrigenYOpcionTexto` y `Seleccionar_CatalogoSinCambio_ReportaOrigenCatalogo` en `AgParticionCarrilesOpcionesTests.cs`, y `Escalador_ConPatronesMixtos_SeleccionaCatalogo` y `ConservaPatronOrigen_NoBorraDiasPatron` en `AgEscaladoTresCasosMatrizTests.cs`.
- **Fase 2222:** Se documentó la política de emisión (Info para `catalogo` siempre y para `enumeracion` solo con `cambiada=true`) y la guía de lectura en `docs/articles/changes/2026-11-14-debug-ranuras-corrida84.md`.

### Verificación de Principios (Ronda 493)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2218 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEscaladorRanuraAula.cs:493-497, 520-523, 1038-1048` |
| F2219 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEscaladorRanuraAula.cs:586-590` |
| F2220 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgParticionCarriles.cs:43-65, 123-128`, `AgPlanificadorOleadas.cs:365-385` |
| F2221 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgOpcionesParticionTests.cs`, `AgParticionCarrilesOpcionesTests.cs`, `AgEscaladoTresCasosMatrizTests.cs` |
| F2222 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `docs/articles/changes/2026-11-14-debug-ranuras-corrida84.md` |

---

## Ronda 495 — Catálogo con texto lossy y sin fallback (Fases 2223–2227)

### Resumen de la Ronda
- **Fase 2223:** Se restauró la continuidad hacia la enumeración contigua cuando el catálogo no mejora la referencia en `AgParticionCarriles.cs:128-135`, evitando la pérdida de mejoras de `Particiones()` y rotulando el origen como `catalogo_fallback`.
- **Fase 2224:** Se reemplazó el uso de `Clave` por `op.CarrilesTexto` en la lista evaluada en `AgParticionCarriles.cs:100-125`, preservando los literales de franja horaria (`5,6M|4,6T,7`).
- **Fase 2225:** Se garantizó que ante un catálogo sin mejora o enumeración sin cambio, `OpcionTexto` conserve la opción intentada (`textoOpcionIntentada`) en lugar de vaciarse en `AgParticionCarriles.cs:115, 155, 248`, actualizando la aserción correspondiente en `AgParticionCarrilesOpcionesTests.cs`.
- **Fase 2226:** Se reemplazó `LoggerHelper.Info` por `LoggerHelper.WarningLimitado` en `AgPlanificadorOleadas.cs:365-385`, deduplicando trazas por clave `OleadasParticion|{diagnostico}|{origen}|{cambiada}`.
- **Fase 2227:** Se crearon las pruebas `Catalogo_4567_OpcionTexto_ConservaFranjas`, `CatalogoSinMejora_EjecutaEnumeracion` y `CatalogoSinMejora_RegistraOpcionIntentada` en `AgParticionCarrilesOpcionesTests.cs`.

### Verificación de Principios (Ronda 495)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2223 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgParticionCarriles.cs:128-135, 152-167, 245-265` |
| F2224 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgParticionCarriles.cs:100-125` |
| F2225 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgParticionCarriles.cs:115, 155, 248`, `AgParticionCarrilesOpcionesTests.cs:228-229` |
| F2226 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorOleadas.cs:365-385` |
| F2227 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgParticionCarrilesOpcionesTests.cs:235-365` |

---

## Ronda 496 — `opcion` con dos significados y contador mixto (Fases 2228–2231)

### Resumen de la Ronda
- **Fase 2228:** Se unificó `OpcionTexto` para transportar fielmente la opción de catálogo evaluada (`textoOpcionIntentada`) y se incorporó `ParticionElegidaTexto` en `AgSeleccionParticion` para capturar la clave contigua elegida (`elegida.Clave`), evitando la sobrecarga semántica del campo en `AgParticionCarriles.cs:40-70, 260-270`.
- **Fase 2229:** Se añadieron contadores segregados en `AgSeleccionParticion`: `CandidatasEvaluadas` (contiguas) y `OpcionesEvaluadas` (catálogo), reflejando ambos en la traza de `AgPlanificadorOleadas.cs:365-388`.
- **Fase 2230:** Se corrigió la guarda de salida temprana en `AgParticionCarriles.cs:155-165` requiriendo `!catalogoEvaluado` para retornar, garantizando la ejecución de la enumeración contigua de fallback.
- **Fase 2231:** Se crearon las pruebas unitarias `CatalogoFallback_OpcionTextoEsLaIntentada`, `CatalogoFallback_ParticionElegidaEsLaContigua`, `CatalogoFallback_OpcionesEvaluadasEsUno` y `SinPatronesSinInfactibles_NoEnumera` en `AgParticionCarrilesOpcionesTests.cs`, y se actualizó la guía operativa en `docs/articles/changes/2026-11-14-debug-ranuras-corrida84.md`.

### Verificación de Principios (Ronda 496)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2228 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgParticionCarriles.cs:40-70, 260-270`, `AgPlanificadorOleadas.cs:375-385` |
| F2229 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgParticionCarriles.cs:40-70, 165-175`, `AgPlanificadorOleadas.cs:380-388` |
| F2230 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgParticionCarriles.cs:155-165` |
| F2231 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgParticionCarrilesOpcionesTests.cs:365-500`, `docs/articles/changes/2026-11-14-debug-ranuras-corrida84.md` |

---

## Ronda 497 — Catálogo exitoso reporta cero candidatas (Fase 2232)

### Resumen de la Ronda
- **Fase 2232:** Se pasó `evaluadasOpciones` en el parámetro `candidatasEvaluadas` del constructor `AgSeleccionParticion` en `AgParticionCarriles.cs:141`, documentando la semántica en el encabezado de clase y validando que en selección exitosa de catálogo `CandidatasEvaluadas > 0`.

### Verificación de Principios (Ronda 497)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2232 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgParticionCarriles.cs:40-45, 138-144`, `AgParticionCarrilesOpcionesTests.cs:520-560` |

---

## Ronda 498 — El árbol no marca la ventana seleccionada (Fases 2233–2237)

### Resumen de la Ronda
- **Fase 2233:** Se introdujo la comparación normalizada por ISO en `AgFormGeneradorHorarios.cs:5980-5990` a través de `AgNivel6Ventanas.Coincide`, permitiendo que valores guardados como `"1,2,3,4,5"`, `"1,2,3,4"` o alias textuales coincidan fielmente con el botón correspondiente.
- **Fase 2234:** Se integraron las 8 opciones de ventana de `AgDiasVentana.Opciones` en el árbol de Nivel 6, actualizando las proporciones horizontales en `CalcularColumnasNivel6` (`wJornada` 26%, `wDias` 56%, `wHoras` 18%) y proveyendo formateo adaptativo compacto en `AgNivel6Ventanas.FormatearRadio`.
- **Fase 2235:** Se sustituyó el hit-test fijo por `AgNivel6Ventanas.ResolverPorX` en `ProcesarClickDimensionNivel6`, verificando los 8 intervalos e imponiendo un límite de paso (`finBoton = Math.Min(posiciones[7] + step, fin)`) para suprimir la selección involuntaria de `VIE-DOM` en clics sobre el margen libre.
- **Fase 2236:** Se homogeneizó la emisión del tooltip en `ActualizarToolTipNivel6`, mostrando la clave canónica y la codificación ISO para cualquier ventana configurada.
- **Fase 2237:** Se implementó la suite completa de pruebas unitarias en `AgNivel6VentanasTests.cs` (índices canónicos e ISO, comparación cruzada, formateo adaptativo, hit-test por intervalo y descarte de espacio muerto).

### Verificación de Principios (Ronda 498)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2233 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgNivel6Ventanas.cs:65-85`, `AgFormGeneradorHorarios.cs:5985-5992` |
| F2234 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgNivel6Ventanas.cs:17-48`, `AgFormGeneradorHorarios.cs:5266-5310` |
| F2235 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgNivel6Ventanas.cs:120-160`, `AgFormGeneradorHorarios.cs:6055-6065` |
| F2236 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.cs:6110-6135` |
| F2237 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgNivel6VentanasTests.cs:1-150` |

---

## Ronda 499 — Clic de día rompe la distinción ISO `4,5,6,7` / `JUE-DOM` (Fases 2238–2242)

### Resumen de la Ronda
- **Fase 2238:** Se introdujo la guarda `EsAliasIsoTerminal` en `AgFormGeneradorHorarios.cs:6055-6065`, impidiendo que un clic sobre el radio `Jue-Dom` reescriba el valor si la celda ya almacena el alias terminal `"4,5,6,7"`.
- **Fase 2239:** Se aseguró la correspondencia directa del catálogo en `AgNivel6Ventanas.cs:15-55` vinculándolo a `AgDiasVentana.Opciones.Keys` y congelando `Claves` mediante `Array.AsReadOnly`.
- **Fase 2240:** Se depuraron los campos `DxLunJue`, `DxMarVie`, `DxJueSab`, `DxVieSab` y `DxVieDom` de `AgColumnasNivel6`, migrando todas las referencias visuales e interactivas al vector dinámico `DxDias[i]`.
- **Fase 2241:** Se resolvió el desborde visual en `CalcularColumnasNivel6` fijando el umbral base en 700 píxeles y recalculando `wDias` dinámicamente si la suma de anchuras supera el espacio disponible.
- **Fase 2242:** Se añadieron pruebas unitarias en `AgNivel6VentanasTests.cs` certificando la preservación del alias `4,5,6,7`, la inmutabilidad de la colección de claves y el formato responsivo de etiquetas.

### Verificación de Principios (Ronda 499)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2238 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgNivel6Ventanas.cs:60-65`, `AgFormGeneradorHorarios.cs:6055-6065` |
| F2239 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgNivel6Ventanas.cs:15-20, 50-55` |
| F2240 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.cs:5245-5260, 5295-5305` |
| F2241 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.cs:5265-5285` |
| F2242 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgNivel6VentanasTests.cs:150-220` |

---

## Ronda 500 — Sin forma de elegir `4,5,6,7` frente a `JUE-DOM` (Fases 2243–2244)

### Resumen de la Ronda
- **Fase 2243:** Se incorporó `"4,5,6,7"` como la novena opción explícita en `AgNivel6Ventanas.cs`, ajustando `Indice` y `Coincide` para priorizar la concordancia exacta e impedir que el cálculo ISO confunda `4,5,6,7` con `JUE-DOM`. Se removió la guarda restrictiva de clic en `AgFormGeneradorHorarios.cs:6051-6060` habilitando la selección bidireccional y se amplió el ancho mínimo del layout (`minimoTotal = 740`, `stepD = Math.Max(44, ...)`).
- **Fase 2244:** Se fijó el orden contractual del catálogo en `AgNivel6Ventanas.cs` y se extendió la suite en `AgNivel6VentanasTests.cs` con pruebas para la resolución de los 9 índices, la no colisión ISO entre `4,5,6,7` y `JUE-DOM`, la reasignación libre por clic y la consistencia total con `AgDiasVentana.Opciones`.

### Verificación de Principios (Ronda 500)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2243 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgNivel6Ventanas.cs:15-28, 60-120`, `AgFormGeneradorHorarios.cs:5265-5285, 6050-6060` |
| F2244 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgNivel6VentanasTests.cs:1-240` |

---

## Ronda 501 — Alias terminal con espacios y residuo sin uso (Fases 2245–2246)

### Resumen de la Ronda
- **Fase 2245:** Se implementó `NormalizarToken` en `AgNivel6Ventanas.cs:60-75` para depurar espacios en listas de días (e.g. `"4, 5, 6, 7"` -> `"4,5,6,7"`), resolviendo la degradación involuntaria a `JUE-DOM` en `Indice` y `Coincide`. Se agregaron casos de prueba unitaria en `AgNivel6VentanasTests.cs` validando el alias con espacios y separadores.
- **Fase 2246:** Se eliminó la declaración residual del campo estático no utilizado `EtiquetasVentanaRadio` en `AgFormGeneradorHorarios.cs:6088`.

### Verificación de Principios (Ronda 501)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2245 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgNivel6Ventanas.cs:60-145`, `AgNivel6VentanasTests.cs:50-95` |
| F2246 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.cs:6080-6090` |

---

## Ronda 502 — Corrección de asimetrías en conservación de patrones y fuentes de datos (Fases 2247–2250)

### Resumen de la Ronda
- **Fase 2247:** Se migró `OpcionConservaPatrones` en `AgEscaladorRanuraAula.cs` a semántica muchos-a-uno, removiendo la limitación inyectiva que anulaba la conservación de patrones en cohortes con más de 2 asignaturas.
- **Fase 2248:** Se eliminó la guarda textual de sábado (`6M`/`6T`), verificando en su lugar la existencia de ranuras para el día 6 directamente en las ranuras del carril evaluado.
- **Fase 2249:** Se condicionó la lectura de `aa.DiasPatron` en `AgPlanificadorPatronSemanal.DiasKDe` a que `aa.OrigenPatron` sea distinto de `AUTOMATICO`, unificando el comportamiento con `PatronesEfectivos`.
- **Fase 2250:** Se aplicó fallback seguro a `contexto.AulaAsignaturas` en `PatronesEfectivos` y `AgEvaluadorDemandaVentana.Evaluar` en caso de que `contexto.AulaAsignaturaPorAulaAsig` no se encuentre inicializado.

### Verificación de Principios (Ronda 502)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F2247 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEscaladorRanuraAula.cs:1050-1085`, `AgEscaladoTresCasosMatrizTests.cs:835-855` |
| F2248 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEscaladorRanuraAula.cs:1065-1085`, `AgEscaladoTresCasosMatrizTests.cs:855-870` |
| F2249 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorPatronSemanal.cs:28-35`, `AgEscaladoTresCasosMatrizTests.cs:870-890` |
| F2250 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEscaladorRanuraAula.cs:1040-1050`, `AgEvaluadorDemandaVentana.cs:124-135`, `AgEscaladoTresCasosMatrizTests.cs:890-910` |












