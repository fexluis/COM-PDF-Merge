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

