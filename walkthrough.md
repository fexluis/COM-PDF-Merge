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






























