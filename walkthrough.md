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






























