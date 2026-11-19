# Indice de Historial - implementation_plan.md

Escision: 2026-11-18 - commit `9b3dfd4` - trozo objetivo: 1000 lineas

Contenido historico congelado en `docs/plans/history/implementation_plan/`. Orden append-only: no se reordena ni se edita.

| Archivo | Fases | Rondas | Lineas |
|---|---|---|---|
| [implementation_plan_F0151_F0272.md](docs/plans/history/implementation_plan/implementation_plan_F0151_F0272.md) | 151-272 | - | 1228 |
| [implementation_plan_F0270_F0461.md](docs/plans/history/implementation_plan/implementation_plan_F0270_F0461.md) | 270-461 | 42, 43, 44 | 1207 |
| [implementation_plan_F0457_F0633.md](docs/plans/history/implementation_plan/implementation_plan_F0457_F0633.md) | 457-633 | 44 | 1234 |
| [implementation_plan_F0634_F0779.md](docs/plans/history/implementation_plan/implementation_plan_F0634_F0779.md) | 634-779 | 44 | 1250 |
| [implementation_plan_F0777_F0964.md](docs/plans/history/implementation_plan/implementation_plan_F0777_F0964.md) | 777-964 | 44 | 1234 |
| [implementation_plan_F0965_F1364.md](docs/plans/history/implementation_plan/implementation_plan_F0965_F1364.md) | 965-1364 | 232 | 1241 |
| [implementation_plan_F1311_F1602.md](docs/plans/history/implementation_plan/implementation_plan_F1311_F1602.md) | 1311-1602 | 264..317 | 1237 |
| [implementation_plan_F1603_F1819.md](docs/plans/history/implementation_plan/implementation_plan_F1603_F1819.md) | 1603-1819 | 318..398 | 473 |
| [implementation_plan_F1820_F2020.md](docs/plans/history/implementation_plan/implementation_plan_F1820_F2020.md) | 1820-2020 | 399..456 | 1226 |
| [`implementation_plan.md`](implementation_plan.md) (vivo) | posteriores al corte | ultimas 5 | 129 |

## Busqueda por fase
Toda `Fase N` se localiza en el primer archivo cuyo rango la contenga; ante solape de rangos, usar la columna Rondas.

## Citas historicas
Las citas `implementation_plan.md:NNNN` anteriores al corte apuntan al archivo monolitico; su ubicacion actual esta en [LINE_MAP.md](docs/plans/history/LINE_MAP.md).

---

## Ronda 457: Activación de tope estructural, exclusión de vacantes consentidas, compuerta estricta y calendario verificable (Fases 2021–2025)

### Resumen de la Ronda
- **Fase 2021:** Cálculo activo de déficit estructural `CalcularDeficitEstructural(AgContexto)` en `AgValidadorVentana.cs` y asignación a `contexto.MaxVacantesForzadas` en pre-vuelo P07 e inicio de optimización.
- **Fase 2022:** Exclusión de vacantes consentidas en `AsignaturasVacantesConsentidas` del cómputo de `forzadas` para la regla dura `R15_VacanteExcedida` en `AgValidadorReglas.cs` y en telemetría `[AG-CAUSAS]` de `AgGeneradorHorarios.cs`.
- **Fase 2023:** En `AgCompuertaPromocion.cs`, validación estricta de vacantes: cálculo incondicional de materias no consentidas, rechazo ante vacantes no consentidas cuando `!permitirVacantes`, y registro/trazabilidad explícita en `corrida.VacantesConsentidas` bajo justificación de promoción.
- **Fase 2024:** En `AgReglaViolacion.cs`, eliminación de alias duplicado ordinal 26 (`R9_VacanteSinCentinela`), fijando el identificador canónico único `R9_VacanteConDocenteReal = 26` y actualizando `AgPesosFitness.cs` y suites de pruebas.
- **Fase 2025:** En `AgOperadorMutacion.cs`, verificación de calendario propio mediante `AgCalendarioConflictos.Crear(genes, contexto).IsSuccess` antes de evaluar disponibilidad o liberar a vacante (evitando enmascarar `R14_BloqueHorarioNoAsignado`), y comparación tolerante case-insensitive en `EsNoAsignable`.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2021, 2022 | `AgValidadorVentana.cs`, `AgValidadorReglas.cs`, `AgGeneradorHorarios.cs`, `AgValidadorReglasTests.cs` | En ejecución (código en disco; pendiente verificación de compilación/tests) |
| 2023 | `AgCompuertaPromocion.cs` | En ejecución (código en disco; pendiente verificación de compilación/tests) |
| 2024 | `AgReglaViolacion.cs`, `AgPesosFitness.cs`, `AgValidadorReglasTests.cs` | En ejecución (código en disco; pendiente verificación de compilación/tests) |
| 2025 | `AgOperadorMutacion.cs`, `AgOperadorMutacionTests.cs` | En ejecución (código en disco; pendiente verificación de compilación/tests) |

---

## Ronda 458: Precisión de déficit estructural, desacoplamiento de factibilidad, optimización de mutación, compuerta no mutable y granularidad de consentimiento (Fases 2026–2030)

### Resumen de la Ronda
- **Fase 2026:** Corrección de `CalcularDeficitEstructural` en `AgValidadorVentana.cs`: agrupar por `IdAsignatura`, descontar asignaturas consentidas y computar grupos concurrentes por aulas distintas (`Distinct().Count()`), retornando `-1` ante contexto incompleto.
- **Fase 2027:** Desacoplamiento de `ValidarFactibilidad` en `AgValidadorVentana.cs` como método puro de solo lectura sin mutar `contexto.MaxVacantesForzadas`. Asignación centralizada en `AgGeneradorHorarios.cs` tras indexación.
- **Fase 2028:** En `AgOperadorMutacion.cs`, optimización de ruta crítica difiriendo `AgCalendarioConflictos.Crear` únicamente a la rama `else` cuando ningún alterno resulta disponible.
- **Fase 2029:** En `AgCompuertaPromocion.cs`, eliminación de mutación autogenerada sobre `corrida.VacantesConsentidas`; requerimiento de parámetro explícito `autorizacionVacantes` y registro trazable vía `LoggerHelper.Warning`.
- **Fase 2030:** En `AgContexto.cs`, incorporación de `ClavesVacantesConsentidas` y método `EstaConsentida(idAsignatura, idSede)` para soportar granularidad a nivel cohorte/sede sin invalidar wildcards de nivel asignatura; cableado en `AgValidadorReglas.cs` y `AgGeneradorHorarios.cs`.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2026, 2027 | `AgValidadorVentana.cs`, `AgGeneradorHorarios.cs`, `AgValidadorVentanaTests.cs` | En ejecución (código en disco; pendiente verificación de compilación/tests) |
| 2028 | `AgOperadorMutacion.cs`, `AgOperadorMutacionTests.cs` | En ejecución (código en disco; pendiente verificación de compilación/tests) |
| 2029 | `AgCompuertaPromocion.cs` | En ejecución (código en disco; pendiente verificación de compilación/tests) |
| 2030 | `AgContexto.cs`, `AgValidadorReglas.cs`, `AgGeneradorHorarios.cs`, `AgValidadorReglasTests.cs` | En ejecución (código en disco; pendiente verificación de compilación/tests) |

---

## Ronda 459: Concurrencia temporal en déficit, proveedor de autorización de vacantes, unificación de consentimiento, productor de claves y cobertura de pruebas (Fases 2031–2035)

### Resumen de la Ronda
- **Fase 2031:** Corrección de `CalcularDeficitEstructural` en `AgValidadorVentana.cs`: cálculo de concurrencia temporal dividiendo aulas distintas por carriles/franjas activos (`Math.Ceiling(aulasDistintas / carriles)`), evitando sobreestimar el déficit a 92 cuando el real es ≈8.
- **Fase 2032:** En `AgPromotorVinculaciones.cs`, parseo de asignaturas autorizadas desde `justificacionVacante` (`ParseAutorizacionVacantes`) y paso a `AgCompuertaPromocion.ValidarPromocion`. En `AgFormGeneradorHorarios.cs`, enriquecimiento de la justificación con las asignaturas no consentidas y provisión de candidatos autorizables a `EsPromovible` para habilitar el botón de promoción.
- **Fase 2033:** Unificación de fuentes de consentimiento en `AgCompuertaPromocion.cs` y `AgValidadorVentana.cs` utilizando `contexto.PermiteVacantes` y `contexto.EstaConsentida(idAsignatura, idSede)`, previniendo rechazos falsos cuando el consentimiento está en claves por sede.
- **Fase 2034:** En `AgFormGeneradorHorarios.cs`, producción de `ClavesVacantesConsentidas` por par asignatura:sede en pre-vuelo y soporte bidireccional en serialización/deserialización de `ParametrosJson["vacantesConsentidas"]` (int y `asig:sede`). Retirado el branch redundante `idAsignatura.ToString()` en `EstaConsentida`.
- **Fase 2035:** Actualización y adición de pruebas unitarias en `AgValidadorVentanaTests.cs` (concurrencia temporal vs aulas totales) y en `AgCompuertaPromocionTests.cs` (autorización de no consentidas, autorización incompleta, claves de sede y parseo de justificación).

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2031 | `AgValidadorVentana.cs`, `AgValidadorVentanaTests.cs` | En ejecución (código en disco; pendiente verificación de compilación/tests) |
| 2032 | `AgCompuertaPromocion.cs`, `AgPromotorVinculaciones.cs`, `AgFormGeneradorHorarios.cs` | En ejecución (código en disco; pendiente verificación de compilación/tests) |
| 2033 | `AgCompuertaPromocion.cs`, `AgContexto.cs`, `AgValidadorVentana.cs` | En ejecución (código en disco; pendiente verificación de compilación/tests) |
| 2034 | `AgContexto.cs`, `AgFormGeneradorHorarios.cs`, `AgCompuertaPromocion.cs` | En ejecución (código en disco; pendiente verificación de compilación/tests) |
| 2035 | `AgValidadorVentanaTests.cs`, `AgCompuertaPromocionTests.cs`, `AgValidadorReglasTests.cs` | En ejecución (código en disco; pendiente verificación de compilación/tests) |

---

## Ronda 460: Concurrencia por cohortes en déficit estructural, diálogo de autorización, parser fail-closed de justificación, preservación de claves en corrida y cobertura de pruebas (Fases 2036–2040)

### Resumen de la Ronda
- **Fase 2036:** En `AgValidadorVentana.cs`, sustitución del divisor arbitrario de bloques por concurrencia efectiva de cohortes (aula, semestre, grupo, sede) sin divisor global, garantizando que el tope estructural no estrangule las liberaciones necesarias (cap >= 8).
- **Fase 2037:** Creación de `FormAutorizacionVacantes.cs` para selección explícita con checkboxes de vacantes a autorizar en `AgFormGeneradorHorarios.cs`. Eliminación de autoautorización incondicional en llamadores y evaluación de `EsPromovible` con `_vacantesAutorizadas`.
- **Fase 2038:** En `AgPromotorVinculaciones.cs`, parser `ParseAutorizacionVacantes` fail-closed exigiendo prefijo explícito `(?:asig(?:naturas)?|materias?)\s*[=:]\s*([0-9][0-9,\s]*)` para prevenir autorizaciones espurias derivadas de texto libre.
- **Fase 2039:** En `AgCompuertaPromocion.cs`, retorno dual `(HashSet<int> Asignaturas, HashSet<string> Claves)` en `VacantesConsentidasEnCorrida`, preservando la granularidad de sede al rehidratar corridas en `AgFormGeneradorHorarios.cs`.
- **Fase 2040:** Actualización de pruebas unitarias en `AgValidadorVentanaTests.cs` (concurrencia por cohorte/sede y caso corrida 109 cap >= 8) y en `AgCompuertaPromocionTests.cs` (rechazo ante autorización parcial de múltiples vacantes y parser fail-closed).

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2036 | `AgValidadorVentana.cs`, `AgValidadorVentanaTests.cs` | En ejecución (código en disco; pendiente verificación de compilación/tests) |
| 2037 | `FormAutorizacionVacantes.cs`, `AgFormGeneradorHorarios.cs`, `AgCompuertaPromocionTests.cs` | En ejecución (código en disco; pendiente verificación de compilación/tests) |
| 2038 | `AgPromotorVinculaciones.cs`, `AgCompuertaPromocion.cs`, `AgCompuertaPromocionTests.cs` | En ejecución (código en disco; pendiente verificación de compilación/tests) |
| 2039 | `AgCompuertaPromocion.cs`, `AgFormGeneradorHorarios.cs` | En ejecución (código en disco; pendiente verificación de compilación/tests) |
| 2040 | `AgValidadorVentanaTests.cs`, `AgCompuertaPromocionTests.cs` | En ejecución (código en disco; pendiente verificación de compilación/tests) |

---

## Ronda 461: Registro de nuevos archivos en csproj, cohorte académica pura en déficit, aislamiento de permisos entre corridas, compuerta solicitable y granularidad completa en justificación (Fases 2041–2045)

### Resumen de la Ronda
- **Fase 2041:** En `ConversorPDF.XLL.csproj`, inclusión explícita de `FormAutorizacionVacantes.cs` (`<SubType>Form</SubType>`) e `ItemVacanteDto.cs` en `<Compile Include=...>` para evitar CS0246 durante compilación.
- **Fase 2042:** En `AgValidadorVentana.CalcularDeficitEstructural`, eliminación de `IdAula` de la clave de cohorte `(Semestre, Grupo, Sede)` para evitar que cada aula física infle el déficit al límite de 92 aulas. Soporte opcional de `InformeFactibilidad.DeficitDocentePorAsignatura` para cálculo directo desde pre-vuelo.
- **Fase 2043:** En `AgFormGeneradorHorarios.cs`, reinicio obligatorio de `_vacantesAutorizadas` y `_vacantesAutorizadasClaves` en el ciclo de vida de la corrida (`CargarCorridaDesdeComboAsync`, `IniciarOptimizacionAsync`, `RecalcularViolaciones`), e intersección estricta con las vacantes no consentidas del individuo cargado para evitar fugas de permisos entre corridas.
- **Fase 2044:** En `AgCompuertaPromocion.cs`, sustitución del workaround `EsPromovible(...) || esViableParaPromover` por el método formal `EsPromovibleSolicitable(...)` (validando HARD == 0 y staging coherente), separando la viabilidad previa para abrir el diálogo de la autorización estricta en el clic de promoción.
- **Fase 2045:** En `ItemVacanteDto.cs`, `FormAutorizacionVacantes.cs`, `AgPromotorVinculaciones.cs`, `AgCompuertaPromocion.cs` y `AgFormGeneradorHorarios.cs`, preservación de la granularidad `asig:sede` en todo el flujo de autorización y justificación, con soporte dual en el parser fail-closed para IDs y claves compuestas. Actualización de tests en `AgValidadorVentanaTests.cs` y `AgCompuertaPromocionTests.cs`.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2041 | `ConversorPDF.XLL.csproj` | En ejecución (código en disco; pendiente verificación de compilación/tests) |
| 2042 | `AgValidadorVentana.cs`, `InformeFactibilidad.cs`, `AgGeneradorHorarios.cs`, `AgValidadorVentanaTests.cs` | En ejecución (código en disco; pendiente verificación de compilación/tests) |
| 2043 | `AgFormGeneradorHorarios.cs` | En ejecución (código en disco; pendiente verificación de compilación/tests) |
| 2044 | `AgCompuertaPromocion.cs`, `AgFormGeneradorHorarios.cs`, `AgCompuertaPromocionTests.cs` | En ejecución (código en disco; pendiente verificación de compilación/tests) |
| 2045 | `ItemVacanteDto.cs`, `FormAutorizacionVacantes.cs`, `AgPromotorVinculaciones.cs`, `AgCompuertaPromocion.cs`, `AgFormGeneradorHorarios.cs`, `AgCompuertaPromocionTests.cs` | En ejecución (código en disco; pendiente verificación de compilación/tests) |

---

## Ronda 462: Productor del déficit en pre-vuelo, soporte de vacantes sin sede (asig:0), botón consistente en finally de promoción y compuerta de promoción de solo lectura (Fases 2046–2049)

### Resumen de la Ronda
- **Fase 2046:** En `AgValidadorVentana.cs`, cálculo de `informe.DeficitDocentePorAsignatura` en `ValidarFactibilidad` evaluando la concurrencia máxima por franja/jornada frente a docentes habilitados del pool (necesidad real acotada entre 8 y 12). Documentación del fallback en `CalcularDeficitEstructural` con emisión de `WARN [AG-TOPE] fallback usado: {deficitTotal}`.
- **Fase 2047:** En `ItemVacanteDto.cs`, clave simplificada a `IdAsignatura.ToString()` cuando `IdSede == 0`. En `FormAutorizacionVacantes.cs`, bifurcación a `Autorizadas` o `AutorizadasClaves` según si la sede es válida. En `AgPromotorVinculaciones.cs`, parseo seguro de tokens `asig:0` sin descarte hacia `asigs`.
- **Fase 2048:** En `AgFormGeneradorHorarios.cs:4975`, sustitución de `EsPromovible` en el bloque `finally` de `btnPromoverProduccion_Click` por `EsPromovibleSolicitable(...)`, alineándolo con los demás puntos de recarga/recalculo y permitiendo reabrir el diálogo de autorización con vacantes no consentidas.
- **Fase 2049:** En `AgCompuertaPromocion.cs`, eliminación de mutación en `contexto.ClavesVacantesConsentidas` dentro de `ValidarPromocion`. Implementación del predicado local de solo lectura `ConsentidaLocal` combinando fuentes de contexto y corrida de forma inmutable.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2046 | `AgValidadorVentana.cs`, `AgValidadorVentanaTests.cs` | En ejecución (código en disco; pendiente verificación de compilación/tests) |
| 2047 | `ItemVacanteDto.cs`, `FormAutorizacionVacantes.cs`, `AgPromotorVinculaciones.cs`, `AgCompuertaPromocionTests.cs` | En ejecución (código en disco; pendiente verificación de compilación/tests) |
| 2048 | `AgFormGeneradorHorarios.cs` | En ejecución (código en disco; pendiente verificación de compilación/tests) |
| 2049 | `AgCompuertaPromocion.cs`, `AgCompuertaPromocionTests.cs` | En ejecución (código en disco; pendiente verificación de compilación/tests) |

---

## Ronda 463: Verificación de Wave 462 sin defectos bloqueantes y análisis de precisión del déficit (Fases 2050–2051)

### Resumen de la Ronda
- **Verificación Wave 462:** Confirmada la correcta implementación de Fases 2046–2049 (productor de déficit en pre-vuelo, soporte de vacantes `asig:0`, botón consistente en finally de promoción y compuerta inmutable). Cero defectos bloqueantes en compilación y runtime.
- **Fase 2050 (Opcional):** Evaluación de concurrencia fina por patrón/franja vs jornada en `AgValidadorVentana.cs`. Clasificada como mejora futura de precisión no bloqueante.
- **Fase 2051 (Opcional):** Análisis del momento de cálculo de déficit en informe P07 previo a diálogo de consentimiento. Clasificada como mejora cosmética de documentación no bloqueante.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2050 | `AgValidadorVentana.cs` | Opcional / Evaluada sin acción bloqueante |
| 2051 | `AgValidadorVentana.cs`, `AgFormGeneradorHorarios.cs` | Opcional / Evaluada sin acción bloqueante |

---

## Ronda 464: Consentimiento sobre demanda real de la corrida y separación de vacantes estructurales (R16) de liberadas (R15) (Fases 2052–2056)

### Resumen de la Ronda
- **Fase 2052:** Creación de `AgConsentimientoVacantes.cs` y `VacanteEstructuralInfo.cs`. El consentimiento de vacantes se calcula estrictamente sobre la demanda real aula-asignatura de la corrida (`ctx.AulaAsignaturas` con `AulasPorId`), eliminando la contaminación de asignaturas del catálogo de programas homónimos. Unificación del cálculo en `AgFormGeneradorHorarios.cs`.
- **Fase 2053:** En `AgContexto.cs`, incorporación de `PoolExiste(idAsignatura)` y `PoolHabilitado(idAsignatura, idPrograma, semestre)`. En `AgFabricaIndividuos.cs`, orden de causas en `ClasificarMotivo` discriminando `SIN_CANDIDATOS` (sin registro en pool) frente a `NO_CUMPLE_REQUISITOS` (docente registrado pero no habilitado para el programa/semestre). Telemetría de `PoolEfectivo` en la carga del contexto.
- **Fase 2054:** En `AgReglaViolacion.cs` y `AgPesosFitness.cs`, incorporación de `R16_VacanteEstructuralNoConsentida`. En `AgValidadorReglas.cs`, separación de R15 (vacantes liberadas por el AG frente a `MaxVacantesForzadas`) de R16 (vacantes estructurales no consentidas, emitiendo una infracción agregada por asignatura).
- **Fase 2055:** Diálogo informado en `AgFormGeneradorHorarios.cs` (pre-vuelo y optimización) agrupando asignaturas demandadas con detalle de cohortes/sede y motivo. Persistencia dual íntegra (`int` y `"asig:sede"`) en `ParametrosJson` y `dict["vacantesConsentidas"]`.
- **Fase 2056:** Registro de `AgConsentimientoVacantes.cs` y `VacanteEstructuralInfo.cs` en `ConversorPDF.XLL.csproj`. Pruebas unitarias y de regresión en `AgConsentimientoVacantesTests.cs`, `AgFabricaIndividuosTests.cs` y `AgValidadorReglasTests.cs` cubriendo el caso Corrida #110 (`HARD=0`, `R16=0`).

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2052 | `AgConsentimientoVacantes.cs`, `VacanteEstructuralInfo.cs`, `AgFormGeneradorHorarios.cs`, `ConversorPDF.XLL.csproj` | Implementado (código y tests en disco; Paso 6 sin compilación) |
| 2053 | `AgContexto.cs`, `AgFabricaIndividuos.cs`, `AgFormGeneradorHorarios.cs` | Implementado (código y tests en disco; Paso 6 sin compilación) |
| 2054 | `AgReglaViolacion.cs`, `AgPesosFitness.cs`, `AgValidadorReglas.cs` | Implementado (código y tests en disco; Paso 6 sin compilación) |
| 2055 | `AgFormGeneradorHorarios.cs`, `AgCompuertaPromocion.cs` | Implementado (código y tests en disco; Paso 6 sin compilación) |
| 2056 | `ConversorPDF.XLL.csproj`, `AgConsentimientoVacantesTests.cs`, `AgFabricaIndividuosTests.cs`, `AgValidadorReglasTests.cs` | Implementado (código y tests en disco; Paso 6 sin compilación) |

---

## Ronda 465: Corrección de regresión centinela, alineación de semilla cohorte y granularidad sin wildcard prematuro (Fases 2057–2061)

### Resumen de la Ronda
- **Fase 2057:** En `AgFabricaIndividuos.cs`, restauración de la exclusión estricta de `IdTerceroVacante` al evaluar candidatos activos e inactivos reales en `ClasificarMotivo`. Evita que un pool que contiene únicamente al centinela clasifique erróneamente como `NO_CUMPLE_REQUISITOS` en lugar de `SIN_CANDIDATOS` (regresión F1996).
- **Fase 2058:** Incorporación de sobrecarga `ClasificarMotivo(c, a, semestre, idPrograma)` en `AgFabricaIndividuos.cs`, invocada tanto en la fábrica cromosómica como en `AgConsentimientoVacantes.cs` con `aula.Semestre` y `idProg`, unificando la evaluación para asignaturas MultiSemestre.
- **Fase 2059:** En `AgConsentimientoVacantes.cs`, agrupación por asignatura y emisión de wildcard (`asigs.Add`) únicamente si **todas** las cohortes demandadas de la asignatura carecen de docente habilitado; de lo contrario, se mantiene el consentimiento acotado por clave `asig:sede`, previniendo enmascarar sedes con docentes.
- **Fase 2060:** En `AgValidadorReglas.cs`, incorporación de `OrigenVacante` y `ClasificarOrigen`. Vacantes con motivos de causas no docentes (`SIN_PLAN_OLEADA`, `SIN_BLOQUE_NIVEL`) se excluyen del cálculo de vacantes liberadas en `R15` y de vacantes estructurales en `R16`, evitando inflar falsos positivos cuando el plan de oleadas es infactible.
- **Fase 2061:** Seguimiento git mediante `git add` de los 3 archivos nuevos (`AgConsentimientoVacantes.cs`, `VacanteEstructuralInfo.cs`, `AgConsentimientoVacantesTests.cs`), actualización de contadores (§2 `ronda=465`, `fase=2061`) y registro de fila R-139 en `REGISTRY.md`. Pruebas unitarias de cobertura añadidas en `AgFabricaIndividuosTests.cs`, `AgConsentimientoVacantesTests.cs` y `AgValidadorReglasTests.cs`.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2057 | `AgFabricaIndividuos.cs`, `AgFabricaIndividuosTests.cs` | Implementado (código y tests en disco; Paso 6 sin compilación) |
| 2058 | `AgFabricaIndividuos.cs`, `AgConsentimientoVacantes.cs`, `AgFabricaIndividuosTests.cs`, `AgConsentimientoVacantesTests.cs` | Implementado (código y tests en disco; Paso 6 sin compilación) |
| 2059 | `AgConsentimientoVacantes.cs`, `AgConsentimientoVacantesTests.cs` | Implementado (código y tests en disco; Paso 6 sin compilación) |
| 2060 | `AgValidadorReglas.cs`, `AgValidadorReglasTests.cs` | Implementado (código y tests en disco; Paso 6 sin compilación) |
| 2061 | `git add`, `REGISTRY.md`, `AgConsentimientoVacantesTests.cs`, `AgFabricaIndividuosTests.cs`, `AgValidadorReglasTests.cs` | Implementado (archivos seguidos; Paso 6 sin compilación) |

---

## Ronda 466: Aislamiento de motivos no docentes en fallos de patrón, filtro de aulas generables y reutilización de consentimiento informado (Fases 2062–2065)

### Resumen de la Ronda
- **Fase 2062:** En `AgFabricaIndividuos.cs:715`, corrección del motivo de vacante cuando falla la construcción del patrón semanal (`resPatron.IsFailure`), asignando `AgMotivoVacante.SIN_PLAN_OLEADA` en lugar de `SIN_DISPONIBILIDAD`. Previene que fallas temporales del plan cuenten falsamente como liberadas por el AG e inflen `R15_VacanteExcedida`.
- **Fase 2063:** En `AgConsentimientoVacantes.cs`, filtrado de cohortes demandadas exigiendo `ctx.ManifiestoCobertura == null || ctx.PuedeGenerarAula(aa.IdAula)`, evitando consentir o listar aulas bloqueadas o excluidas de la corrida.
- **Fase 2064:** En `AgFormGeneradorHorarios.cs`, optimización de UX evaluando `_contextoActual.AsignaturasVacantesConsentidas.SetEquals(asigsOpt) && _contextoActual.ClavesVacantesConsentidas.SetEquals(clavesOpt)`. Si el consentimiento de P07 está vigente e idéntico, se reutiliza sin duplicar el diálogo modal ante el usuario.
- **Fase 2065:** Trazabilidad formal en `REGISTRY.md` con fila R-140 y contadores §2 (`ronda=466`, `fase=2065`). Pruebas unitarias de soporte implementadas en `AgFabricaIndividuosTests.cs`, `AgValidadorReglasTests.cs` y `AgConsentimientoVacantesTests.cs`.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2062 | `AgFabricaIndividuos.cs`, `AgFabricaIndividuosTests.cs`, `AgValidadorReglasTests.cs` | Implementado (código y tests en disco; Paso 6 sin compilación) |
| 2063 | `AgConsentimientoVacantes.cs`, `AgConsentimientoVacantesTests.cs` | Implementado (código y tests en disco; Paso 6 sin compilación) |
| 2064 | `AgFormGeneradorHorarios.cs` | Implementado (código y tests en disco; Paso 6 sin compilación) |
| 2065 | `REGISTRY.md`, `AgConsentimientoVacantesTests.cs`, `AgFabricaIndividuosTests.cs`, `AgValidadorReglasTests.cs` | Implementado (trazabilidad y tests en disco; Paso 6 sin compilación) |

---

## Ronda 467: Consentimiento informado idempotente y unificado, erradicación de doble diálogo P07/optimización y cobertura de pruebas (Fases 2066–2068)

### Resumen de la Ronda
- **Fase 2066:** Extracción de método unificado e idempotente `AsegurarConsentimientoVacantes(bool interactivo)` en `AgFormGeneradorHorarios.cs` (líneas 1446–1496), llamado tanto en diagnóstico pre-vuelo P07 (`EjecutarDiagnosticoPreVueloAsync`) como en optimización (`IniciarOptimizacionAsync`). Elimina duplicación de lógica y garantiza reutilización limpia basada en `SetEquals`.
- **Fase 2067:** Cobertura de pruebas unitarias en `AgConsentimientoVacantesTests.cs` evaluando: inclusión total de aulas cuando el manifiesto es nulo, exclusión de aulas no generables cuando el manifiesto está construido, reutilización vía `SetEquals` ante mismo universo, y requerimiento de confirmación (no-reutilización) ante cambios en claves compuestas.
- **Fase 2068:** Trazabilidad formal en `REGISTRY.md` con fila R-141 y contadores §2 (`ronda=467`, `fase=2068`). Sincronización de planes vivos (`implementation_plan.md`, `walkthrough.md`) y seguimiento en git.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2066 | `AgFormGeneradorHorarios.cs` | Implementado (código en disco; Paso 6 sin compilación) |
| 2067 | `AgConsentimientoVacantesTests.cs` | Implementado (tests en disco; Paso 6 sin compilación) |
| 2068 | `REGISTRY.md`, `implementation_plan.md`, `walkthrough.md` | Implementado (trazabilidad y documentación; Paso 6 sin compilación) |

---

## Ronda 468: Erradicación definitiva del doble diálogo de consentimiento, unificación de logs y helper puro EsVigente (Fases 2069–2072)

### Resumen de la Ronda
- **Fase 2069:** Eliminación de la invocación de `AsegurarConsentimientoVacantes` en el pre-vuelo P07 (`EjecutarDiagnosticoPreVueloAsync`). El consentimiento informado queda asegurado en un único punto canónico en `IniciarOptimizacionAsync` tras construirse el manifiesto y evaluarse el escalado, erradicando el doble diálogo modal.
- **Fase 2070:** Homogeneización de etiquetas de log en `AgFormGeneradorHorarios.cs:1446-1496` bajo el prefijo unificado `[CONSENTIMIENTO] Reutilizando universo vigente...`, `[CONSENTIMIENTO] Cancelado por el usuario...` y `[CONSENTIMIENTO] Aceptado por el usuario...`, evitando la mezcla engañosa con `[OPTIMIZACIÓN]` previa a la corrida.
- **Fase 2071:** Extracción del método puro `AgConsentimientoVacantes.EsVigente(AgContexto, HashSet<int>, HashSet<string>)` y cobertura de pruebas en `AgConsentimientoVacantesTests.cs`: `EsVigente_WhenMismoUniverso_ShouldBeTrue`, `EsVigente_WhenFaltaUnaClave_ShouldBeFalse` y `EsVigente_WhenAsignaturasNulas_ShouldBeFalse`.
- **Fase 2072:** Registro formal en `REGISTRY.md` con fila R-142 y contadores §2 (`ronda=468`, `fase=2072`), verificación de umbrales de rotación e indexación en control de versiones.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2069 | `AgFormGeneradorHorarios.cs` | Implementado (código en disco; Paso 6 sin compilación) |
| 2070 | `AgFormGeneradorHorarios.cs` | Implementado (código en disco; Paso 6 sin compilación) |
| 2071 | `AgConsentimientoVacantes.cs`, `AgConsentimientoVacantesTests.cs` | Implementado (código y tests en disco; Paso 6 sin compilación) |
| 2072 | `REGISTRY.md`, `implementation_plan.md`, `walkthrough.md` | Implementado (trazabilidad y documentación; Paso 6 sin compilación) |

---

## Ronda 469: Precondición de manifiesto antes de consentimiento en corrida y cierre de ciclo (Fases 2073–2074)

### Resumen de la Ronda
- **Fase 2073:** Reordenamiento estricto en `IniciarOptimizacionAsync` de `AgFormGeneradorHorarios.cs:2690-2705`: el guard `_contextoActual.ManifiestoCobertura == null || _contextoActual.ValidacionPlanAprobado == null` se evalúa **antes** de `AsegurarConsentimientoVacantes`. Esto garantiza que nunca se dispare el diálogo modal ni se registren conjuntos de vacantes si la corrida carece de manifiesto válido.
- **Fase 2074:** Cierre y verificación integral de las olas 464–469, trazabilidad en `REGISTRY.md` con fila R-143 y contadores §2 (`ronda=469`, `fase=2074`), ejecución de gate y empaquetado release v1.6.4.31.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2073 | `AgFormGeneradorHorarios.cs` | Implementado |
| 2074 | `REGISTRY.md`, `implementation_plan.md`, `walkthrough.md`, release scripts | En ejecución / Cierre |

---

## Ronda 470: Corrección de advertencia Fecha Límite Excedida en Sesiones Académicas según ventana canónica de período (Fases 2075–2078)

### Resumen de la Ronda
- **Fase 2075:** Creación del servicio desacoplado `ICierrePeriodoResolver` y `CierrePeriodoResolver` en `ConversorPDF.XLL.Academic.Application` junto a su DTO `ResultadoCierrePeriodo`, estableciendo la precedencia canónica: ventana temporal del período (`periodo.TryObtenerVentana`) como fuente primaria, y fallback al 15/12 del año resuelto solo si el período carece de ventana válida. Inclusión de constantes documentadas `DiaCierreSemestre = 15` y `MesCierreSemestre = 12` (hallazgo H8).
- **Fase 2076:** Sustitución de la regla rígida del 15/12 en `FormSesionesAcademicas.cs:683-691` por evaluación con `_cierrePeriodoResolver.Evaluar(periodoSeleccionado, dtpFinal.Value.Date, DateTime.Today)`, preservando el diálogo modal confirmable y emitiendo traza diagnóstica `LoggerHelper.Warning` con el origen evaluado.
- **Fase 2077:** Extensión de la consulta de catálogos en `CatalogCacheService.cs` para incluir `PeriodoInicio` y `PeriodoFin` con protección ante esquemas antiguos (`pragma_table_info`), e invalidación explícita del caché en `FormPeriodos.cs` (guardar/eliminar) y `AgFormGeneradorHorarios.cs` (extensión de ventana en corrida). Inclusión de los nuevos archivos en `ConversorPDF.XLL.csproj`.
- **Fase 2078:** Creación de la suite de pruebas unitarias `CierrePeriodoResolverTests.cs` en `ConversorPDF.Tests` con xUnit + FluentAssertions, cubriendo la regresión del defecto (período 2027-01 con fecha 02/02/2027), rebase de fin de período, fallbacks de año y formato de mensajes.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2075 | `ICierrePeriodoResolver.cs`, `CierrePeriodoResolver.cs`, `ResultadoCierrePeriodo.cs` | Implementado |
| 2076 | `FormSesionesAcademicas.cs` | Implementado |
| 2077 | `CatalogCacheService.cs`, `FormPeriodos.cs`, `AgFormGeneradorHorarios.cs`, `ConversorPDF.XLL.csproj` | Implementado |
| 2078 | `CierrePeriodoResolverTests.cs` | Implementado |

---

## Ronda 471: Corrección de mojibake en servicio nuevo y coherencia de caché de períodos (Fases 2079–2082)

### Resumen de la Ronda
- **Fase 2079:** Reescritura limpia en UTF-8 con BOM de `CierrePeriodoResolver.cs`, `ICierrePeriodoResolver.cs`, `ResultadoCierrePeriodo.cs` y `CierrePeriodoResolverTests.cs`, erradicando la doble codificación CP1252/UTF-8 (mojibake) en literales de mensaje al usuario y documentación XML.
- **Fase 2080:** Reemplazo de la invalidación destructiva global `CatalogCacheService.Invalidate()` por métodos granulares no destructivos `CatalogCacheService.ActualizarPeriodo` y `EliminarPeriodo`, actualizando las invocaciones en `FormPeriodos.cs` (guardar/eliminar) y `AgFormGeneradorHorarios.cs:792` para evitar vaciar catálogos en memoria enlazados en UI.
- **Fase 2081:** Validación de rango válido de año (1..9999) en `CierrePeriodoResolver.ResolverAnioFallback` ante nombres malformados o con año cero (`"0000-01"`), evitando `ArgumentOutOfRangeException` en `new DateTime(0, 12, 15)`. Cobertura en `CierrePeriodoResolverTests.cs`.
- **Fase 2082:** Eliminación de la variable muerta `fechaLimite` en `FormSesionesAcademicas.cs:681` tras la sustitución de la validación del 15/12 en F2076, documentando que el bucle de generación se rige por horas acumuladas y safety guard.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2079 | `CierrePeriodoResolver.cs`, `ICierrePeriodoResolver.cs`, `ResultadoCierrePeriodo.cs`, `CierrePeriodoResolverTests.cs` | Implementado |
| 2080 | `CatalogCacheService.cs`, `FormPeriodos.cs`, `AgFormGeneradorHorarios.cs` | Implementado |
| 2081 | `CierrePeriodoResolver.cs`, `CierrePeriodoResolverTests.cs` | Implementado |
| 2082 | `FormSesionesAcademicas.cs` | Implementado |

---

## Ronda 472: Diagnóstico forense del bloqueo P08 y trazabilidad DEBUG del consentimiento, tope estructural y vacantes (Fases 2083–2089)

### Resumen de la Ronda
- **Fase 2083:** En `AgConsentimientoVacantes.cs`, creación de `FormatearDetalle` y `OrdenSemestre` con orden numérico arábigo estándar (1..12) vía `AgFormatoAsignatura.SemestreArabigo`. Integración en `AgFormGeneradorHorarios.cs` con el resolver en cadena `ResolverEtiquetaSedeConsentimiento(idSede)` (nombre → subunidad → subunidad de contexto → sede {id} / sede N/D). Cobertura con 3 tests en `AgConsentimientoVacantesTests.cs`.
- **Fase 2084:** Instrumentación DEBUG `[AG-CONSENT]` en `AgConsentimientoVacantes.Calcular` (decisión de wildcard/parcial/ninguna, cohortes demandadas, pool de activos y detección de aviso cuando `IdSede == 0`) y en `AgFormGeneradorHorarios.AsegurarConsentimientoVacantes` (universo antes y después del diálogo).
- **Fase 2085:** Instrumentación DEBUG `[AG-TOPE]` en pre-vuelo de `AgValidadorVentana.cs:485-516` (maxConcurrentes, docs, déficit y estado de consentimiento por sede), en fallback de `CalcularDeficitEstructural` (cálculo por asignatura y total) y en `AgGeneradorHorarios.cs:106-109` (fuente del tope estructural).
- **Fase 2086:** Desglose forense del campeón infactible en `AgGeneradorHorarios.cs` bajo `LoggerHelper.ModoVerboso`: `[AG-HARD]` (reglas duras y violaciones individuales), `[AG-VAC]` (conteo de consentidas vs liberadas por AG), `[AG-SNAPSHOT]` (detalles del gen vacante o en conflicto con tercero del pool) y `[AG-PARADA]` (motivo de parada, estancamiento e historial de últimas 10 generaciones).
- **Fase 2087:** Trazabilidad causal `[AG-VACACION]` de liberaciones en `AgAsignadorTercero.DiagnosticarMotivoVacante`, `AgCoherenciaGrupo.Sellar` (motivo modal vs fallback) y en la rama de mutación de `AgOperadorMutacion.cs` (`SinAlternoDisponible` vs `CalendarioNoVerificable`).
- **Fase 2088:** En `AgDiagnosticoInfactibilidad.Describir`, reporte explícito de `ReglasDuras=[...]` y separación de vacantes `CONSENTIDAS` vs `LIBERADAS_POR_AG`. Inserción de reglas duras como segunda línea en `ResumenCorto`. Emisión de advertencia `[AG-GATE]` en `AgFormGeneradorHorarios.cs:3256-3266` previo al modal de infactibilidad. Cobertura en `AgDiagnosticoInfactibilidadTests.cs`.
- **Fase 2089:** Compilación limpia de `ConversorPDF.XLL` y ejecución exitosa de la suite de pruebas completa en `ConversorPDF.Tests` (1801 pruebas superadas, 0 errores). Cierre documental y sincronización de registro.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2083 | `AgConsentimientoVacantes.cs`, `AgFormGeneradorHorarios.cs`, `AgConsentimientoVacantesTests.cs` | Implementado |
| 2084 | `AgConsentimientoVacantes.cs`, `AgFormGeneradorHorarios.cs` | Implementado |
| 2085 | `AgValidadorVentana.cs`, `AgGeneradorHorarios.cs` | Implementado |
| 2086 | `AgGeneradorHorarios.cs` | Implementado |
| 2087 | `AgAsignadorTercero.cs`, `AgCoherenciaGrupo.cs`, `AgOperadorMutacion.cs` | Implementado |
| 2088 | `AgDiagnosticoInfactibilidad.cs`, `AgFormGeneradorHorarios.cs`, `AgDiagnosticoInfactibilidadTests.cs` | Implementado |
| 2089 | `REGISTRY.md`, `implementation_plan.md`, `walkthrough.md` | Implementado |

---

## Ronda 473: Diagnóstico forense con trazabilidad correcta, pero causa raíz de R15 sin corregir (unidades y consentimiento sede 0) (Fases 2090–2095)

### Resumen de la Ronda
- **Fase 2090:** Unificación de la unidad de conteo en R15 (`AgValidadorReglas.cs:725-745`) a cohortes (`cohortesLiberadas`, clave `{asig}|{sede}|{sem}|{grp}`) preservando el número total de sesiones para telemetría. Unificación del cálculo de déficit estructural en `AgValidadorVentana.cs` con el helper canónico `CalcularDeficitPorAsignatura(AgContexto c, int idAsignatura, out int cohortes, out int docs)`, consumido tanto por el pre-vuelo de factibilidad como por el fallback.
- **Fase 2091:** Corrección de fuga de consentimiento ante sedes no resueltas o nulas (`IdSede <= 0`): registro de clave compuesta `{asig}:0` en `AgConsentimientoVacantes.Calcular` y soporte de `Math.Max(0, idSede)` en `AgContexto.EstaConsentida`.
- **Fase 2092:** Extracción canónica de `AgOrigenVacante` (Domain/Enums) y `AgClasificadorVacante` (Application/Services/Genetica). Sustitución de clasificadores ad-hoc duplicados en `AgValidadorReglas.cs`, `AgGeneradorHorarios.cs` (`[AG-VAC]`) y `AgDiagnosticoInfactibilidad.cs` (`cLiberadas`).
- **Fase 2093:** Corrección del cálculo de motivo en `AgConsentimientoVacantes.FormatearDetalle` para obtener la moda estadística (`GroupBy(x => x.Motivo).OrderByDescending(m => m.Count()).ThenBy(m => m.Key)`) en lugar de `FirstOrDefault()`.
- **Fase 2094:** Eliminación del fallback ambiguo `_contextoActual.Subunidad` en `AgFormGeneradorHorarios.ResolverEtiquetaSedeConsentimiento` (prioriza `_contextoActual.SedesPorId`, luego `_sedes`, y finalmente `$"sede {idSede}"`). Ordenación determinista de sedes por clave en `FormatearDetalle`.
- **Fase 2095:** Cobertura de pruebas unitarias (`AgValidadorReglasTests.cs`, `AgConsentimientoVacantesTests.cs`, `AgValidadorVentanaTests.cs`, `AgMotivoVacanteTests.cs`). Sincronización documental en `REGISTRY.md` (R-147), `implementation_plan.md` y `walkthrough.md`.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2090 | `AgValidadorReglas.cs`, `AgValidadorVentana.cs` | Implementado |
| 2091 | `AgConsentimientoVacantes.cs`, `AgContexto.cs` | Implementado |
| 2092 | `AgOrigenVacante.cs`, `AgClasificadorVacante.cs`, `AgValidadorReglas.cs`, `AgGeneradorHorarios.cs`, `AgDiagnosticoInfactibilidad.cs`, `ConversorPDF.XLL.csproj` | Implementado |
| 2093 | `AgConsentimientoVacantes.cs` | Implementado |
| 2094 | `AgFormGeneradorHorarios.cs`, `AgConsentimientoVacantes.cs` | Implementado |
| 2095 | `AgValidadorReglasTests.cs`, `AgConsentimientoVacantesTests.cs`, `AgValidadorVentanaTests.cs`, `AgMotivoVacanteTests.cs`, `REGISTRY.md`, `implementation_plan.md`, `walkthrough.md` | Implementado |

---

## Ronda 474: R15 en cohortes aplicado, pero el déficit perdió la concurrencia por jornada (regresión F2031/F2036) y persisten fugas de clasificación (Fases 2096–2102)

### Resumen de la Ronda
- **Fase 2096:** Restitución de la concurrencia por jornada en el cálculo de déficit estructural de `AgValidadorVentana.CalcularDeficitPorAsignatura` (`maxConcurrentes = demanda.GroupBy(d => d.Jornada)...DefaultIfEmpty(0).Max()`), retornando `Math.Max(0, maxConcurrentes - docs)`. Corrección del test a `DeficitDocentePorAsignatura[10].Should().Be(1)` y adición de test con dos jornadas y un docente con déficit 0.
- **Fase 2097:** Creación del helper canónico `AgClasificadorVacante.ClaveCohorte(AgGen g)` con discriminante por aula y gen (`{asig}|{sede}|aula:{aula}|gen:{id}`) cuando faltan `Semestre` o `Grupo`, evitando que genes incompletos colapsen en una única cohorte en `AgValidadorReglas.cs` (R15) y `AgGeneradorHorarios.cs` (`[AG-VAC]`).
- **Fase 2098:** Adición de `AgOrigenVacante.Indeterminado` para tratar motivos vacíos, nulos o desconocidos (`AgMotivoVacante.Ninguno`). Ignorado explícito en `AgValidadorReglas.cs` para evitar falsos positivos de R15 y R16 ante genes mutados antes del sellado.
- **Fase 2099:** Corrección del detalle de vacantes en `AgDiagnosticoInfactibilidad.Describir`, incorporando el desglose de `ESTRUCTURALES_NC`, `NO_DOCENTE` e `INDETERMINADAS`, garantizando la partición exacta `CONSENTIDAS + LIBERADAS_POR_AG + ESTRUCTURALES_NC + NO_DOCENTE + INDETERMINADAS == Vacantes`.
- **Fase 2100:** Optimización algorítmica de O(G×A) a O(N) mediante la precomputación del diccionario base en `AgValidadorVentana.ConstruirDeficitBase(AgContexto c)`, consumido en un solo pase por `ValidarFactibilidad`, `CalcularDeficitEstructural` y `CalcularDeficitPorAsignatura`.
- **Fase 2101:** Salvaguarda en `AgConsentimientoVacantes.FormatearDetalle` ante `resolverSede` retornando vacío o espacios, cayendo a `$"sede {id}"` o `"sede N/D"`. Renombrado de test obsoleto en `AgValidadorVentanaTests.cs`.
- **Fase 2102:** Cobertura de pruebas unitarias (`AgValidadorVentanaTests.cs`, `AgMotivoVacanteTests.cs`, `AgValidadorReglasTests.cs`, `AgDiagnosticoInfactibilidadTests.cs`, `AgConsentimientoVacantesTests.cs`) y sincronización documental de cierre en `REGISTRY.md` (R-148), `implementation_plan.md` y `walkthrough.md`.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2096 | `AgValidadorVentana.cs`, `AgValidadorVentanaTests.cs` | Implementado |
| 2097 | `AgClasificadorVacante.cs`, `AgValidadorReglas.cs`, `AgGeneradorHorarios.cs`, `AgMotivoVacanteTests.cs` | Implementado |
| 2098 | `AgOrigenVacante.cs`, `AgClasificadorVacante.cs`, `AgValidadorReglas.cs`, `AgMotivoVacanteTests.cs`, `AgValidadorReglasTests.cs` | Implementado |
| 2099 | `AgDiagnosticoInfactibilidad.cs`, `AgDiagnosticoInfactibilidadTests.cs` | Implementado |
| 2100 | `AgValidadorVentana.cs` | Implementado |
| 2101 | `AgConsentimientoVacantes.cs`, `AgValidadorVentanaTests.cs`, `AgConsentimientoVacantesTests.cs` | Implementado |
| 2102 | `REGISTRY.md`, `implementation_plan.md`, `walkthrough.md` | Implementado |

---

## Ronda 475: Error de compilación en la suite y R15 aún compara agregados distintos (max por jornada vs total de cohortes) (Fases 2103–2107)

### Resumen de la Ronda
- **Fase 2103:** Eliminación de la llave de cierre redundante en `AgValidadorReglasTests.cs:1885` originada en la fusión de F2102, resolviendo el error `CS1022`/`CS1513` que impedía la compilación del proyecto de pruebas.
- **Fase 2104:** Alineación de la métrica de agregación en R15 (`AgValidadorReglas.cs:730-760`): `liberadasPorAsignatura` almacena `(Jornada, Cohorte)` y totaliza la suma de máximos concurrentes por jornada (`liberadasPorAsignatura.Values.Sum(...)`), igualando la unidad del tope estructural `MaxVacantesForzadas`. Cobertura en `AgValidadorReglasTests.cs`.
- **Fase 2105:** En `AgValidadorVentana.cs`, confinamiento de la construcción de `sedesPorAsig` dentro de `if (LoggerHelper.ModoVerboso)` y normalización de `Jornada` a mayúsculas con `Trim().ToUpperInvariant()`, garantizando consistencia insensible a mayúsculas y eliminando allocations O(N) innecesarias en producción.
- **Fase 2106:** Precisión del rotulado de telemetría en `AgGeneradorHorarios.cs:410-440` (`[AG-VAC]`), separando claramente `estructuralesAsignaturas` de `estructuralesSesiones` para erradicar contradicciones con el reporte de violaciones R16.
- **Fase 2107:** Gate de verificación obligatorio: compilación limpia de la solución y ejecución exitosa de la suite completa en `ConversorPDF.Tests` (1824 pruebas superadas, 0 errores). Cierre y sincronización de registro `R-149`.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2103 | `AgValidadorReglasTests.cs` | Implementado |
| 2104 | `AgValidadorReglas.cs`, `AgValidadorReglasTests.cs` | Implementado |
| 2105 | `AgValidadorVentana.cs` | Implementado |
| 2106 | `AgGeneradorHorarios.cs` | Implementado |
| 2107 | `REGISTRY.md`, `implementation_plan.md`, `walkthrough.md`, `ConversorPDF.Tests` (1824 tests) | Implementado |































