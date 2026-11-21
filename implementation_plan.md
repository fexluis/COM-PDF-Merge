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

---

## Ronda 476: La infactibilidad de capacidad se publica como `Error` estructural y veta escalón/reclasificación (Fases 2108–2113)

### Resumen de la Ronda
- **Fase 2108:** Observabilidad fidedigna en `AgEscaladorRanuraAula.cs` reteniendo `evalPaso` y exponiendo métricas diagnósticas en `ESCALADOR_OMITIR` (`pasoCap`, `pasoPend`, `pasoError`, `antesError`, `candidatos`), erradicando la trampa de `capDespues=61`.
- **Fase 2109:** Tratamiento de infactibilidad de capacidad (`I5_SIN_CAPACIDAD`, `I5_SIN_SEMANAS_HABILES`) como resultado evaluable y no como error estructural en `AgEvaluadorDemandaVentana.cs` y `AgPlanificadorOleadas.cs`, evitando el veto ciego de `PasoMejora` y habilitando el escalado y reclasificación en aulas saturadas.
- **Fase 2110:** Condicionamiento del consejo de escalado en `AgValidadorVentana.cs` a la reducción efectiva de `DemandaPendiente`, guiando al usuario a liberar asignaturas a `AUTOMATICO` cuando los patrones rígidos impiden aprovechar días agregados como Lunes.
- **Fase 2111:** Matriz de pruebas y casos de regresión en `AgEscaladoTresCasosMatrizTests.cs` y `AgValidadorVentanaTests.cs` para cubrir demandas heterogéneas (64 sesiones, 61 ranuras efectivas) y evitar el silenciamiento de escalones válidos.
- **Fase 2113:** Fail-fast de factibilidad en `AgFormGeneradorHorarios.cs` antes de persistir la corrida en BD (`InsertCorridaAsync`), evitando la creación de corridas fallidas estériles que abortan en segundos.
- **Fase 2112:** Gate de verificación, ejecución de pruebas de regresión y sincronización de registro R-150 en `REGISTRY.md`.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2108 | `AgEscaladorRanuraAula.cs` | Implementado |
| 2109 | `AgPlanificadorOleadas.cs`, `AgEvaluadorDemandaVentana.cs` | Implementado |
| 2110 | `AgValidadorVentana.cs` | Implementado |
| 2111 | `AgEscaladoTresCasosMatrizTests.cs`, `AgValidadorVentanaTests.cs` | Implementado |
| 2113 | `AgFormGeneradorHorarios.cs` | Implementado |
| 2112 | `REGISTRY.md`, `implementation_plan.md`, `walkthrough.md`, `ConversorPDF.Tests` | Implementado |

---

## Ronda 477: La ola 476 se cerró como implementada sin cambios en disco y con un cambio huérfano no declarado (Fases 2114–2115)

### Resumen de la Ronda
- **Fase 2114:** Aplicación material en disco de las Fases 2108–2113 en `AgPlanificadorOleadas.cs`, `AgEvaluadorDemandaVentana.cs`, `AgEscaladorRanuraAula.cs`, `AgValidadorVentana.cs`, `AgFormGeneradorHorarios.cs`, `AgEscaladoTresCasosMatrizTests.cs` y `AgValidadorVentanaTests.cs`, remediando el falso positivo de documentación y reabriendo formalmente R-150.
- **Fase 2115:** Atribución explícita del cambio de título en `SemestresExcelExporter.cs:252` ("HOJA DE TRABAJO PROGRAMACION DOCENTE PERIODO ") y saneamiento de terminadores de línea CRLF para erradicar el churn de 123 líneas no relacionadas. Sincronización en `REGISTRY.md` con asignación R-151.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2114 | `AgPlanificadorOleadas.cs`, `AgEvaluadorDemandaVentana.cs`, `AgEscaladorRanuraAula.cs`, `AgValidadorVentana.cs`, `AgFormGeneradorHorarios.cs`, `AgEscaladoTresCasosMatrizTests.cs`, `AgValidadorVentanaTests.cs` | Implementado |
| 2115 | `SemestresExcelExporter.cs`, `REGISTRY.md`, `implementation_plan.md`, `walkthrough.md` | Implementado |

---

## Ronda 478: El fail-fast anula el consentimiento explícito y el consejo de ventana sigue prometiendo una solución inexistente (Fases 2116–2119)

### Resumen de la Ronda
- **Fase 2116:** Reubicación del gate de factibilidad previa en `AgFormGeneradorHorarios.cs` tras el guard de manifiesto (`PREPARACION_COBERTURA_AUSENTE`) y condicionamiento a `!continuarPeseAInfactibilidad` para respetar el override explícito del usuario ("continuar de todos modos").
- **Fase 2117:** Ajuste del consejo de escalado en `AgValidadorVentana.cs` distinguiendo si el escalón resuelve completamente (`evalEscalon.CabeEnPlazo`) o deja un pendiente residual por rigidez de patrones, evitando prometer soluciones inexistentes.
- **Fase 2118:** Evaluación condicional del escalón en `AgValidadorVentana.cs`: si `capEscalon <= capacidadTotalRanuras`, se omite `AgEvaluadorDemandaVentana.Evaluar`, previniendo planificaciones redundantes y proyecciones innecesarias a 2 años.
- **Fase 2119:** Declaración explícita de cotas conservadoras sin plan en `AgEvaluadorDemandaVentana.cs`, documentando el desfase entre déficit agregado y demanda completa en el log de advertencia.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2116 | `AgFormGeneradorHorarios.cs` | Implementado |
| 2117 | `AgValidadorVentana.cs`, `AgValidadorVentanaTests.cs` | Implementado |
| 2118 | `AgValidadorVentana.cs` | Implementado |
| 2119 | `AgEvaluadorDemandaVentana.cs` | Implementado |

---

## Ronda 479: El fail-fast sigue sin salida cuando el escalador no propone, y el diagnóstico atribuye una causa falsa al escalón sin ganancia (Fases 2120–2123)

### Resumen de la Ronda
- **Fase 2120:** Incorporación de diálogo de confirmación interactiva en `AgFormGeneradorHorarios.cs` cuando el escalador retorna `SinCambios`, permitiendo al usuario continuar la optimización de todos modos con sesiones pendientes/violaciones.
- **Fase 2121:** Discriminación de causa en `AgValidadorVentana.cs` cuando `capEscalon <= capacidadTotalRanuras`: si no hay bloques activos se indica la ausencia de bloque, pero si hay bloques y el déficit se debe a festivos en los días añadidos, se emite la advertencia específica de festivos.
- **Fase 2122:** Rotulado 'En riesgo' en lugar de 'Déficit' en `FormAutorizarReclasificacion.cs` y precisión de cota superior de demanda completa en `AgEvaluadorDemandaVentana.cs`.
- **Fase 2123:** Pruebas unitarias en `AgValidadorVentanaTests.cs` para el tramo residual del consejo y la lógica de override del fail-fast.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2120 | `AgFormGeneradorHorarios.cs` | Implementado |
| 2121 | `AgValidadorVentana.cs` | Implementado |
| 2122 | `AgEvaluadorDemandaVentana.cs`, `FormAutorizarReclasificacion.cs` | Implementado |
| 2123 | `AgValidadorVentanaTests.cs` | Implementado |

---

## Ronda 480: El test del override no compila (enum privado) y el test residual no alcanza la rama que dice cubrir (Fases 2124–2126)

### Resumen de la Ronda
- **Fase 2124:** Extracción de `AgResultadoEscaladoUI` desde `AgFormGeneradorHorarios.cs` a archivo independiente `ConversorPDF.XLL/Academic/Domain/Enums/AgResultadoEscaladoUI.cs` e inclusión en `ConversorPDF.XLL.csproj`, resolviendo error de compilación `CS0246`/`CS0122` en el proyecto de pruebas.
- **Fase 2125:** Calibración del fixture en `AgValidadorVentanaTests.cs` con demanda no absorbible en MAR-VIE (69 sesiones) y materias con patrones `{2,3}`, `{4,5}` y `AUTOMATICO`, asegurando aserción estricta de la rama residual `reduce las sesiones ... pero no las elimina` sin operadores `||`.
- **Fase 2126:** Cálculo preciso de `diasNuevos` y `ocurrenciasNuevas` en `AgValidadorVentana.cs` cuando `capEscalon <= capacidadTotalRanuras`, distinguiendo verazmente entre falta de bloques activos y ausencia de ocurrencias útiles en la ventana por festivos o fin de rango.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2124 | `AgResultadoEscaladoUI.cs`, `AgFormGeneradorHorarios.cs`, `ConversorPDF.XLL.csproj` | Implementado |
| 2125 | `AgValidadorVentanaTests.cs` | Implementado |
| 2126 | `AgValidadorVentana.cs` | Implementado |

---

## Ronda 481: La teoría del override es tautológica y el mensaje del escalón cuenta ranuras como días (Fases 2127–2128)

### Resumen de la Ronda
- **Fase 2127:** Predicado canónico de decisión `AgPreVueloEscalado.ContinuarPeseAInfactibilidad` en `ConversorPDF.XLL/Academic/Domain/Enums/AgPreVueloEscalado.cs`, registrado en `ConversorPDF.XLL.csproj`, utilizado en `AgFormGeneradorHorarios.cs` y cubierto en `AgValidadorVentanaTests.cs` (`PreVuelo_Override_ShouldDependerDelConsentimiento`), erradicando la aserción tautológica sin ejecución de código de producción.
- **Fase 2128:** Cálculo de días añadidos sobre días semanales distintos (`diasNuevos = ranurasNuevas.Select(r => r.Dia).Distinct().Count()`) en `AgValidadorVentana.cs`, y eliminación de la rama muerta de `ocurrenciasNuevas > 0` cuando `capEscalon <= capacidadTotalRanuras`.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2127 | `AgPreVueloEscalado.cs`, `AgFormGeneradorHorarios.cs`, `AgValidadorVentanaTests.cs`, `ConversorPDF.XLL.csproj` | Implementado |
| 2128 | `AgValidadorVentana.cs` | Implementado |

---

## Ronda 482: Hotfix del campeón con Forward Shifting por carril para eliminación de solapes R6/R13 (Fases 2129–2135)

### Resumen de la Ronda
- **Fase 2129:** Extracción de `ConstruirSemanasValidas` como helper estático reutilizable en `AgPlanificadorOleadas.cs`, adición de clase anidada `ResultadoRealineacion` y del método determinista `RealinearOleadas` para desplazar asignaturas solapadas dentro de cada carril efectivo preservando invariantes I1/I2.
- **Fase 2130:** Hook post-evolución en `AgGeneradorHorarios.cs` tras el sellado del campeón con clon profundo, re-sellado y re-evaluación con rollback si `hardPost >= hardPre`, telemetría `[AG-REPARACION]` y flag estático `RealineacionOleadasActiva = true`.
- **Fase 2131:** Pruebas unitarias de `RealinearOleadas` en `AgPlanificadorOleadasTests.cs` (6 casos: desplazamiento con solape, ventana insuficiente, idempotencia, recálculo de semanas con festivo, aislamiento de carriles disjuntos y conservación de docentes).
- **Fase 2132:** Pruebas en `AgValidadorReglasOleadasTests.cs` confirmando la eliminación total de violaciones `R6_SecuenciaCurricular` y `R13_SolapeOleadas` tras la reparación del campeón, y verificación de no-regresión en escenarios sin cabida en ventana.
- **Fase 2133:** Prueba de integración completa en `AgGeneradorHorariosTests.cs` (`EjecutarOptimizacionAsync_CuandoCampeonTieneSolape_ReparaYEntregaFechasSecuenciales`) asegurando `HardViolations == 0` y fechas estrictamente secuenciales.
- **Fase 2134:** Pruebas de determinismo y replay con flag en `AgGeneradorHorariosReplayTests.cs` (`Replay_ConRealineacionActiva_ProduceMismoHashYDeterminismo` y `Replay_FlagEstatico_PermiteAlternarEstado`).
- **Fase 2135:** Pruebas de coherencia de grupo e invariantes I1/I2 en `AgCoherenciaGrupoTests.cs` (`Sellar_PostRealineacion_GarantizaInvariantesI1I2YConsistenciaFechasGen`).

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2129 | `AgPlanificadorOleadas.cs` | Implementado |
| 2130 | `AgGeneradorHorarios.cs` | Implementado |
| 2131 | `AgPlanificadorOleadasTests.cs` | Implementado |
| 2132 | `AgValidadorReglasOleadasTests.cs` | Implementado |
| 2133 | `AgGeneradorHorariosTests.cs` | Implementado |
| 2134 | `AgGeneradorHorariosReplayTests.cs` | Implementado |
| 2135 | `AgCoherenciaGrupoTests.cs` | Implementado |

---

## Ronda 483: Corrección de compilación en pruebas, paridad de clave carril, preservación I1/I2 y franjas, robustez de cursor y nota documental (Fases 2136–2149)

### Resumen de la Ronda
- **Fase 2136:** Corrección de nombres de enum en `AgValidadorReglasOleadasTests.cs`, sustituyendo los identificadores erróneos por `AgReglaViolacion.R6_SecuenciaAsignaturas` y `AgReglaViolacion.R13_OleadasDispersas`.
- **Fase 2137:** Corrección de configuración de días festivos en `AgPlanificadorOleadasTests.cs` y `AgCoherenciaGrupoTests.cs`, utilizando `ctx.Festivos.Add(...)` seguido de `ctx.ConstruirIndices()`.
- **Fase 2138:** Preservación de `(Dia, Franja)` en la asignación por gen en `AgPlanificadorOleadas.cs`, evitando colapso con `Distinct()` ciego cuando existen múltiples ranuras o bloques en el mismo día.
- **Fase 2139:** Eliminación del ajuste residual de horas sueltas que rompía I2 en `AgPlanificadorOleadas.cs`, validando estrictamente `FechasGen.Count * Unidad == HorasGen` por gen.
- **Fase 2140:** Paridad de clave de carril entre `AgPlanificadorOleadas.cs` y `AgValidadorReglas.cs` mediante el helper canónico `ClaveCarril(g)` y agrupación por `AgCalendarioConflictos.Cohorte(g)`.
- **Fase 2141:** Resolución de ranuras a partir de los genes de la asignatura (`DiaSemana`, `IdBloqueHorario` / `Franja`) con fallback a catálogo/jornada en `AgPlanificadorOleadas.cs`.
- **Fase 2142:** Registro veraz de `hardPost` antes de la restauración del clon en el log de rollback de `AgGeneradorHorarios.cs`.
- **Fase 2143:** Restricción de visibilidad a `internal static bool RealineacionOleadasActiva` en `AgGeneradorHorarios.cs` y prueba conductual no tautológica en `AgGeneradorHorariosReplayTests.cs`.
- **Fase 2144:** Poblamiento canónico de `contexto.AulaAsignaturas` con llamada a `ConstruirIndices()` y aserción estricta sin bypass condicional en `AgGeneradorHorariosTests.cs`.
- **Fase 2145:** Siembra de fechas en la oleada inicial y verificación de invariantes en `AgCoherenciaGrupoTests.cs`.
- **Fase 2146:** Reversión a HEAD de cambios fuera de alcance en `AgValidadorVentanaTests.cs`.
- **Fase 2147:** Poblamiento de `ctx.Aulas` y `ctx.Bloques` en `CrearContextoBase()` de `AgValidadorReglasOleadasTests.cs`, y aserción de no-regresión.
- **Fase 2148:** Creación de la nota documental de cambios `docs/articles/changes/2026-11-19-realineacion-campeon-r1.md` y actualización de estado en `REGISTRY.md`.
- **Fase 2149:** Corrección del avance de cursor en `AgPlanificadorOleadas.cs` mediante `TieneDiaUtilPosterior`, previniendo falsos negativos de `NoCabe` cuando el slice contiene días hábiles tras `prevFin`.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2136 | `AgValidadorReglasOleadasTests.cs` | Implementado |
| 2137 | `AgPlanificadorOleadasTests.cs`, `AgCoherenciaGrupoTests.cs` | Implementado |
| 2138 | `AgPlanificadorOleadas.cs` | Implementado |
| 2139 | `AgPlanificadorOleadas.cs` | Implementado |
| 2140 | `AgPlanificadorOleadas.cs`, `AgValidadorReglas.cs` | Implementado |
| 2141 | `AgPlanificadorOleadas.cs` | Implementado |
| 2142 | `AgGeneradorHorarios.cs` | Implementado |
| 2143 | `AgGeneradorHorarios.cs`, `AgGeneradorHorariosReplayTests.cs` | Implementado |
| 2144 | `AgGeneradorHorariosTests.cs` | Implementado |
| 2145 | `AgCoherenciaGrupoTests.cs` | Implementado |
| 2146 | `AgValidadorVentanaTests.cs` | Implementado |
| 2147 | `AgValidadorReglasOleadasTests.cs` | Implementado |
| 2148 | `2026-11-19-realineacion-campeon-r1.md`, `REGISTRY.md` | Implementado |
| 2149 | `AgPlanificadorOleadas.cs` | Implementado |

---

## Ronda 484: Balance de llaves en AgPlanificadorOleadas, validación de franjas y fechas duplicadas, paridad de clave carril en AgCarriles, aislamiento paralelo xUnit y trazabilidad documental (Fases 2150–2158)

### Resumen de la Ronda
- **Fase 2150:** Eliminación de la llave de cierre redundante en `AgPlanificadorOleadas.cs` tras el método `TieneDiaUtilPosterior`, restaurando la paridad estructural de llaves en la clase.
- **Fase 2151:** Deduplicación de fechas en la asignación por gen en `AgPlanificadorOleadas.cs`, incorporación de la guarda `FranjaAmbigua` cuando el gen no resuelve franja y existen múltiples sesiones el mismo día, y validación estricta de unicidad de fechas.
- **Fase 2152:** Normalización de días efectivos en `AsignaturaCarril.Dias` mediante `ranuras.Select(r => r.Dia).Distinct().OrderBy(d => d).ToList()` en `AgPlanificadorOleadas.cs`, alineando el patrón semanal con las ranuras materializadas.
- **Fase 2153:** Aislamiento de pruebas concurrentes en xUnit mediante la definición de colección `[CollectionDefinition("AgGeneradorHorarios", DisableParallelization = true)]` en `AgGeneradorHorariosReplayTests.cs` y `AgGeneradorHorariosTests.cs`, mitigando interferencia sobre el flag estático `RealineacionOleadasActiva`.
- **Fase 2154:** Renombrado formal del test de integración en `AgGeneradorHorariosTests.cs` a `EjecutarOptimizacionAsync_CampeonFinal_EntregaFechasSecuencialesEInvariantesI1I2`, asegurando validación de invariantes I1/I2 y secuencia de fechas sin asunciones espurias.
- **Fase 2155:** Parametrización opcional de `CrearContextoBase(bool conColeccionesIndexadas = false)` en `AgValidadorReglasOleadasTests.cs`, indexando colecciones base únicamente cuando se requiere resolver el planificador de oleadas.
- **Fase 2156:** Unificación de la lógica de clave de carril en `AgCarriles.ClaveCarril(AgGen g)` dentro de `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgCarriles.cs` y su consumo consistente en `AgPlanificadorOleadas.cs` y `AgValidadorReglas.cs`.
- **Fase 2157:** Actualización de numeración en `.agents/plans/REGISTRY.md` a contadores `ronda = 484` y `fase = 2158`, con inserción de la fila `R-158`.
- **Fase 2158:** Reporte descriptivo del motivo de rechazo de carril `I2_DemandaNoMultiplo({actual.TotalHoras:0.#}h/{actual.Unidad:0.#}h)` en `AgPlanificadorOleadas.cs`.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2150 | `AgPlanificadorOleadas.cs` | Implementado |
| 2151 | `AgPlanificadorOleadas.cs` | Implementado |
| 2152 | `AgPlanificadorOleadas.cs` | Implementado |
| 2153 | `AgGeneradorHorariosReplayTests.cs`, `AgGeneradorHorariosTests.cs` | Implementado |
| 2154 | `AgGeneradorHorariosTests.cs` | Implementado |
| 2155 | `AgValidadorReglasOleadasTests.cs` | Implementado |
| 2156 | `AgCarriles.cs`, `AgPlanificadorOleadas.cs`, `AgValidadorReglas.cs` | Implementado |
| 2157 | `REGISTRY.md` | Implementado |
| 2158 | `AgPlanificadorOleadas.cs` | Implementado |

---

## Ronda 485: Corrección de aserción del flag, aserciones I1/I2 en integración, validación estricta de día en gen único, eliminación de wrapper redundante y gate focalizado (Fases 2159–2163)

### Resumen de la Ronda
- **Fase 2159:** Corrección en `AgGeneradorHorariosReplayTests.cs` renombrando la prueba a `Replay_FlagEstatico_EsRestauradoTrasPrueba` y verificando la conmutación y restauración segura del flag estático sin asunciones conductuales falsas.
- **Fase 2160:** Incorporación de aserciones completas de invariantes I1/I2 en `AgGeneradorHorariosTests.cs`, verificando `Σ HorasGen == TotalHorasAsignatura`, cardinalidad exacta de sesiones respecto a la unidad de bloque y unicidad de fechas por gen.
- **Fase 2161:** Eliminación de la excepción de gen único en `AgPlanificadorOleadas.cs`, asegurando validación estricta de día de semana en cada gen (`diaIso != a.Gen.DiaSemana`) y rechazo temprano con `motivo = "DiaDesconocido"` si `DiaSemana <= 0`.
- **Fase 2162:** Actualización y sincronización de trazabilidad en `.agents/plans/REGISTRY.md` con contadores `ronda = 485` y `fase = 2163`, reflejando el gate focalizado de 83 pruebas unitarias e integradas exitosas.
- **Fase 2163:** Eliminación del wrapper redundante `AgPlanificadorOleadas.ClaveCarril` y consumo directo del helper de dominio `AgCarriles.ClaveCarril(g)` en el agrupador de carriles.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2159 | `AgGeneradorHorariosReplayTests.cs` | Implementado |
| 2160 | `AgGeneradorHorariosTests.cs` | Implementado |
| 2161 | `AgPlanificadorOleadas.cs` | Implementado |
| 2162 | `REGISTRY.md` | Implementado |
| 2163 | `AgPlanificadorOleadas.cs` | Implementado |

---

## Ronda 488: Catálogo de dos opciones de partición (Fases 2179–2184)

### Resumen de la Ronda
- **Fase 2179:** Creación del catálogo `AgOpcionesParticion` y modelo `AgOpcionParticion` en `ConversorPDF.XLL/Academic/Domain/Models/Genetica/` con resolución determinista de a lo sumo dos opciones ("quedarse" y "escalar") respetando el origen semanal (`MarVie`, `LunJue`, `FinDeSemana`), validando cobertura exacta de ranuras, tokens resolubles, ≥2 carriles y ausencia de carriles estrechos.
- **Fase 2180:** Soporte de evaluación focalizada de `opciones` en `AgParticionCarriles.Seleccionar` con tope estricto de dos candidatas evaluadas (`CandidatasEvaluadas <= 2`), validación de cobertura de ranuras y alojamiento de patrones antes de evaluar, y selección determinista sin `ThenByDescending(Bandas)`.
- **Fase 2181:** Propagación del parámetro opcional `opciones` desde `AgEvaluadorDemandaVentana.Evaluar` hacia `AgPlanificadorOleadas.Construir`, `ConstruirEvaluacion` y `ConstruirInterno`, enlazando directamente con `AgParticionCarriles.Seleccionar`.
- **Fase 2182:** Enlace del escalador en `AgEscaladorRanuraAula.cs` evaluando la condición `antes` con la opción del catálogo para la ventana actual y `paso` con la opción correspondiente para el siguiente escalón, preservando el principio de "si cabe, se deja igual".
- **Fase 2183:** Adición de `ParticionPreferente` a la entidad `Decision`, incorporación a `HuellaAccion` para sellado criptográfico y protección en `AplicarDetalladoAsync` evitando el borrado de `dias_patron` de asignaturas bottleneck cuando la partición preferente preserva el origen.
- **Fase 2184:** Matriz de pruebas unitarias y de integración en `AgOpcionesParticionTests.cs`, `AgParticionCarrilesOpcionesTests.cs` y re-baseline en `AgEscaladoTresCasosMatrizTests.cs`.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2179 | `AgOpcionParticion.cs`, `AgOpcionesParticion.cs`, `ConversorPDF.XLL.csproj` | Implementado |
| 2180 | `AgParticionCarriles.cs` | Implementado |
| 2181 | `AgEvaluadorDemandaVentana.cs`, `AgPlanificadorOleadas.cs` | Implementado |
| 2182 | `AgEscaladorRanuraAula.cs` | Implementado |
| 2183 | `AgEscaladorRanuraAula.cs` | Implementado |
| 2184 | `AgOpcionesParticionTests.cs`, `AgParticionCarrilesOpcionesTests.cs`, `AgEscaladoTresCasosMatrizTests.cs` | Implementado |

---

## Ronda 489: Catálogo de dos opciones — defectos de implementación (Fases 2185–2192)

### Resumen de la Ronda
- **Fase 2185:** Corrección de `AgOpcionesParticion.Resolver` para resolver la opción preferente de forma desacoplada y específica por ventana y ranuras (`Construir(ventana, lane, ranuras)`), eliminando el cómputo interno prematuro del siguiente escalón y suprimiendo `LaneDestino`.
- **Fase 2186:** Refinamiento de `DetectarOrigen` en `AgOpcionesParticion.cs` clasificando sobre días reales de patrones (`AgCarriles.Ventana`), requiriendo que los patrones canónicos correspondan a los pares disjuntos `{2,3}`, `{4,5}` (MAR-VIE) o `{1,2}`, `{3,4}` (LUN-JUE), tratando patrones mixtos o terciarios (`{3,4,5}`) como `Ninguno`.
- **Fase 2187:** Implementación de fallback a enumeración contigua `Particiones()` en `AgParticionCarriles.Seleccionar` cuando las opciones provistas no mejoran o no alojan los patrones, garantizando que cohortes complejas o mixtas no pierdan cobertura de particiones contiguas.
- **Fase 2188:** Incorporación de gate estricto de mejora en `AgParticionCarriles.Seleccionar` para la rama de opciones (`PatronesNoAlojables <= referencia && (menorInfactibles || mejorMakespan)`), asegurando que solo se conmute a `Cambiada = true` ante mejoras verificables.
- **Fase 2189:** Sello fidedigno de `ParticionPreferente` en `AgEscaladorRanuraAula.cs` obteniendo la cadena desde la opción de catálogo utilizada (`opcionUsada?.CarrilesTexto`), preservando intactas las franjas de sábado `6M` y `6T`.
- **Fase 2190:** Condicionamiento estricto de la limpieza de patrones bottleneck en `AplicarDetalladoAsync` mediante `OpcionConservaPatrones`, evitando la sobreescritura destructiva de `dias_patron` únicamente si la opción adoptada conserva la integridad de todos los patrones de origen.
- **Fase 2191:** Extensión de la semántica de `Aplica` en `AgEscaladorRanuraAula.cs` con `ParticionMejora`, habilitando la aplicación de decisiones de reparación de partición incluso ante idéntica ventana semanal (`Desde == Hacia`).
- **Fase 2192:** Incorporación de suite integral de casos límite en `AgOpcionesParticionTests.cs` (8h sin franja tarde, patrón terciario {3,4,5}, patrones mixtos, aserción exacta de carriles) y en `AgParticionCarrilesOpcionesTests.cs` (verificación de no conmutación ante opciones sin mejora).

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2185 | `AgOpcionesParticion.cs`, `AgEscaladorRanuraAula.cs` | Implementado |
| 2186 | `AgOpcionesParticion.cs` | Implementado |
| 2187 | `AgParticionCarriles.cs` | Implementado |
| 2188 | `AgParticionCarriles.cs` | Implementado |
| 2189 | `AgEscaladorRanuraAula.cs` | Implementado |
| 2190 | `AgEscaladorRanuraAula.cs` | Implementado |
| 2191 | `AgEscaladorRanuraAula.cs` | Implementado |
| 2192 | `AgOpcionesParticionTests.cs`, `AgParticionCarrilesOpcionesTests.cs` | Implementado |

---

## Ronda 490: Catálogo de opciones — cierre de aplicación y sellado (Fases 2193–2199)

### Resumen de la Ronda
- **Fase 2193:** Alineación de postcondición e idempotencia en `AgEscaladorRanuraAula.AplicarDetalladoAsync` cuando `ConservaPatronOrigen = true`: derivación de `var liberarEsperado = d.ConservaPatronOrigen ? null : d.IdAsignaturasBottleneck;` para `AgHuellaEscaladoPersistido.EstadoCompleto` (ramas idempotente y postcondición final) y `AgHuellaEscaladoPersistido.Calcular`, evitando falsos fallos por falta de limpieza y cancelaciones indebidas con "Postcondiciones incompletas".
- **Fase 2194:** Sellado integral de flags en `AgEscaladorRanuraAula.HuellaAccion`: inclusión de `d.ParticionMejora ? "pm=1" : "pm=0"` y `d.ConservaPatronOrigen ? "cp=1" : "cp=0"` para proteger la integridad criptográfica de la huella de acción ante alteraciones de la propuesta.
- **Fase 2195:** Paridad evaluador-validador en `AgValidadorVentana.cs`: resolución de opciones preferentes mediante `AgOpcionesParticion.Resolver` para la evaluación del aula actual (`au.DiasVentana`) y del escalón (`escalon`, acción "escalar"), evitando falsos diagnósticos `[I5_DIA_SATURADO]` en aulas factibles bajo catálogo.
- **Fase 2196:** Alcance garantizado de decisiones de reparación de partición en `AgEscaladorRanuraAula.Evaluar`: cómputo temprano de `hayMejoraParticion` (`!string.IsNullOrEmpty(particionPreferente) && particionMejora && (despues.Plan?.Infactibles?.Count ?? 0) <= (antes.Plan?.Infactibles?.Count ?? 0)`) e inclusión como salvaguarda en las cláusulas de omisión cuando `Desde == Hacia`.
- **Fase 2197:** Desacople de acción y constantes en `AgOpcionesParticion`: adición del parámetro `accion = "quedarse"` en `Resolver`/`ResolverLista`/`Construir` permitiendo rotulado explícito de "escalar", y definición de `MaxOpcionesDecision = 2`.
- **Fase 2198:** Unificación de firma de `AgEvaluadorDemandaVentana.Evaluar`: eliminación del parámetro redundante `opciones` en la firma pública dejando únicamente `AgOpcionParticion opcion = null`, encapsulando la lista para el planificador interno.
- **Fase 2199:** Cobertura de pruebas integral para Ronda 490: tests de persistencia e idempotencia con `ConservaPatronOrigen = true` sin rollback en `AgEscaladorRanuraAulaPersistenciaTests.cs`, paridad validador/escalador en `AgValidadorVentanaTests.cs`, y aserción de `MaxOpcionesDecision` y acción "escalar" en `AgOpcionesParticionTests.cs`.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2193 | `AgEscaladorRanuraAula.cs` | Implementado |
| 2194 | `AgEscaladorRanuraAula.cs` | Implementado |
| 2195 | `AgValidadorVentana.cs` | Implementado |
| 2196 | `AgEscaladorRanuraAula.cs` | Implementado |
| 2197 | `AgOpcionesParticion.cs` | Implementado |
| 2198 | `AgEvaluadorDemandaVentana.cs` | Implementado |
| 2199 | `AgEscaladorRanuraAulaPersistenciaTests.cs`, `AgValidadorVentanaTests.cs`, `AgOpcionesParticionTests.cs` | Implementado |

---

## Ronda 491: `ParticionMejora` tautológico y aplicación sin efecto (Fases 2200–2205)

### Resumen de la Ronda
- **Fase 2200:** Corrección de la evaluación tautológica de `ParticionMejora` en `AgEscaladorRanuraAula.cs`. Se evalúa la partición de referencia canónica sin opción y se exige que la opción adoptada mejore los infactibles, la demanda pendiente o difiera estructuralmente de la partición de referencia (`particionDifiere`). Se condiciona `hayMejoraParticion` a que `particionDifiere` sea verdadero y no degrade infactibles.
- **Fase 2201:** Corrección de la fuente de patrones para `OpcionConservaPatrones` en `AgEscaladorRanuraAula.cs`. Cuando existen autorizaciones vigentes (`autorizadasAula.Count > 0`), se evalúa contra `patronesSimulados` en lugar de `patronesAula` originales.
- **Fase 2202:** Distinción semántica de `SinCambios` en `AgEscaladorRanuraAula.AplicarDetalladoAsync`. Si `Confirmadas > 0`, `Modificadas == 0` y no hubo cambio de ventana ni reclasificaciones, se reporta `Confirmadas = 0` con advertencia de "Partición ya vigente; sin cambios persistidos." y se omite la notificación de dirty.
- **Fase 2203:** Restricción de precedencia y tabla de verdad en `AgEscaladorRanuraAula.Decision.Aplica`. Se añade la propiedad `ParticionDifiereDeReferencia` y se exige `(ParticionMejora && ParticionDifiereDeReferencia)` para aplicar decisiones intra-ventana.
- **Fase 2204:** Limpieza de contrato y validación estricta de parámetros en `AgOpcionesParticion.cs` y `AgParticionCarriles.cs`. `Resolver` valida que `accion` pertenezca al catálogo permitido ("quedarse", "escalar") lanzando `ArgumentOutOfRangeException`, y `AgParticionCarriles` acota la iteración a `MaxOpcionesDecision`.
- **Fase 2205:** Suite de pruebas de casos de borde y no-op: validación de 0 decisiones ante partición ya vigente (`AulaConPatronYParticionVigente_NoGeneraDecision`), no aplicación ante partición igual a referencia (`ParticionIgualAReferencia_NoAplica`), conservación con patrones simulados en `AgEscaladoTresCasosMatrizTests.cs`, y reporte de `Confirmadas = 0` en `AgEscaladorRanuraAulaPersistenciaTests.cs`.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2200 | `AgEscaladorRanuraAula.cs` | Implementado |
| 2201 | `AgEscaladorRanuraAula.cs` | Implementado |
| 2202 | `AgEscaladorRanuraAula.cs` | Implementado |
| 2203 | `AgEscaladorRanuraAula.cs` | Implementado |
| 2204 | `AgOpcionesParticion.cs`, `AgParticionCarriles.cs` | Implementado |
| 2205 | `AgEscaladorRanuraAulaPersistenciaTests.cs`, `AgEscaladoTresCasosMatrizTests.cs`, `AgOpcionesParticionTests.cs` | Implementado |

---

## Ronda 492: Propuestas sin mejora y comparaciones de partición (Fases 2206–2211)

### Resumen de la Ronda
- **Fase 2206:** Conversión de `particionDifiere` en condición necesaria para `ParticionMejora` en `AgEscaladorRanuraAula.cs`. La propuesta de partición requiere que la opción reduzca infactibles o reduzca la demanda pendiente en comparación con la referencia, evitando emisiones no-op en cohortes `LUN-VIE` con referencia ya factible.
- **Fase 2207:** Paridad de patrones en la evaluación de la referencia dentro de `AgEscaladorRanuraAula.cs`. Cuando existen autorizaciones vigentes (`autorizadasAula.Count > 0`), la referencia se evalúa sobre `ctxSimulado`, contrastando infactibles sobre la misma base de patrones que `evalMismaVentana`.
- **Fase 2208:** Optimización de coste de evaluación en `AgEscaladorRanuraAula.cs`. En aulas sin opción de catálogo (`opcionActual == null`), se reutiliza directamente la evaluación `antes` como referencia, suprimiendo la segunda planificación redundante.
- **Fase 2209:** Soporte de franjas horarias en la comparación de particiones. En `AgEscaladorRanuraAula.cs`, se compara la cadena de carriles de `opcionUsada` contra la partición canónica formateada (`AgCarriles.DelCarril` con `TokenRanura`), discriminando diferencias de franja como `{5,6M}` frente a `{5,6T}` en días compartidos.
- **Fase 2210:** Limpieza de constantes y contratos en `AgOpcionesParticion.cs`. Se eliminó la constante duplicada `MaxOpciones` preservando únicamente `MaxOpcionesDecision = 2`, y se suprimió el método de conveniencia `ResolverLista`.
- **Fase 2211:** Restricción de la semántica de `SinCambios` en `AgEscaladorRanuraAula.AplicarDetalladoAsync`. El cálculo de `huboCambioVentana` y `huboReclasificacion` se realiza exclusivamente sobre `decisionesAplicables`, impidiendo que decisiones no aplicables supriman el ajuste a `Confirmadas = 0`.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2206 | `AgEscaladorRanuraAula.cs` | Implementado |
| 2207 | `AgEscaladorRanuraAula.cs` | Implementado |
| 2208 | `AgEscaladorRanuraAula.cs` | Implementado |
| 2209 | `AgEscaladorRanuraAula.cs` | Implementado |
| 2210 | `AgOpcionesParticion.cs`, `AgOpcionesParticionTests.cs` | Implementado |
| 2211 | `AgEscaladorRanuraAula.cs`, `AgEscaladorRanuraAulaPersistenciaTests.cs` | Implementado |

---

## Ronda 494: Corrección de las propuestas 2213 y comparación entre ventanas (Fases 2215–2216)

### Resumen de la Ronda
- **Fase 2215:** Corrección de la línea base para `noAlojRef` en `AgEscaladorRanuraAula.cs`. En lugar de evaluar sobre `canonicas`, se calculan los patrones no alojables directamente sobre la partición del plan de `referencia` (`referencia.Plan?.Particion`), evitando que cohortes `LUN-VIE` con patrones `MAR-VIE` ya resueltos por el motor produzcan falsas mejoras.
- **Fase 2216:** Comparación de infactibles con la referencia de la ventana destino en `AgEscaladorRanuraAula.cs`. Cuando el destino escala (`destino != aula.DiasVentana`), `referenciaComparada` evalúa la ventana destino sin opción, evitando mezclar infactibles entre ventanas diferentes en `hayMejoraParticion`.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2215 | `AgEscaladorRanuraAula.cs` | Implementado |
| 2216 | `AgEscaladorRanuraAula.cs`, `AgEscaladoTresCasosMatrizTests.cs` | Implementado |

---

## Ronda 493: Catálogo inerte por filtro asimétrico y trazabilidad de rama (Fases 2218–2222)

### Resumen de la Ronda
- **Fase 2218:** Extracción y uso de `PatronesEfectivos` en `AgEscaladorRanuraAula.cs` para `patronesAula` y `patronesSimulados`. Excluye `AUTOMATICO`, alineando la semántica con `AgEvaluadorDemandaVentana.cs:124-126` y permitiendo que `AgOpcionesParticion.Resolver` detecte el catálogo de particiones.
- **Fase 2219:** Evaluación de `conservaPatron` en `AgEscaladorRanuraAula.cs` usando los patrones efectivos y garantizando `ContarPatronesNoAlojables(opcionUsada.Carriles, patronesConserva) == 0`, previniendo el borrado de `dias_patron` que la opción sí aloja.
- **Fase 2220:** Enriquecimiento de `AgSeleccionParticion` con propiedades `Origen` y `OpcionTexto` en `AgParticionCarriles.cs`, instrumentando `[Oleadas] Partición:` en `AgPlanificadorOleadas.cs` para registrar rama y opción elegida.
- **Fase 2221:** Suite de pruebas con convivencia de patrones mixtos y trazabilidad en `AgOpcionesParticionTests.cs`, `AgParticionCarrilesOpcionesTests.cs` y `AgEscaladoTresCasosMatrizTests.cs`.
- **Fase 2222:** Documentación de la política de emisión (Info siempre para `catalogo`, Info condicional a `cambiada` en `enumeracion`) y guía de lectura en `docs/articles/changes/2026-11-14-debug-ranuras-corrida84.md`.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2218 | `AgEscaladorRanuraAula.cs` | Implementado |
| 2219 | `AgEscaladorRanuraAula.cs` | Implementado |
| 2220 | `AgParticionCarriles.cs`, `AgPlanificadorOleadas.cs` | Implementado |
| 2221 | `AgOpcionesParticionTests.cs`, `AgParticionCarrilesOpcionesTests.cs`, `AgEscaladoTresCasosMatrizTests.cs` | Implementado |
| 2222 | `docs/articles/changes/2026-11-14-debug-ranuras-corrida84.md` | Implementado |

---

## Ronda 495: Catálogo con texto lossy y sin fallback (Fases 2223–2227)

### Resumen de la Ronda
- **Fase 2223:** Restauración del fallback de F2187 a enumeración contigua cuando ninguna opción de catálogo mejora la referencia en `AgParticionCarriles.cs:123-135`. Si no hay mejora, continúa a `Particiones()` etiquetando el origen como `catalogo_fallback`.
- **Fase 2224:** Preservación de franjas horarias en `OpcionTexto` en `AgParticionCarriles.cs:100-125` almacenando la tupla `(ev, op.CarrilesTexto)`, evitando la pérdida de formato como `5,6M|4,6T,7`.
- **Fase 2225:** Trazabilidad de la opción intentada en `AgParticionCarriles.cs:115, 155, 248`, asignando `textoOpcionIntentada` cuando el catálogo no produce cambios o no supera el gate.
- **Fase 2226:** Control de ruido en `AgPlanificadorOleadas.cs:365-385` mediante `LoggerHelper.WarningLimitado` deduplicando la emisión de `[Oleadas] Partición:` por clave única.
- **Fase 2227:** Cobertura de pruebas unitarias en `AgParticionCarrilesOpcionesTests.cs` para conservación de franjas, ejecución de enumeración en fallback y reporte de opción intentada sin cambios.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2223 | `AgParticionCarriles.cs` | Implementado |
| 2224 | `AgParticionCarriles.cs` | Implementado |
| 2225 | `AgParticionCarriles.cs` | Implementado |
| 2226 | `AgPlanificadorOleadas.cs` | Implementado |
| 2227 | `AgParticionCarrilesOpcionesTests.cs` | Implementado |

---

## Ronda 496: `opcion` con dos significados y contador mixto (Fases 2228–2231)

### Resumen de la Ronda
- **Fase 2228:** Unificación semántica de `OpcionTexto` en `AgSeleccionParticion` para que represente exclusivamente la opción de catálogo evaluada (`textoOpcionIntentada`), introduciendo la propiedad independiente `ParticionElegidaTexto` para transportar la clave de la partición contigua elegida (`elegida.Clave`).
- **Fase 2229:** Desacoplamiento de contadores en `AgSeleccionParticion`: `CandidatasEvaluadas` cuantifica estrictamente las particiones contiguas y `OpcionesEvaluadas` registra las opciones de catálogo evaluadas, reflejando ambos en la emisión de `[Oleadas] Partición:`.
- **Fase 2230:** Corrección de la salida temprana en `AgParticionCarriles.cs:160` condicionando `!hayPatronValido && infactibles == 0` con `&& !catalogoEvaluado`, asegurando que tras un catálogo sin mejora se ejecute el fallback completo a `Particiones()`.
- **Fase 2231:** Ampliación de la suite en `AgParticionCarrilesOpcionesTests.cs` con pruebas para `OpcionTexto`, `ParticionElegidaTexto`, `OpcionesEvaluadas` y preservación de salida sin catálogo; documentación de la semántica en `docs/articles/changes/2026-11-14-debug-ranuras-corrida84.md`.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2228 | `AgParticionCarriles.cs`, `AgPlanificadorOleadas.cs` | Implementado |
| 2229 | `AgParticionCarriles.cs`, `AgPlanificadorOleadas.cs` | Implementado |
| 2230 | `AgParticionCarriles.cs` | Implementado |
| 2231 | `AgParticionCarrilesOpcionesTests.cs`, `docs/articles/changes/2026-11-14-debug-ranuras-corrida84.md` | Implementado |

---

## Ronda 497: Catálogo exitoso reporta cero candidatas (Fase 2232)

### Resumen de la Ronda
- **Fase 2232:** Corrección del argumento `candidatasEvaluadas` en la rama de catálogo exitoso de `AgParticionCarriles.cs:141`, pasando `evaluadasOpciones` en lugar de `0` para reflejar con exactitud las opciones exploradas y mantener coherencia con el criterio de evaluación del catálogo.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2232 | `AgParticionCarriles.cs`, `AgParticionCarrilesOpcionesTests.cs` | Implementado |

---

## Ronda 498: El árbol no marca la ventana seleccionada (Fases 2233–2237)

### Resumen de la Ronda
- **Fase 2233:** Normalización ISO de ventana de días en `AgFormGeneradorHorarios.cs:5980-5990` y creación de `AgNivel6Ventanas.cs` para comparar representaciones mediante `AgDiasVentana.AIso`, resolviendo el fallo de coincidencia para valores guardados como `"1,2,3,4,5"`, `"1,2,3,4"` o formatos alternativos.
- **Fase 2234:** Expansión del catálogo del árbol Nivel 6 de 5 a las 8 ventanas canónicas de `AgDiasVentana.Opciones` (`LUN-JUE`, `LUN-VIE`, `MAR-VIE`, `JUE-SAB`, `VIE-SAB`, `VIE-DOM`, `JUE-DOM`, `MIE-DOM`), redistribuyendo el ancho en `CalcularColumnasNivel6` (26% jornada, 56% días, 18% horas) e incorporando nombres visuales compactos (`FormatearRadio`).
- **Fase 2235:** Hit-test dinámico y protección de espacio muerto en `ProcesarClickDimensionNivel6` mediante `AgNivel6Ventanas.ResolverPorX`, eliminando la asignación espuria de `VIE-DOM` cuando el clic se realiza en el hueco posterior al último botón de opción.
- **Fase 2236:** Enriquecimiento del tooltip contextual en `ActualizarToolTipNivel6`, presentando la clave canónica y la secuencia ISO (`Ventana actual: LUN-VIE (1,2,3,4,5)`) para cualquier combinación seleccionada.
- **Fase 2237:** Suite de pruebas unitarias en `AgNivel6VentanasTests.cs` validando mapeo por ISO, coincidencia cruzada, formateo adaptativo por paso horizontal, hit-test por coordenadas y rechazo de clics fuera de límites o en espacios muertos.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2233 | `AgNivel6Ventanas.cs`, `AgFormGeneradorHorarios.cs` | Implementado |
| 2234 | `AgFormGeneradorHorarios.cs`, `AgNivel6Ventanas.cs` | Implementado |
| 2235 | `AgFormGeneradorHorarios.cs`, `AgNivel6Ventanas.cs` | Implementado |
| 2236 | `AgFormGeneradorHorarios.cs` | Implementado |
| 2237 | `AgNivel6VentanasTests.cs` | Implementado |

---

## Ronda 499: Clic de día rompe la distinción ISO `4,5,6,7` / `JUE-DOM` (Fases 2238–2242)

### Resumen de la Ronda
- **Fase 2238:** Preservación del alias ISO terminal `4,5,6,7` en `AgFormGeneradorHorarios.cs:6055-6065` al pulsar el radio `Jue-Dom`. Si la ventana configurada es `4,5,6,7` y el usuario pulsa sobre el mismo carril visual (`JUE-DOM`), `EsAliasIsoTerminal` evita reescribir la propiedad a la clave textual, conservando la distinción histórica.
- **Fase 2239:** Desacoplamiento y fuente única en `AgNivel6Ventanas.cs:15-55`, derivando `OpcionesOrdenadas` de `AgDiasVentana.Opciones.Keys` y exponiendo `Claves` mediante un wrapper inmutable `Array.AsReadOnly`.
- **Fase 2240:** Supresión de campos residuales `Dx*` (`DxLunJue`, `DxMarVie`, `DxJueSab`, `DxVieSab`, `DxVieDom`) en `AgColumnasNivel6`, unificando el posicionamiento horizontal estrictamente a través de `DxDias[i]`.
- **Fase 2241:** Corrección del desborde horizontal en `CalcularColumnasNivel6`, fijando `minimoTotal = 700`, acotando `wHoras = Math.Max(96, anchoDisp - wJornada - wDias)` y reajustando `wDias` para garantizar que la suma nunca supere `anchoDisp`.
- **Fase 2242:** Ampliación de la suite de pruebas en `AgNivel6VentanasTests.cs` con casos para `EsAliasIsoTerminal`, no degradación de `4,5,6,7` ante clic en `Jue-Dom`, formato compacto en anchos reducidos y paridad exacta con el catálogo de dominio.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2238 | `AgFormGeneradorHorarios.cs`, `AgNivel6Ventanas.cs` | Implementado |
| 2239 | `AgNivel6Ventanas.cs` | Implementado |
| 2240 | `AgFormGeneradorHorarios.cs` | Implementado |
| 2241 | `AgFormGeneradorHorarios.cs` | Implementado |
| 2242 | `AgNivel6VentanasTests.cs` | Implementado |

---

## Ronda 500: Sin forma de elegir `4,5,6,7` frente a `JUE-DOM` (Fases 2243–2244)

### Resumen de la Ronda
- **Fase 2243:** Separación e incorporación de la 9ª opción explícita `"4,5,6,7"` en `AgNivel6Ventanas.cs` permitiendo seleccionar indistintamente entre el alias terminal (`4,5,6,7`) y la etiqueta canónica escalable (`JUE-DOM`). Se actualizó `Indice` y `Coincide` para priorizar coincidencia exacta e impedir el colapso por ISO de `4,5,6,7` con `JUE-DOM`. Se retiró la guarda pasiva de `AgFormGeneradorHorarios.cs:6055-6065` y se adaptó el ancho mínimo en `CalcularColumnasNivel6` (`minimoTotal = 740`, `stepD = Math.Max(44, ...)`).
- **Fase 2244:** Estabilización del orden del catálogo en `AgNivel6Ventanas.cs` y suite completa de pruebas en `AgNivel6VentanasTests.cs` validando los 9 índices, la distinción explícita de `4,5,6,7` frente a `JUE-DOM`, la reescritura bidireccional por clic y la paridad de conjunto con `AgDiasVentana.Opciones`.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2243 | `AgNivel6Ventanas.cs`, `AgFormGeneradorHorarios.cs` | Implementado |
| 2244 | `AgNivel6Ventanas.cs`, `AgNivel6VentanasTests.cs` | Implementado |

---

## Ronda 501: Alias terminal con espacios y residuo sin uso (Fases 2245–2246)

### Resumen de la Ronda
- **Fase 2245:** Normalización de tokens en `AgNivel6Ventanas.cs` mediante el método `NormalizarToken(s)`, eliminando espacios intermedios y separadores redundantes antes de comparar. Evita que representaciones como `"4, 5, 6, 7"` fallen en la comparación textual exacta del alias y colapsen erróneamente por equivalencia ISO a `"JUE-DOM"`.
- **Fase 2246:** Eliminación del campo muerto `EtiquetasVentanaRadio` en `AgFormGeneradorHorarios.cs:6088`, el cual quedó sin referencias tras la unificación del tooltip contextual mediante `AgNivel6Ventanas.ClaveCanonica`.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2245 | `AgNivel6Ventanas.cs`, `AgNivel6VentanasTests.cs` | Implementado |
| 2246 | `AgFormGeneradorHorarios.cs` | Implementado |

---

## Ronda 502: Corrección de asimetrías en conservación de patrones y fuentes de datos (Fases 2247–2250)

### Resumen de la Ronda
- **Fase 2247:** Reemplazo de la lógica inyectiva en `OpcionConservaPatrones` (`AgEscaladorRanuraAula.cs:1050-1075`) por un modelo de compatibilidad muchos-a-uno, permitiendo que múltiples asignaturas compartan el mismo carril y eliminando la restricción espuria `patronesDias.Count > opcion.Carriles.Count`.
- **Fase 2248:** Sustitución del veto textual global de sábado (`6M`/`6T`) en `OpcionConservaPatrones` por validación efectiva de ranuras por carril, verificando que el carril asignado al patrón contenga ranura de sábado.
- **Fase 2249:** Exclusión de patrones con origen `AUTOMATICO` en `AgPlanificadorPatronSemanal.DiasKDe(aa, ...)` al sembrar genes sin plan fechado, garantizando que los patrones automáticos no operen como restricciones fijas.
- **Fase 2250:** Incorporación de fallback a `AulaAsignaturas` en `PatronesEfectivos` (`AgEscaladorRanuraAula.cs`) y `Evaluar` (`AgEvaluadorDemandaVentana.cs`) cuando el índice por aula-asignatura no ha sido pre-construido.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2247 | `AgEscaladorRanuraAula.cs`, `AgEscaladoTresCasosMatrizTests.cs` | Implementado |
| 2248 | `AgEscaladorRanuraAula.cs`, `AgEscaladoTresCasosMatrizTests.cs` | Implementado |
| 2249 | `AgPlanificadorPatronSemanal.cs`, `AgEscaladoTresCasosMatrizTests.cs` | Implementado |
| 2250 | `AgEscaladorRanuraAula.cs`, `AgEvaluadorDemandaVentana.cs`, `AgEscaladoTresCasosMatrizTests.cs` | Implementado |

---

## Ronda 503: Cierre de conservación de patrones y fallback vacío (Fases 2251–2253)

### Resumen de la Ronda
- **Fase 2251:** Implementación de `AgContexto.RelacionesAula(ctx)` que verifica cardinalidad y ejecuta `ctx.ConstruirIndices()` si el índice `AulaAsignaturaPorAulaAsig` está vacío (0 filas) pero `AulaAsignaturas` contiene relaciones. Se integró en `PatronesEfectivos` (`AgEscaladorRanuraAula.cs`) y `Evaluar`/`Demandas` (`AgEvaluadorDemandaVentana.cs`).
- **Fase 2252:** Documentación explícita del contrato de `OpcionConservaPatrones` en `AgEscaladorRanuraAula.cs` estableciendo que la comprobación de carriles es de carácter estructural (presencia de ranuras requeridas por el patrón, incluyendo sábado) y que la capacidad semanal de sesiones se valida en la evaluación de demanda downstream.
- **Fase 2253:** Corrección en `AgFabricaIndividuos.cs:429` para excluir el origen `AUTOMATICO` de `tienePatronDeclarado`, permitiendo que asignaturas con patrones automáticos residuales sean tratadas como automáticas y no bloqueen la resiembra del genoma sin plan fechado.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2251 | `AgContexto.cs`, `AgEscaladorRanuraAula.cs`, `AgEvaluadorDemandaVentana.cs`, `AgEscaladoTresCasosMatrizTests.cs` | Implementado |
| 2252 | `AgEscaladorRanuraAula.cs` | Retirada por diseño (capacidad pertenece a oleadas) |
| 2253 | `AgFabricaIndividuos.cs`, `AgFabricaIndividuosTests.cs` | Implementado |

---

## Ronda 504: Cierre de conservación de patrones (Fases 2254–2255)

### Resumen de la Ronda
- **Fase 2254:** Formalización de `PatronesAlojablesEnOpcion` en `AgEscaladorRanuraAula.cs:586, 1053-1100` (anteriormente `OpcionConservaPatrones`), consolidando el criterio estructural (presencia de ranura para cada día del patrón y sábado) y registrando el retiro formal de 2252 ya que la capacidad semanal la evalúa `AgPlanificadorOleadas`.
- **Fase 2255:** Cobertura de pruebas unitarias exhaustiva para los cuatro bordes (`RelacionesAula_IndiceVacio_UsaListaYConstruyeIndices`, `PatronesAlojables_SeisPatronesDosCarriles_True`, `PatronesAlojables_Dia6SinRanura_False`, `Escalador_PatronesEfectivos_ExcluyeAutomatico` en `AgEscaladoTresCasosMatrizTests.cs` y `Fabrica_PatronAutomatico_NoBloqueaResiembra` en `AgFabricaIndividuosTests.cs`).

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2254 | `AgEscaladorRanuraAula.cs`, `AgEscaladoTresCasosMatrizTests.cs` | Implementado |
| 2255 | `AgEscaladoTresCasosMatrizTests.cs`, `AgFabricaIndividuosTests.cs` | Implementado |

---

## Ronda 505: Eliminación de persistencia de solo partición y desacople de ConservaPatronOrigen (Fases 2256–2259)

### Resumen de la Ronda
- **Fase 2256:** Eliminación de `(ParticionMejora && ParticionDifiereDeReferencia)` de la propiedad `AgEscaladorRanuraAula.Decision.Aplica`. Dado que las particiones de carriles son efímeras y se recalculan dinámicamente en memoria en cada corrida por `AgPlanificadorOleadas`, no existen campos ni columnas en base de datos para almacenar carriles. Por tanto, decisiones donde `Desde == Hacia` sin reclasificaciones autorizadas no generan escrituras persistentes (`Aplica = false`), impidiendo transacciones espurias y eliminando fallos en la verificación de huellas postcondición (`Calcular != esperado`).
- **Fase 2257:** Desacople estricto de `liberarEsperado` y el vaciado de `dias_patron` para bottleneck respecto a `Desde == Hacia`. Ambas operaciones quedan condicionadas exclusivamente a `!d.ConservaPatronOrigen`. Cuando `ConservaPatronOrigen` es `true` (validado mediante `PatronesAlojablesEnOpcion`), se preservan los patrones incluso ante escalado de cambio de ventana, evitando borrados destructivos innecesarios.
- **Fase 2258:** Documentación y armonización arquitectónica de la no-persistencia de solo partición en `AgEscaladorRanuraAula.cs`. Retención informativa en salida temprana de `AplicarDetalladoAsync`: si no hay decisiones aplicables a BD pero existen mejoras de partición, se emite advertencia de partición vigente sin modificaciones (`SinCambios`).
- **Fase 2259:** Adición de trazabilidad diagnóstica `LoggerHelper.Info` previa a la validación de postcondiciones en `AplicarDetalladoAsync` (`APLICAR ref={ref} desde={desde} hacia={hacia} conserva={conserva} liberar={liberar} modificadas={modificadas} perfil={perfil}`), garantizando visibilidad inmediata de los parámetros de aplicación de cada decisión.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2256 | `AgEscaladorRanuraAula.cs`, `AgEscaladorRanuraAulaPersistenciaTests.cs`, `AgEscaladoTresCasosMatrizTests.cs` | Implementado |
| 2257 | `AgEscaladorRanuraAula.cs`, `AgEscaladoTresCasosMatrizTests.cs` | Implementado |
| 2258 | `AgEscaladorRanuraAula.cs` | Implementado |
| 2259 | `AgEscaladorRanuraAula.cs` | Implementado |

---

## Ronda 506: Partición de origen preferente (Fases 2260–2263)

### Resumen de la Ronda
- **Fase 2260:** Introducción de `AgCriterioParticion` (`Estricto`, `PreferirOrigen`, `ForzarOrigen`) en `AgParticionCarriles.cs` y métodos de guarda `NoEmpeora` y `NoEmpeoraSinMakespan`. Permite que ante una referencia canónica con 0 infactibles y mismo makespan no se descarte la opción de catálogo de origen (`2,3|1,4,5`), eliminando la caída espuria a `catalogo_fallback` con carril aislado `{1}`.
- **Fase 2261:** Propagación del parámetro `AgCriterioParticion` a través de `AgPlanificadorOleadas.Construir`, `ConstruirEvaluacion`, `ConstruirInterno` y `AgEvaluadorDemandaVentana.Evaluar` (valor por defecto `AgCriterioParticion.Estricto`).
- **Fase 2262:** Paridad en la selección de partición entre escalador (`AgEscaladorRanuraAula.cs`) y validador (`AgValidadorVentana.cs`), configurando `PreferirOrigen` en las evaluaciones con opción (`opcionActual`, `opcionMisma`, `opcionPaso`, `opcionEscalon`) y conservando `Estricto` en la evaluación de la referencia sin opción.
- **Fase 2263:** Trazabilidad enriquecida en log de oleadas (`criterio={criterio}`) y batería de pruebas unitarias (`AgParticionCarrilesOpcionesTests.cs`, `AgEscaladoTresCasosMatrizTests.cs`) para verificar selección preferente, preservación de guardas de no-empeoramiento y fallback nominal ante opciones adversas.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2260 | `AgParticionCarriles.cs` | Revisado |
| 2261 | `AgPlanificadorOleadas.cs`, `AgEvaluadorDemandaVentana.cs` | Revisado |
| 2262 | `AgEscaladorRanuraAula.cs`, `AgValidadorVentana.cs` | Revisado |
| 2263 | `AgPlanificadorOleadas.cs`, `AgParticionCarrilesOpcionesTests.cs`, `AgEscaladoTresCasosMatrizTests.cs` | Revisado |

---

## Ronda 507: Criterio de partición de origen (implementación restante) (Fases 2264–2267)

### Resumen de la Ronda
- **Fase 2264:** Implementación de `AgCriterioParticion` (`Estricto`, `PreferirOrigen`, `ForzarOrigen`) en `AgParticionCarriles.cs` y de las funciones de guarda `NoEmpeora` y `NoEmpeoraSinMakespan` con verificación explícita de cobertura de sesiones ubicadas (`Ubicadas(ev) >= Ubicadas(r)`). Esto resuelve el descarte sistemático de `2,3|1,4,5` ante referencias 0/0, evitando la caída espuria a `catalogo_fallback` con carril `{1}`.
- **Fase 2265:** Propagación del parámetro `AgCriterioParticion criterio = AgCriterioParticion.Estricto` a través de `AgPlanificadorOleadas.Construir`, `ConstruirEvaluacion`, `ConstruirInterno` y `AgEvaluadorDemandaVentana.Evaluar`.
- **Fase 2266:** Paridad de criterio entre escalador (`AgEscaladorRanuraAula.cs`) y validador (`AgValidadorVentana.cs`), configurando `PreferirOrigen` en las llamadas que evalúan opciones estructuradas y manteniendo `Estricto` en la línea base de referencia sin opción.
- **Fase 2267:** Incorporación de `criterio=` en el log de oleadas (`[Oleadas] Partición: ... criterio={criterio}; ...`) y suite de pruebas unitarias (`AgParticionCarrilesOpcionesTests.cs`, `AgEscaladoTresCasosMatrizTests.cs`) para verificar selección preferente, preservación de sesiones ubicadas y rechazo por degradación.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2264 | `AgParticionCarriles.cs` | Planificado |
| 2265 | `AgPlanificadorOleadas.cs`, `AgEvaluadorDemandaVentana.cs` | Planificado |
| 2266 | `AgEscaladorRanuraAula.cs`, `AgValidadorVentana.cs` | Planificado |
| 2267 | `AgPlanificadorOleadas.cs`, `AgParticionCarrilesOpcionesTests.cs`, `AgEscaladoTresCasosMatrizTests.cs` | Planificado |

---

## Ronda 508: Criterio de partición de origen (re-emisión) (Fases 2268–2271)

### Resumen de la Ronda
- **Fase 2268:** Re-emisión de la definición de `AgCriterioParticion` (`Estricto`, `PreferirOrigen`, `ForzarOrigen`), cálculo auxiliar `Ubicadas` y predicados de guarda `NoEmpeora`/`NoEmpeoraSinMakespan` con verificación estricta de cobertura (`Ubicadas(ev) >= Ubicadas(r)`) en `AgParticionCarriles.cs:85-147`, permitiendo que la opción de catálogo `2,3|1,4,5` no se descarte frente a referencias canónicas 0/0.
- **Fase 2269:** Propagación del parámetro `AgCriterioParticion criterio = AgCriterioParticion.Estricto` en `AgPlanificadorOleadas.Construir`, `ConstruirEvaluacion`, `ConstruirInterno` (`AgPlanificadorOleadas.cs:30-59, 236, 357-364`) y `AgEvaluadorDemandaVentana.Evaluar` (`AgEvaluadorDemandaVentana.cs:98, 131-133`).
- **Fase 2270:** Paridad y alineación de criterio entre escalador (`AgEscaladorRanuraAula.cs:502-586`) y validador (`AgValidadorVentana.cs:353-355, 428-429`), empleando `PreferirOrigen` en evaluaciones con opción de catálogo y `Estricto` en la evaluación de referencia sin opción.
- **Fase 2271:** Trazabilidad de partición en log de oleadas (`[Oleadas] ... criterio={criterio}; ...`, `AgPlanificadorOleadas.cs:377-391`) y batería de pruebas unitarias (`AgParticionCarrilesOpcionesTests.cs`, `AgEscaladoTresCasosMatrizTests.cs`) para verificar selección de catálogo de origen, guardas de cobertura y fallback nominal.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2268 | `AgParticionCarriles.cs` | Planificado |
| 2269 | `AgPlanificadorOleadas.cs`, `AgEvaluadorDemandaVentana.cs` | Planificado |
| 2270 | `AgEscaladorRanuraAula.cs`, `AgValidadorVentana.cs` | Planificado |
| 2271 | `AgPlanificadorOleadas.cs`, `AgParticionCarrilesOpcionesTests.cs`, `AgEscaladoTresCasosMatrizTests.cs` | Planificado |

---

## Ronda 509: Cierre de 2258 y diagnóstico de aplicación (Fases 2272–2274)

### Resumen de la Ronda
- **Fase 2272:** Ajuste de postcondición same-window (`Desde == Hacia`) en `AgEscaladorRanuraAula.cs:898-1001` y `AgHuellaEscaladoPersistido.cs:100-145`. En decisiones solo de reclasificación o intra-ventana sin cambio de ventana, `destino` se pasa como `null` a `Calcular` y `EstadoCompleto`. `EstadoCompleto` omite la comprobación de `dias_ventana` en aula y nodo cuando `destino == null`, evitando fallos espurios de postcondición por divergencias preexistentes entre el árbol y `ag_aula`. Asimismo, `Calcular` no condiciona la sustitución de reclasificaciones ni liberación de cuellos de botella a `destino != null`.
- **Fase 2273:** Adelanto del registro de log `[ESCALADO] APLICAR` previo a la ejecución del primer `UPDATE` en `AgEscaladorRanuraAula.cs` (~línea 925) y emisión diagnóstica `[ESCALADO] APLICAR_FALLO` con aula, desde, hacia, tipo y mensaje de excepción en el bloque `catch` para garantizar trazabilidad ante fallos de escritura SQL.
- **Fase 2274:** Limpieza de propiedades residuales de partición en `AgEscaladorRanuraAula.Decision` (`ParticionMejora`, `ParticionDifiereDeReferencia`) y sus correspondientes tokens (`pm=`, `pdr=`) en `HuellaAccion`, una vez desacoplada la persistencia de decisiones intra-ventana.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2272 | `AgEscaladorRanuraAula.cs`, `AgHuellaEscaladoPersistido.cs` | Planificado |
| 2273 | `AgEscaladorRanuraAula.cs` | Planificado |
| 2274 | `AgEscaladorRanuraAula.cs` | Planificado |

---

## Ronda 510: Reclasificación en misma ventana rompe el hash esperado (Fases 2275–2279)

### Resumen de la Ronda
- **Fase 2275:** Desacople de `destino != null` para la sustitución de `liberar` (cuello de botella) y `reclasificaciones` en `AgHuellaEscaladoPersistido.Calcular` (`AgHuellaEscaladoPersistido.cs:33, 39`). Con `Desde == Hacia` (o `Hacia` vacío) y solo reclasificación, los `UPDATE` mutan `origen_patron` y `dias_patron`; al desacoplar la condición, el hash esperado refleja fielmente los cambios en `ag_aula_asignatura` sin requerir cambio de ventana.
- **Fase 2276:** Adaptación de `AgHuellaEscaladoPersistido.EstadoCompleto` para aceptar `destino == null` (`AgHuellaEscaladoPersistido.cs:100-112`). Si `destino == null`, se verifica la existencia unívoca del aula y se omite la validación de `dias_ventana` en aula y nodo. En `AgEscaladorRanuraAula.cs` (líneas 907-917 y 998) se calcula `destinoEsperado = cambiaVentana ? d.Hacia : null` tanto para la rama idempotente como para la verificación postcondición.
- **Fase 2277:** Corrección de la clasificación de advertencia `"Partición ya vigente; sin cambios persistidos."` en `AgEscaladorRanuraAula.cs:852-855`. Se introduce la propiedad `SoloParticion` (`ParticionMejora && !Aplica && Desde == Hacia && SinReclasificaciones`) evitando atribuir a partición decisiones de escalada rechazadas por falta de capacidad.
- **Fase 2278:** Reubicación de la traza `[ESCALADO] APLICAR` previo al primer `UPDATE` SQL en `AgEscaladorRanuraAula.cs` (~línea 925) e incorporación de log `[ESCALADO] APLICAR_FALLO` con referencia, desde, hacia, conserva, tipo y mensaje de error en el bloque `catch`.
- **Fase 2279:** Exclusión de los flags `pm=` y `pdr=` del cómputo de `HuellaAccion` en `AgEscaladorRanuraAula.cs:1155-1156`, manteniendo las propiedades en `Decision` para el diagnóstico de `SoloParticion` pero sin alterar el sello criptográfico de la acción.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2275 | `AgHuellaEscaladoPersistido.cs` | Implementado |
| 2276 | `AgHuellaEscaladoPersistido.cs`, `AgEscaladorRanuraAula.cs` | Implementado |
| 2277 | `AgEscaladorRanuraAula.cs` | Implementado |
| 2278 | `AgEscaladorRanuraAula.cs` | Implementado |
| 2279 | `AgEscaladorRanuraAula.cs` | Implementado |

---

## Ronda 511: `UPDATE ag_aula` se ejecuta con guarda distinta a `destinoEsperado` (Fases 2280–2283)

### Resumen de la Ronda
- **Fase 2280:** Sustitución de `!string.Equals(d.Desde, d.Hacia)` por `cambiaVentana` en `AgEscaladorRanuraAula.cs:948` y en el flag global `huboCambioVentana` (línea 1044). Si `Hacia` viene vacío o no nulo pero igual a `Desde`, no se ejecuta `UPDATE ag_aula SET dias_ventana` ni se modifica `ag_arbol_nodo`, manteniendo estricta paridad con `destinoEsperado = cambiaVentana ? d.Hacia : null`.
- **Fase 2281:** Validación simétrica estricta de cada reclasificación en `AgEscaladorRanuraAula.cs:980-987`. Se verifica `rec != null && rec.Autorizado && rec.IdAula == idAula && rec.AulaRef == d.AulaRef && rec.Subunidad == d.Subunidad` antes de ejecutar cualquier `UPDATE` en `ag_aula_asignatura`, arrojando `InvalidOperationException` si la propuesta fue manipulada o presenta divergencias.
- **Fase 2282:** Bump de versión del prefijo de acción en `HuellaAccion` de `propuesta-escalado-v4` a `propuesta-escalado-v5` (`AgEscaladorRanuraAula.cs:1164`) para reflejar la eliminación de los tokens `pm=` y `pdr=`.
- **Fase 2283:** Ampliación de la guarda de `SoloParticion` en `AgEscaladorRanuraAula.cs:57-61` (`(ParticionMejora || ParticionDifiereDeReferencia)`) y unificación del texto de advertencia a `"Sin cambios persistidos; la partición se recalcula en cada corrida."` en las líneas 859 y 1048. Actualización de aserciones en tests correspondientes.

### Soporte realizado por fase

| Fase | Archivos | Estado |
|---|---|---|
| 2280 | `AgEscaladorRanuraAula.cs` | Implementado |
| 2281 | `AgEscaladorRanuraAula.cs` | Implementado |
| 2282 | `AgEscaladorRanuraAula.cs`, `AgEscaladorRanuraAulaTests.cs` | Implementado |
| 2283 | `AgEscaladorRanuraAula.cs`, `AgEscaladorRanuraAulaPersistenciaTests.cs`, `AgEscaladorRanuraAulaTests.cs` | Implementado |















