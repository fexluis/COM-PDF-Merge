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































