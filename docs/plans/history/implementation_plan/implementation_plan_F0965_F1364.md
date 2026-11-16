<!-- historico congelado - generado por .agents/tools/split-history.ps1 - commit d0ee6fe - no editar -->
# implementation_plan - Fases 965-1364 (Rondas 232)

> Indice: [../../../implementation_plan.md](../../../implementation_plan.md) - Escision: 2026-11-16

---
### 157.137 Ronda 181 — F964 Aplicada con Error de Compilación + 4 Defectos en Código Nuevo (Fases 965 a 969)

#### 1. Fase 965 — `horasAcumuladas` eliminada por F964 → CS0103 ×4
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgExpansorSesiones.cs`
- **Diagnóstico:** El reemplazo del bloque de inicialización en F964 eliminó la declaración `double horasAcumuladas = 0.0;`, produciendo 4 errores CS0103 en líneas de acumulación y control.
- **Solución Propuesta:** Restaurar `double horasAcumuladas = 0.0;` tras la inicialización de `sesiones`.

#### 2. Fase 966 — `GetUltimaCorridaAsync` declarada sin implementar → CS0535
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Domain/Interfaces/IAgScheduleRepository.cs`, `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgScheduleRepository.cs`
- **Diagnóstico:** El contrato `IAgScheduleRepository` expone `GetUltimaCorridaAsync(int, int, string, CancellationToken)`, pero la implementación en Dapper no existía en `AgScheduleRepository.cs`.
- **Solución Propuesta:** Implementar consulta Dapper `SELECT ... FROM ag_corrida WHERE idperiodo = @idPeriodo AND idprograma = @idPrograma AND (@subunidad IS NULL OR subunidad = @subunidad) ORDER BY id DESC LIMIT 1` con soporte de `vacantes` y cancelación.

#### 3. Fase 968 — Coherencia cuantiza con `4.0` fijo → posgrado 8 h viola F964 y aborta el individuo
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgCoherenciaGrupo.cs`
- **Diagnóstico:** `AgCoherenciaGrupo` cuantizaba sesiones y proyectaba horas usando la constante `4.0` en la rama principal. En posgrados de 8h (`HorasDia=8`), generaba asignaciones como 20h (`20 / 8 = 2.5`), provocando rechazo fail-fast en F964 y aborto total de la expansión.
- **Solución Propuesta:** Promover función local `ObtenerUnidad(AgGen g)` al ámbito del grupo utilizando `AgUnidadesSemanales.HorasSesionAula` y reemplazar todos los factores `4.0` por `unidad` dinámica.

#### 4. Fase 967 — `AgBloqueSolape` y `R14_BloqueSolapadoPorIntervalo` sin emisor → regla inerte
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorReglas.cs`
- **Diagnóstico:** El código enum 33 (`R14_BloqueSolapadoPorIntervalo`) y el evaluador de intervalos `AgBloqueSolape.SeSolapan` carecían de llamada en `AgValidadorReglas`, permitiendo que bloques solapados pasaran la compuerta.
- **Solución Propuesta:** Conectar evaluación en `AgValidadorReglas.Evaluar` sobre genes con bloque asignado, validando solape temporal por intervalo cuando compartan ámbito o recurso (aula, docente, cohorte/grupo) e incrementando `hardCount`.

#### 5. Fase 969 — Un gen inválido cancela la materialización completa del individuo
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgExpansorSesiones.cs`
- **Diagnóstico:** `ExpandirIndividuo` retornaba `Result.Failure` al primer gen no cuantizado, descartando todas las demás sesiones válidas sin reportar el detalle de los genes fallidos.
- **Solución Propuesta:** Aislar la falla por gen acumulando errores, registrando advertencia diagnóstica y permitiendo el retorno de las sesiones válidas generadas, fallando únicamente si el individuo queda sin ninguna sesión expandible.

---

### 157.138 Ronda 182 — Discriminación de Cohorte en R14, Optimización Temporal y Divisibilidad de Carga (Fases 970 a 972)

#### 1. Fase 970 — F967 compara cohortes entre programas/sedes distintas → R14 hard masivo
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorReglas.cs`
- **Diagnóstico:** La condición `g1.Grupo == g2.Grupo && g1.Semestre == g2.Semestre` en `mismoRecurso` no evaluaba el programa ni la sede. Programas o sedes diferentes que comparten nomenclatura de semestres y grupos (ej. Semestre "1", Grupo "1") provocaban que dos genes sin conflicto de recursos se considerasen la misma cohorte, emitiendo violaciones duras `R14_BloqueSolapadoPorIntervalo` y tornando casi cualquier individuo en `Infactible`.
- **Solución Propuesta:** Condicionar la equivalencia de cohorte a `g1.IdPrograma == g2.IdPrograma && g1.IdSede == g2.IdSede && !string.IsNullOrWhiteSpace(g1.Grupo) && !string.IsNullOrWhiteSpace(g1.Semestre) && g1.Grupo == g2.Grupo && g1.Semestre == g2.Semestre`.

#### 2. Fase 971 — F967 evalúa O(n²) sobre todos los genes → bloqueo en poblaciones reales
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorReglas.cs`
- **Diagnóstico:** La evaluación de solapes recorría todos los pares de genes con bloque mediante doble bucle $O(n^2)$ sin particionamiento previo. Con tamaños de genoma típicos (2.000 genes), esto requería ~2M de iteraciones por individuo en cada generación, congelando la ejecución.
- **Solución Propuesta:** Particionar previamente los genes por día (`conBloque.GroupBy(x => x.Bloque.DiaSemanaNumero)`), ejecutando la comparación de pares únicamente dentro de la lista de cada día y aplicando corte temprano del bucle interno (`cortado = true; break;`).

#### 3. Fase 972 — F968 usa unidad de 8 h sin validar múltiplo → I1 queda irresoluble
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgCoherenciaGrupo.cs`
- **Diagnóstico:** En grupos de posgrado con sesiones de 8h, si la carga total requerida (`esperado`) no es múltiplo exacto de 8h (ej. 60h), redondear sobre unidad 8.0h produce desfases permanentes $\Sigma \neq I1$ (64h) que el rebalanceo de residuos no puede reconciliar.
- **Solución Propuesta:** Comprobar si `esperado` es divisible exactamente por `unidad` (`Math.Abs(esperado / unidad - Math.Round(esperado / unidad)) > 0.001`); de no serlo, degradar a `unidad = 4.0` (regla D4′).

---

### 157.139 Ronda 183 — Reconciliación de Bloques en Fábrica, Remanente de Coherencia y Limpieza R14 (Fases 973 a 975)

#### 1. Fase 973 — Fábrica asigna bloque de 8 h sin validar que la carga sea múltiplo → F964 rechaza y F969 omite el gen
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`
- **Diagnóstico:** La fábrica asignaba franjas horarias con `horasEfectivas = HorasSesionAula(aula, nivelAula)` (8h en posgrado) sin considerar si la carga de la asignatura era múltiplo de 8h. En asignaturas de 60h, Coherencia reducía la unidad a 4h pero el gen conservaba el bloque de 8h, causando que `AgExpansorSesiones` rechazara la cuantización en F964 y F969 omitiera el gen.
- **Solución Propuesta:** Evaluar `cargaUnidad / horasEfectivas` (usando `totalHoras` en día único y `totalHoras / diasK.Count` en multi-día) y degradar `horasEfectivas = 4` cuando `horasEfectivas > 4` y la carga no sea divisible entre 8h.

#### 2. Fase 974 — F972 valida `esperado` pero no `horasRestantes` → residuo de fijados rompe I1
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgCoherenciaGrupo.cs`
- **Diagnóstico:** Si `esperado` es múltiplo de 8 (ej. 64h) pero los genes fijados suman un valor que deja un remanente no divisible entre 8 (ej. 20h fijadas $\rightarrow$ 44h restantes), la unidad permanecía en 8.0h provocando que la suma proyectada discrepara de `esperado` sin posibilidad de ajuste residual.
- **Solución Propuesta:** Degradar `unidad = 4.0` si `esperado` o `horasRestantes` no son múltiplos exactos de `unidad`.

#### 3. Fase 975 — `cortado` + `continue` es código muerto en F971
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorReglas.cs`
- **Diagnóstico:** La bandera `cortado` y la condición `if (cortado) continue;` al final del bucle `i` eran superfluas, ya que `break;` cortaba el bucle interior `j` e `i` continuaba naturalmente al siguiente elemento.
- **Solución Propuesta:** Remover la variable `cortado` y la cláusula `continue`, conservando el `break;` directo dentro del ciclo de pares.

---

### 157.140 Ronda 184 — Sincronización de Horas Semanales y Cuantización de Fallback en Fábrica (Fases 976 y 977)

#### 1. Fase 976 — `HorasSemanales` fijo en 4.0 contradice bloque de 8 h → carga por gen subestimada
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`
- **Diagnóstico:** Asignar bloques de 8h en programas de posgrado mientras se retenía `HorasSemanales = 4.0` en `AgFabricaIndividuos` subestimaba a la mitad el conteo de horas calculadas por sesión (`sesionesEsteDia * HorasSemanales`), trasladando un desbalance masivo al ajuste posterior de residuo I1 y generando horas no múltiplos de 8h que F964 rechazaba.
- **Solución Propuesta:** Derivar dinámicamente `gen.HorasSemanales = bloque.DuracionHoras > 0 ? bloque.DuracionHoras : 4.0` tras asignar el bloque tanto en la rama de día único como en la de multi-día, eliminando la constante rígida en la instanciación.

#### 2. Fase 977 — Fallback `totalHoras / diasK.Count` sin cuantizar → F964 rechaza el gen
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`
- **Diagnóstico:** El cálculo de fallback en multi-día dividía linealmente `totalHoras / diasK.Count` sin redondear a múltiplos de la unidad de bloque, generando duraciones no cuantizadas (ej. 10h con bloque de 4h) que eran descartadas por la compuerta F964 y excluidas por F969.
- **Solución Propuesta:** Cuantizar el fallback a la unidad de sesión del gen mediante `Math.Max(unidadDia, Math.Round(cargaDia / unidadDia) * unidadDia)` con `unidadDia = gen.HorasSemanales > 0 ? gen.HorasSemanales : 4.0`.

---

### 157.141 Ronda 185 — Sincronización de Unidad de Sesión en Validador de Ventana y Contención D14 (Fases 978 y 979)

#### 1. Fase 978 — `sesionesRequeridas = ceil(hAsig / 4.0)` bloquea posgrado 8 h
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorVentana.cs`
- **Diagnóstico:** En asignaturas de posgrado con sesiones de 8h (`HorasDia = 8`), calcular las sesiones requeridas como `ceil(hAsig / 4.0)` duplicaba artificialmente la demanda de sesiones (ej. 16 sesiones requeridas para 64h en vez de 8). En ventanas semestrales típicas de 12 semanas (capacidad real = 12 ocurrencias de días K), se disparaba un bloqueo falso de capacidad insuficiente ($16 > 12$).
- **Solución Propuesta:** Calcular `unidadAsig` mediante `AgUnidadesSemanales.HorasSesionAula(aulaAsig, AgNivelAcademicoCatalogo.Desde(progAsig?.Modalidad))` y obtener `sesionesRequeridas = (int)Math.Ceiling(hAsig / unidadAsig)`.

#### 2. Fase 979 — Validación D14 usa `auCfg.HorasDia` crudo → exige bloque de 8 h inexistente
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorVentana.cs`
- **Diagnóstico:** La validación D14 pasaba `auCfg.HorasDia` crudo a `Validar`. En aulas de pregrado con `horas_dia = 8`, esto provocaba la búsqueda de bloques horarios de 8h (inexistentes en catálogo para pregrado según la regla D4′ de bloques de 4h), resultando en un error bloqueante en pre-vuelo.
- **Solución Propuesta:** Evaluar `horasEfectivasCfg = (int)AgUnidadesSemanales.HorasSesionAula(auCfg, AgNivelAcademicoCatalogo.Desde(progCfg?.Modalidad))` y pasarlo a `Validar` en lugar del valor crudo de la entidad de aula.

---

### 157.142 Ronda 186 — Resolución de Unidad desde Aula Ganadora y Conteo de Capacidad por Sesiones (Fases 980 y 981)

#### 1. Fase 980 — Unidad 4 h cuando `aa == null` → F978 inoperante en primera corrida
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorVentana.cs`
- **Diagnóstico:** `unidadAsig` se calculaba de forma aislada a partir del mapeo previo `aa` antes de evaluar `configs`. En primeras corridas o ante asignaturas sin aula preasignada (`aa == null`), caía en fallback de 4.0h y volvía a sobredimensionar a 16 sesiones requeridas las asignaturas de 64h en posgrado, reactivando bloqueos falsos.
- **Solución Propuesta:** Extraer función `UnidadDe(au)` y diferir la resolución de `unidadAsig` y `sesionesRequeridas` al aula ganadora de `configs` (`mejorCfg`), asegurando simetría total entre la unidad de demanda y la oferta de capacidad.

#### 2. Fase 981 — Capacidad cuenta 1 sesión por día ignorando AM/PM
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorVentana.cs`
- **Diagnóstico:** `CapacidadDe` sumaba ocurrencias de días en calendario asumiendo 1 sesión por día. En aulas con `HorasDia = 8` y asignaturas de 4h (pregrado con franjas mañana y tarde), un aula puede alojar 2 sesiones por día, por lo que la capacidad real se subestimaba al 50%.
- **Solución Propuesta:** Calcular `sesionesPorDia = au != null ? Math.Max(1, (int)Math.Round(au.HorasDia / unidadAu)) : 1` dentro de `CapacidadDe` y multiplicar las ocurrencias del día K por este factor.

---

### 157.143 Ronda 187 — Acotamiento de Capacidad Diaria por Franjas Reales de Catálogo (Fase 982)

#### 1. Fase 982 — `sesionesPorDia` apuesta 2 sesiones/día sin verificar franjas AM/PM existentes
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorVentana.cs`
- **Diagnóstico:** `sesionesPorDia` asumía que la jornada del aula podía dividirse en sesiones de duración `unidadAu` sin verificar si el catálogo de bloques (`contexto.Bloques`) contenía efectivamente múltiples franjas activas para esa duración. Si un aula de 8h/día solo disponía de una franja activa de 4h (ej. solo MAÑANA), se le atribuía el doble de capacidad en pre-vuelo, difiriendo el conflicto al fitness y generando individuos con sesiones imposibles de acomodar.
- **Solución Propuesta:** Limitar `sesionesPorDia` con `Math.Min(..., Math.Max(1, franjasReales))`, contabilizando el número de franjas activas y distintas con `DuracionHoras == unidadAu` registradas en `contexto.Bloques`.

---

### 157.144 Ronda 188 — Unidad de Sesión en Patrón Semanal y Overload Funcional de Días K (Fases 983 y 984)

#### 1. Fase 983 — `Construir` divide por 4.0 fijo → posgrado 8 h declara patrón infactible
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorPatronSemanal.cs` y `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`
- **Diagnóstico:** `Construir` utilizaba un divisor hardcodeado de 4.0h para calcular el número de sesiones requeridas $S = \lceil totalHoras / 4.0 \rceil$. En programas de posgrado (asignatura de 64h con sesiones de 8h), proyectaba 16 sesiones en lugar de 8, declarando infactible el patrón cuando la ventana contenía menos de 16 semanas hábiles, cancelando el patrón semanal y degradando la asignación a un solo día.
- **Solución Propuesta:** Incorporar parámetro `double unidadSesion = 4.0` en `Construir` y calcular sesiones con `unidad = unidadSesion > 0 ? unidadSesion : 4.0`. En `AgFabricaIndividuos`, resolver la unidad efectiva según el aula/modalidad y pasarla a la invocación de `Construir`.

#### 2. Fase 984 — Sin overload `DiasKDe(diasPatron)` → carriles `[M,X]`/`[J,V]` inalcanzables
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorPatronSemanal.cs`
- **Diagnóstico:** `DiasKDe` solo admitía como entrada `AgAulaAsignatura` persistida. Patrones textuales declarados (semillas de UI como '1,3' o '4,5') no podían resolverse de forma funcional contra la ventana del aula, impidiendo la siembra y expresión del modelo de carriles alternos T5-T7.
- **Solución Propuesta:** Crear overload estático puro `DiasKDe(string diasPatron, string diasVentana = null, AgAula aula = null)` para procesar patrones textuales, intersectarlos contra la ventana del aula y recurrir al comportamiento canónico ante patrones vacíos o inválidos.

---

### 157.145 Ronda 189 — Reparación de Cruce: Preservación de Carril Descendiente y Propagación de Unidad Posgrado D4' (Fases 985 a 987)

#### 1. Fase 987 — Cruce reemplaza `dst.DiasPatron` con el patrón de otra aula → fusión de carriles
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgOperadorCruce.cs`
- **Diagnóstico:** `diasK` se derivaba incondicionalmente de `aaSrc` del progenitor, sobrescribiendo `dst.DiasPatron`. Si el descendiente pertenecía a un carril paralelo y el progenitor a otro, ambos carriles colapsaban en uno solo en oleadas T5-T7.
- **Solución Aplicada:** Declarada función local `AulaSrcDst(idAula)` e inicializado `diasK` con `AgPlanificadorPatronSemanal.DiasKDe(dst.DiasPatron, null, AulaSrcDst(dst.IdAula))` si `dst.DiasPatron` ya está presente; solo recurrir a `aaSrc` si está vacío.

#### 2. Fase 985 — Cruce reconstruye patrón con `unidad=4.0` → posgrado 8 h queda sin `PatronSemanal`
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgOperadorCruce.cs`
- **Diagnóstico:** Invocación de `_planificadorPatron.Construir` sin `unidadSesion` utilizaba el valor por defecto de 4.0h. En posgrados de 64h requiriendo 8 sesiones de 8h, calculaba $S = 16$ y fallaba en ventanas cortas, anulando `dst.PatronSemanal`.
- **Solución Aplicada:** Resolver `unidadCruce` dinámicamente mediante `AgUnidadesSemanales.HorasSesionAula(aulaDst, nivel)` desde el aula y modalidad del programa destino, pasándola como argumento a `Construir`.

#### 3. Fase 986 — Selección de bloque fija `DuracionHoras == 4.0` → posgrado recibe bloque de 4 h
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgOperadorCruce.cs`
- **Diagnóstico:** `bloquePatron` filtraba rígidamente por `b.DuracionHoras == 4.0`. En programas de posgrado con catálogo de 8h, fallaba la coincidencia y el fallback asignaba cualquier bloque del día (posiblemente de 4h), violando silenciosamente D4′.
- **Solución Aplicada:** Priorizar bloques cuya duración coincida con `unidadCruce` (`Math.Abs(b.DuracionHoras - unidadCruce) < 0.001`), con fallbacks sucesivos a 4.0h y a cualquier bloque del día.

---

### 157.146 Ronda 190 — Compatibilidad de Regla R8 y Columna K con Carriles y Posgrado 8h (Fases 988 y 989)

#### 1. Fase 988 — `R8_DiaFueraDePatron` usa `gr.First().DiasPatron` → todo carril secundario es violación dura
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorReglas.cs`
- **Diagnóstico:** La validación R8 utilizaba el patrón de días del primer gen del grupo (`gr.First().DiasPatron`) para contrastar todos los días del grupo. Si un grupo contenía múltiples carriles (oleadas alternas `[1,3]` y `[4,5]`), los días del carril secundario se marcaban fuera de patrón como infracción dura, exterminando la solución en el proceso genético.
- **Solución Aplicada:** Validar cada gen de forma individual contra su propio `DiasPatron` (`gr.Where(g => !string.IsNullOrWhiteSpace(g.DiasPatron))`), aislando la verificación a nivel de carril.

#### 2. Fase 989 — Columna K cae a 4 h cuando falta el bloque → posgrado reportado mal
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgMaterializadorSesiones.cs`
- **Diagnóstico:** Al componer `DiaHorasSemanales` en la vinculación, si no se resolvía el bloque horario en catálogo, la duración recurría a un fallback fijo de 4.0h, distorsionando las horas de asignaturas de posgrado (8h) en la columna K y en `RebuildSelectedDays`.
- **Solución Aplicada:** Extraer la duración de `subGrp.Select(x => x.HorasSemanales).Where(h => h > 0).DefaultIfEmpty(4.0).Max()` para reflejar fielmente la unidad real del gen cuando no haya bloque en catálogo.

---

### 157.147 Ronda 191 — Columna K Multi-Franja (AM/PM) y Contrato Canónico de 3 Partes (Fases 990 y 991)

#### 1. Fase 990 — Materializador elige `Min()` de bloque por día → franja reportada no es la programada
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgMaterializadorSesiones.cs`
- **Diagnóstico:** Las sesiones se agrupaban exclusivamente por `DiaSemana` tomando deterministamente `Min()` del bloque horario. En carriles con sesiones el mismo día en franjas distintas (ej. SÁBADO AM 08:00-12:00 y PM 14:00-18:00), el gen PM se reportaba erróneamente con horario AM al tener menor IdBloque, falseando la columna K y colapsando a una sola franja en `RebuildSelectedDays`.
- **Solución Aplicada:** Agrupación compuesta por `new { g.DiaSemana, g.IdBloqueHorario }` ordenada por día y bloque, emitiendo una entrada canónica individual por cada franja/bloque sin colapsar jornadas distintas ni emitir advertencias espurias.

#### 2. Fase 991 — `DiasHorasSemanaFormatter` emite 2 segmentos (inerte e incompatible con `RebuildSelectedDays`)
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/DiasHorasSemanaFormatter.cs`
- **Diagnóstico:** `DiasHorasSemanaFormatter` generaba claves de 2 segmentos (`"DIA|horario"`). Al no proveer el segmento intermedio de horas exigido por `RebuildSelectedDays` (`int.TryParse(segments[1])`), el parseo descartaba silenciosamente las entradas, dejando vacía la columna K al reabrir la sesión.
- **Solución Aplicada:** Adecuación al contrato canónico de 3 partes (`$"{dia}|{horas:0.##}|{horario}"`) con fallback de horas (`s.Horas > 0 ? s.Horas : 4.0`), asegurando compatibilidad round-trip con `RebuildSelectedDays` y preservando claves únicas por franja.

---

### 157.148 Ronda 192 — Erradicación del Enmascaramiento de Fechas con `DateTime.MinValue` (Fases 992 y 993)

#### 1. Fase 992 — Sesión sin fecha se persiste como `0001-01-01` en `ag_sesiones`
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgScheduleRepository.cs`
- **Diagnóstico:** En `GuardarStagingAsync`, si una sesión carecía de fecha, el fallback `(s.FechaSesion ?? DateTime.MinValue).ToString("yyyy-MM-dd")` persistía `"0001-01-01"` en la columna obligatoria `fechasesion TEXT NOT NULL`. Esto provocaba que `AcademicDateSyncService` la ignorara (filtro `> 2000-01-01`), `FormVerVinculaciones` mostrara la fecha ficticia `01/01/0001` y los ordenamientos de viáticos se alteraran, ocultando el error de origen.
- **Solución Aplicada:** Inclusión de compuerta preventiva previa al bucle de persistencia (`sesiones.FirstOrDefault(x => !x.FechaSesion.HasValue || x.FechaSesion.Value.Year < 1900)`), abortando la transacción y retornando `Result.Failure` explicativo, y persistiendo directamente `s.FechaSesion.Value.ToString("yyyy-MM-dd")`.

#### 2. Fase 993 — Promoción a producción replica el centinela `DateTime.MinValue`
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPromotorVinculaciones.cs`
- **Diagnóstico:** Al promover de staging a la tabla operativa `sesiones`, se aplicaba `FechaSesion = ss.FechaSesion ?? DateTime.MinValue`, enmascarando sesiones inválidas y propagando fechas residuales a producción donde no existe aislamiento por corrida.
- **Solución Aplicada:** Reemplazo del fallback por expresión `throw` explícita (`ss.FechaSesion ?? throw new Exception(...)`), asegurando fallo fail-fast antes de contaminar la base productiva.

---

### 157.149 Ronda 193 — Prevención de Colapso de Franjas y Unificación del Contrato Canónico (Fases 994 a 998)

#### 1. Fase 994 — Día con dos franjas (AM/PM) colapsa al reabrir la vinculación
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/FormSesionesAcademicas.cs`
- **Diagnóstico:** Al reabrir sesiones en `FormSesionesAcademicas`, `RebuildSelectedDays` utilizaba el nombre del día como clave de diccionario sin contemplar múltiples franjas en la misma jornada (ej. SÁBADO AM y PM). La segunda franja sobrescribía a la primera y al persistir se eliminaba silenciosamente la mitad del horario del carril.
- **Solución Aplicada:** En `RebuildSelectedDays`, al detectar una entrada previa para el mismo día, se emite advertencia en `LoggerHelper.Warning` y se conserva deterministamente la franja con mayor carga horaria, impidiendo el descarte ciego de la sesión principal.

#### 2. Fase 995 — `int.TryParse` descarta horas no enteras y formato invariante
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/FormSesionesAcademicas.cs`, `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgMaterializadorSesiones.cs`, `ConversorPDF.XLL/Academic/Application/Services/DiasHorasSemanaFormatter.cs`
- **Diagnóstico:** `int.TryParse(segments[1])` fallaba ante horas fraccionarias (`"3.5"` o `"2,5"`), omitiendo el día. Además, la interpolación `{horas:0.##}` con cultura local (es-CO) emitía comas decimales (`"7,5"`), que al dividirse con `Split(',')` fragmentaban corruptamente las cadenas de días.
- **Solución Aplicada:** Migración a `double.TryParse(..., NumberStyles.Float, CultureInfo.InvariantCulture)` y redondeo seguro en `RebuildSelectedDays`; serialización con `horas.ToString("0.##", CultureInfo.InvariantCulture)` en `AgMaterializadorSesiones` y `DiasHorasSemanaFormatter`.

#### 3. Fase 996 — `DiasHorasSemanaFormatter` emite abreviaturas; contrato canónico usa nombres completos
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/DiasHorasSemanaFormatter.cs`
- **Diagnóstico:** El formateador producía abreviaturas (`"LUN"`, `"MAR"`, `"MIE"`), incompatibles con el catálogo canónico del sistema (`"LUNES"`, `"MARTES"`, ...). Esto rompía `DayHoursDropdown._diasSemana`, `DayPatternParser` y `RebuildSelectedDays`.
- **Solución Aplicada:** Delegación de `NormalizarDia` y `ObtenerNombreDia` en `DiaSemanaNormalizador.ObtenerNombreDia(...)`, homologando el vocabulario a nombres completos sin tildes.

#### 4. Fase 997 — Separador `" / "` en `DesdeSesiones` rompe el round-trip
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/DiasHorasSemanaFormatter.cs`
- **Diagnóstico:** `DesdeSesiones` unía elementos con `" / "`, contraviniendo el estándar separado por comas `", "` esperado por `RebuildSelectedDays` y `DayHoursDropdown`.
- **Solución Aplicada:** Unificación del delimitador a `", "` y documentación de la deuda técnica para su posterior integración en exportadores y grillas.

#### 5. Fase 998 — IDs fantasma tras rollback de F992
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgScheduleRepository.cs`
- **Diagnóstico:** En `GuardarStagingAsync`, al abortar por sesión sin fecha válida, las instancias `AgVinculacion` en memoria retenían los `dbId` asignados previo al `trans.Rollback()`, dejando IDs inexistentes en el grafo de objetos.
- **Solución Aplicada:** Restauración de los IDs originales en memoria mapeando desde `mapVinc` antes de ejecutar el rollback y retornar `Result.Failure`.

---

### 157.150 Ronda 194 — Corrección de Especificador Militar `HH` y Prevención de Descarte Silencioso (Fases 999 a 1001)

#### 1. Fase 999 — `NormalizarHorario` formatea con `hh` (12 h) mutilando horarios PM a AM
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/HorarioMilitarizador.cs`, `ConversorPDF.XLL/Academic/Domain/Models/FranjaHoraria.cs`
- **Diagnóstico:** El formato interpolado `{f.Inicio:hh\\:mm} A {f.Fin:hh\\:mm}` y `{Inicio:hh\\:mm} - {Fin:hh\\:mm}` utilizaba el especificador de 12 horas `hh` sin designador meridiano en vez del estándar militar de 24 horas `HH`. Horarios vespertinos como `"13:00 A 17:00"` eran reescritos como `"01:00 A 05:00"`, provocando que `DayHoursDropdown` no los encontrara en su catálogo de opciones, asignara `DBNull.Value` y descartara permanentemente el día al guardar.
- **Solución Aplicada:** Corrección del especificador de formato a `HH\\:mm` tanto en `HorarioMilitarizador.NormalizarHorario` como en `FranjaHoraria.ToString()`, garantizando idempotencia en el rango 00:00 - 23:59.

#### 2. Fase 1000 — Datos normalizados previamente por V63 quedaron corruptos
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV71_AuditarHorariosMeridiano.cs`, `ConversorPDF.XLL/Core/Helpers/SqliteHelper.cs`, `ConversorPDF.XLL/ConversorPDF.XLL.csproj`
- **Diagnóstico:** La migración V63 unificó horarios utilizando `HorarioMilitarizador` cuando este contenía el defecto `hh`. Las filas persistidas en `sesiones.horariosesion`, `ag_sesiones.horariosesion` y `vinculaciones.dia_horas_semanales` no pueden corregirse automáticamente de forma determinista por ambigüedad intrínseca AM/PM.
- **Solución Aplicada:** Creación de la migración no destructiva `MigrationV71_AuditarHorariosMeridiano` (versión 71), la cual audita registros afectados mediante consultas de solo lectura y emite alertas diagnósticas en `LoggerHelper.Warning` para orientar la regeneración de staging y auditoría de producción.

#### 3. Fase 1001 — `DayHoursDropdown` descarta el día sin emitir traza diagnóstica
- **Archivo Afectado:** `ConversorPDF.XLL/UI/DayHoursDropdown.cs`
- **Diagnóstico:** Cuando un horario normalizado no coincidía con ninguna de las opciones predefinidas en el ComboBox de la celda, se asignaba silenciosamente `DBNull.Value`, provocando que al guardar la fila el día se eliminara sin dejar evidencia en logs.
- **Solución Aplicada:** Inclusión de registro de advertencia con `LoggerHelper.Warning` en el bloque `else` de resolución de horario en `DayHoursDropdown.CargarDesdeDiaHorasSemanales`, detallando el horario ausente, el día y las opciones disponibles antes de asignar `DBNull.Value`.

---

### 157.151 Ronda 195 — Fugas de Subunidad, Preservación de Carriles Concurrentes y Afinamiento de Auditoría (Fases 1002 a 1006)

#### 1. Fase 1002 — `GetBloquesHorariosAsync` ignora el parámetro `subunidad`
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgScheduleRepository.cs`
- **Diagnóstico:** El método recibía el parámetro `subunidad` pero la consulta SQL filtraba exclusivamente por `activo = 1`, permitiendo que bloques de otras subunidades se mezclaran en el plan y alteraran la asignación de franjas por subunidad.
- **Solución Aplicada:** Incorporada condición `@Subunidad = '' OR COALESCE(subunidad, '') = @Subunidad` en la cláusula `WHERE`, pasando el parámetro mediante `CommandDefinition`.

#### 2. Fase 1003 — Carriles concurrentes colapsados por `GroupBy(...).Select(g => g.First())`
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`
- **Diagnóstico:** La agrupación por tupla `(AgClaveCohorte, Jornada, DiasVentana)` reducía múltiples aulas con la misma cohorte a una sola instancia representativa. En patrones alternos (ej. MAR-VIE con carriles `[M,X]` y `[J,V]`), el segundo par de días era descartado y no se planificaba.
- **Solución Aplicada:** Deduplicación por `a.Id` en lugar de por atributos de cohorte, conservando la totalidad de carriles concurrentes de la misma cohorte.

#### 3. Fase 1004 — Comparación exacta `DuracionHoras == 4.0` sobre columna REAL
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`
- **Diagnóstico:** La comparación de igualdad estricta `b.DuracionHoras == 4.0` en tipos de punto flotante (`double`) fallaba ante imprecisiones numéricas (`3.9999999` o `4.0000001`) provenientes de la columna `REAL` en base de datos, provocando caídas no deseadas al fallback de 8 horas o a bloques arbitrarios.
- **Solución Aplicada:** Sustituida la igualdad estricta por comparación con margen de tolerancia: `Math.Abs(b.DuracionHoras - 4.0) < 0.001` en las líneas 292 y 364.

#### 4. Fase 1005 — Auditoría V71 con criterio demasiado amplio (falsos positivos)
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV71_AuditarHorariosMeridiano.cs`
- **Diagnóstico:** El filtro `LIKE '% A %'` contabilizaba la totalidad de los horarios que contuvieran el separador `" A "`, incluyendo registros canónicos y válidos, inflando el recuento y perdiendo precisión diagnóstica.
- **Solución Aplicada:** Restringida la consulta diagnóstica para contar únicamente registros no vacíos cuyo `horariosesion` no exista en el catálogo `ag_bloque_horario` (`hora_inicio || ' A ' || hora_fin`), detectando exactamente aquellos alterados o desalineados.

#### 5. Fase 1006 — Semilla V62 introduce bloques de domingo y variantes 8 h duplicadas
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV62_AddAgAulaVentanasYBloquesSemilla.cs`
- **Diagnóstico:** El bucle de inserción iteraba de 1 a 7, sembrando bloques para domingo (día 7), e incluía una variante redundante de 8 horas para `NOCHE` (`14:00-22:00`) que colisionaba con `TARDE`.
- **Solución Aplicada:** Ajustado el límite del bucle a `dia <= 6` para excluir domingos de la semilla y eliminada la plantilla duplicada de 8 horas en `NOCHE`.

---

### 157.152 Ronda 196 — Preservación de Bloques Globales, Repartición de Carriles y Depuración de Semilla (Fases 1007 a 1009)

#### 1. Fase 1007 — F1002 excluye los bloques globales (`subunidad = ''`) cuando hay subunidad activa
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgScheduleRepository.cs`
- **Diagnóstico:** La cláusula `WHERE (@Subunidad = '' OR COALESCE(subunidad, '') = @Subunidad)` introducida en F1002 excluía todos los bloques base/globales sembrados con `subunidad = ''` siempre que el usuario seleccionara una subunidad específica en el generador, resultando en 0 bloques horarios y provocando vacantes masivas o fallos de ejecución.
- **Solución Aplicada:** Modificada la cláusula a `@Subunidad = '' OR COALESCE(subunidad, '') IN ('', @Subunidad)` y ordenado con `ORDER BY CASE WHEN COALESCE(subunidad, '') = @Subunidad THEN 0 ELSE 1 END, dia_semana, hora_inicio` para dar prioridad a los bloques especializados de la subunidad sin ignorar los bloques globales.

#### 2. Fase 1008 — `DiasKDe` trunca toda ventana a 2 días → un solo carril
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorPatronSemanal.cs`, `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`
- **Diagnóstico:** `DiasKDe` ejecutaba `Take(2)` de manera fija sobre `diasVentana`, de modo que ventanas de 4 días como `MAR-VIE` (`"2,3,4,5"`) siempre retornaban los dos primeros días (`[2,3]`). Esto ocasionaba que todas las asignaturas de una cohorte colisionaran en los mismos días y el carril complementario (`[4,5]`) nunca fuera asignado.
- **Solución Aplicada:** Incorporado parámetro opcional `int carril = 0` en `DiasKDe` calculando el offset dinámico (`Skip(idx * 2).Take(2)`), e instrumentado el contador `idxAsig` por aula en `AgFabricaIndividuos` para alternar secuencialmente los carriles entre las asignaturas programadas.

#### 3. Fase 1009 — Semilla de domingo ya persistida y variantes 8 h duplicadas
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV72_LimpiarBloquesDomingoYDuplicados.cs`, `ConversorPDF.XLL/Core/Helpers/SqliteHelper.cs`, `ConversorPDF.XLL/ConversorPDF.XLL.csproj`
- **Diagnóstico:** La corrección en el código de la migración V62 (F1006) solo aplica a bases nuevas; bases de datos existentes ya ejecutaron V62 y retienen bloques para `dia_semana = 7` (domingo) y bloques duplicados de 8 horas en `NOCHE` (`14:00-22:00`).
- **Solución Aplicada:** Creación y registro de la migración no destructiva `MigrationV72_LimpiarBloquesDomingoYDuplicados` (versión 72), eliminando de `ag_bloque_horario` los registros globales de domingo y las variantes de 8h en `NOCHE`.

---

### 157.153 Ronda 197 — Intersección de Patrón con Ventana y Soporte de Ventanas Impares (Fases 1010 a 1012)

#### 1. Fase 1010 — `DiasKDe(aa, aula, …)` no intersecta `aa.DiasPatron` con la ventana del aula
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorPatronSemanal.cs`
- **Diagnóstico:** Cuando `aa.DiasPatron` contenía días sembrados por carga horaria (ej. 64 h → `"4,5,6"` o 80 h → `"2,3,4,5,6"`), `DiasKDe` retornaba esos días de forma directa sin intersectar con `aula.DiasVentana`. Si el aula tenía ventana `MAR-VIE` (`[2,3,4,5]`), se programaban sesiones en sábado (`6`) y se saltaba la partición en carriles de F1008.
- **Solución Aplicada:** Añadida intersección explícita con `AgDiasVentana.Dias(aula.DiasVentana)` dentro de la rama `!string.IsNullOrWhiteSpace(aa.DiasPatron)`, asegurando paridad funcional con el overload textual de `DiasKDe`.

#### 2. Fase 1011 — Ventanas impares pierden el último día (`carriles` truncado)
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorPatronSemanal.cs`
- **Diagnóstico:** El cálculo de carriles utilizaba división entera (`diasVentana.Count / porCarril`). Para ventanas de 3 días como `JUE-SAB` (`4,5,6`) o `VIE-DOM` (`5,6,7`), el resultado era `3 / 2 = 1`, dejando un único carril `[4,5]` y omitiendo permanentemente el día sábado (`6`).
- **Solución Aplicada:** Modificado el cálculo de cantidad de carriles a `(int)Math.Ceiling(diasVentana.Count / (double)porCarril)`, permitiendo que ventanas impares generen un carril complementario (ej. `[4,5]` y `[6]`).

#### 3. Fase 1012 — Validador de capacidad calcula con carril 0 para todos los carriles
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorVentana.cs`
- **Diagnóstico:** Las comprobaciones de factibilidad en `CapacidadDe` y validación D14 invocaban `DiasKDe` sin parámetro de carril, asumiendo implícitamente el carril 0 (de 2 días).
- **Solución Aplicada:** Explicitada la llamada a `DiasKDe(..., carrilGen)` con carril representativo en pre-vuelo y documentada la relación entre los carriles y la capacidad efectiva por curso.

---

### 157.154 Ronda 198 — Sembrado de Patrón por Carril y Advertencia Multicarril en Pre-Vuelo (Fases 1013 a 1014)

#### 1. Fase 1013 — `dias_patron` sembrado desde horas anula F1008 y excede 2 días
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** El sembrado automático en `GenerarEncabezadosAulasAsync` escribía `dias_patron` = `"2,3,4,5,6"` (≥80 h) o `"4,5,6"` (≥64 h). Al quedar `AgAulaAsignatura.DiasPatron` no vacío, `DiasKDe(aa, aula, carril)` retornaba por la rama de patrón existente e ignoraba el parámetro `carril`, provocando que todas las asignaturas de la cohorte compartieran los mismos días (un solo carril real) y distribuyeran la carga en 3 a 5 días con bloques atípicos de 1 a 2 horas.
- **Solución Aplicada:** Instrumentado contador `idxAulaAsig` por aula y asignado `diasPatron` de a 2 días rotado por asignatura (`idxAulaAsig % carriles`) extrayendo de `AgDiasVentana.Dias(item.DiasVentana)` mediante `Skip(lane * 2).Take(2)`.

#### 2. Fase 1014 — F1012 introdujo variable inerte; capacidad sigue calculada con carril 0
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorVentana.cs`
- **Diagnóstico:** `int carrilGen = 0;` en `CapacidadDe` era constante, manteniendo el cálculo de capacidad acotado silenciosamente al carril 0 sin advertir que en aulas multicarril con carriles residuales de longitud impar (ej. `[6]`) la capacidad calculada sobreestima la oferta real del carril complementario.
- **Solución Aplicada:** Reemplazado el código inerte por la determinación de `carriles` sobre `AgDiasVentana.Dias(au.DiasVentana)` y la emisión de advertencia diagnóstica no redundante mediante `LoggerHelper.WarningLimitado` tanto en `CapacidadDe` como en la validación D14 de configuraciones de aula.

---

### 157.155 Ronda 199 — Depuración de Patrones Legacy Multidía y Determinismo en Reparto de Carriles (Fases 1015 a 1016)

#### 1. Fase 1015 — `dias_patron` legacy (3-5 días) persiste y anula F1013
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV73_LimpiarDiasPatronMultidia.cs`, `ConversorPDF.XLL/Core/Helpers/SqliteHelper.cs`, `ConversorPDF.XLL/ConversorPDF.XLL.csproj`
- **Diagnóstico:** Filas preexistentes en `ag_aula_asignatura` sembradas bajo versiones anteriores conservaban valores como `"2,3,4,5,6"` o `"4,5,6"`. La instrucción `INSERT OR IGNORE` en `GenerarEncabezadosAulasAsync` no actualizaba estas filas y reutilizaba el valor legacy no vacío, forzando a `DiasKDe` a eludir la partición de carriles de F1008/F1013 y fragmentando las horas en bloques de 1 a 2 horas.
- **Solución Aplicada:** Implementada y registrada la migración `MigrationV73_LimpiarDiasPatronMultidia` (versión 73) para resetear a `''` cualquier `dias_patron` con más de 2 días en `ag_aula_asignatura`, permitiendo que el motor genético gobierne los días de clase por ventana y carril.

#### 2. Fase 1016 — Carril asignado sin `ORDER BY` → reparto no reproducible
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** La consulta `sqlAsigs` carecía de cláusula de ordenamiento, provocando que `idxAulaAsig % carriles` dependiera del orden arbitrario de iteración retornado por SQLite, introduciendo no determinismo entre ejecuciones y rompiendo la auditabilidad de las corridas.
- **Solución Aplicada:** Añadida cláusula `ORDER BY a.id;` a `sqlAsigs` para garantizar un orden determinista y reproducible en la rotación de asignaturas por carril.

---

### 157.156 Ronda 200 — Preservación de 2 Días en Carril Residual y Fallback de Patrón sobre Ventana (Fases 1017 a 1018)

#### 1. Fase 1017 — Carril residual de 1 día reduce capacidad frente a la calculada
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorPatronSemanal.cs`, `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** En ventanas impares (ej. `JUE-SAB` = `[4,5,6]`), el carril residual 1 tomaba un único día (`[6]`) con la mitad de capacidad que el carril 0 (`[4,5]`). Como el validador pre-vuelo evalúa con el carril 0, las asignaturas rotadas al carril residual no alcanzaban a cubrir las horas requeridas y degradaban el fitness sin bloquear.
- **Solución Aplicada:** Calculado el inicio de ventana mediante desplazamiento deslizante `Math.Min(idx * porCarril, Math.Max(0, diasVentana.Count - porCarril))`, garantizando que el carril residual siempre conserve 2 días (ej. `[5,6]` para `JUE-SAB`), replicado en el sembrado de `AgFormGeneradorHorarios.cs`.

#### 2. Fase 1018 — Patrón fuera de ventana se devuelve sin filtrar
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorPatronSemanal.cs`
- **Diagnóstico:** Cuando `aa.DiasPatron` no intersectaba con `aula.DiasVentana` (ej. patrón `[2,3]` y ventana cambiada a `JUE-SAB`), `dentro.Count == 0` retornaba la lista original sin filtrar, programando sesiones en días no autorizados fuera de la ventana del aula.
- **Solución Aplicada:** Añadido fallback en `DiasKDe` que emite `LoggerHelper.WarningLimitado` y resuelve el carril de 2 días directamente sobre la ventana del aula cuando el patrón está íntegramente fuera de ella.

---

### 157.157 Ronda 201 — Paridad en Overload Textual y Carriles Disjuntos contra Infracción R4 (Fases 1019 a 1020)

#### 1. Fase 1019 — Overload textual sin paridad F1018 → días fuera de ventana en cruce
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorPatronSemanal.cs`, `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgOperadorCruce.cs`
- **Diagnóstico:** El overload textual `DiasKDe(string, string, AgAula)` retornaba `pedidos` crudo cuando `dentro.Count == 0`, permitiendo que en el operador de cruce (`AgOperadorCruce.cs:166`) se propagaran días ajenos a la ventana del aula destino, eludiendo el fallback seguro introducido en F1018. Adicionalmente, la rama `aaSrc` en cruce invocaba `DiasKDe` sin pasar el aula destino.
- **Solución Aplicada:** Incorporado parámetro opcional `int carril = 0` y fallback con `LoggerHelper.WarningLimitado` delegando en `DiasKDe((AgAulaAsignatura)null, aula, carril)` ante patrones fuera de ventana, y pasado `AulaSrcDst(dst.IdAula)` en `AgOperadorCruce.cs:168`.

#### 2. Fase 1020 — Carriles solapados violan R4 (misma aula/grupo, mismo día)
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorPatronSemanal.cs`, `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** En ventanas impares (ej. `JUE-SAB` = `[4,5,6]`), el clamp deslizante de F1017 generaba carriles solapados `[4,5]` y `[5,6]`, haciendo que dos asignaturas de la misma cohorte coincidieran el día 5 y provocando violaciones duras sistemáticas de R4 (`R4_DosAsignaturasMismoDia`) y de aula R2.
- **Solución Aplicada:** Reemplazado el cálculo de carriles por división entera estricta (`diasVentana.Count / porCarril`) asegurando carriles disjuntos, dejando el día residual fuera de rotación con advertencia diagnóstica no redundante (`LoggerHelper.WarningLimitado`), y espejado en el sembrado de `AgFormGeneradorHorarios.cs`.

---

### 157.158 Ronda 202 — Tolerancia a ISO Numérico en Ventana y Diferimiento de Patrón sin Ventana (Fases 1021 a 1022)

#### 1. Fase 1021 — `AgDiasVentana.Dias` sin tolerancia a ISO numérico
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgDiasVentana.cs`
- **Diagnóstico:** `AgDiasVentana.AIso` únicamente validaba contra las 5 claves canónicas del diccionario (ej. `"MAR-VIE"`). Cuando el valor de `dias_ventana` ya venía persistido como ISO numérico (ej. `"1,2,3,4"` o `"2,3,4,5"`), `Opciones.TryGetValue` fallaba retornando cadena vacía, lo que provocaba que `Dias` devolviera lista vacía y `DiasKDe` recurriera al default `{2,3}` ignorando la ventana real del aula.
- **Solución Aplicada:** Añadido soporte en `AIso` para parsear listas de enteros delimitadas por coma o punto y coma en el rango ISO [1..7] si la búsqueda por clave falla, retornando el formato normalizado ordenado y sin duplicados.

#### 2. Fase 1022 — Sembrado con ventana vacía fuerza `4,5`/`2,3`
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** En la inicialización de `ag_aula_asignatura`, cuando el aula no tenía `item.DiasVentana` declarado, el fallback forzaba patrones estáticos `"4,5"` o `"2,3"` según la intensidad horaria, lo que restringía rígidamente los días de cursado sin considerar la jornada o restricciones de la sede antes de la optimización genética.
- **Solución Aplicada:** Asignado `diasPatron = string.Empty` en la rama sin ventana declarada y registrada traza diagnóstica `AulaAsignaturaSinVentana`, permitiendo que el motor genético y `DiasKDe` deriven los días en función de las restricciones reales del aula.

---

### 157.159 Ronda 203 — Colisión de nombre Excel 'De' (Fases 1033 a 1038)

#### 1. Fase 1033 — Eliminar el nombre Excel 'De' en AgSemanaIso y AgClaveCohorte
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgSemanaIso.cs`, `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgClaveCohorte.cs`, `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`
- **Diagnóstico:** Excel-DNA 1.8.0 registra por defecto todos los métodos públicos estáticos de clases públicas cuyos parámetros y retornos pertenecen a tipos soportados. El nombre Excel se deriva exclusivamente del nombre del método sin prefijo de clase. `AgSemanaIso.De(DateTime)` y `AgClaveCohorte.De(int, int, int, string, string, string)` colisionaban registrando ambas la función `=DE(...)`, provocando el mensaje `Repeated function name: 'De' - previous registration will be overwritten.` en el log de registro.
- **Solución Aplicada:**
  1. Se eliminó el alias `De(DateTime)` en `AgSemanaIso.cs` preservando `DiaIso(DateTime)` como método canónico.
  2. Se renombraron las sobrecargas en `AgClaveCohorte.cs` a `DesdeCohorte`, `DesdeAula` y `DesdeGen`.
  3. Se actualizó el call site en `AgFabricaIndividuos.cs:84` a `AgClaveCohorte.DesdeAula(a)`.

#### 2. Fase 1034 — Test guardián de registro de métodos Excel-DNA
- **Archivo Afectado:** `ConversorPDF.Tests/Unit/ExcelDnaRegistrationGuardTests.cs`
- **Diagnóstico:** Ausencia de prueba automatizada de regresión que verifique las reglas de descubrimiento y registro de Excel-DNA 1.8.0 sobre el ensamblado `ConversorPDF.XLL`, permitiendo que futuros métodos estáticos públicos con nombres idénticos o con el identificador colisionado `De` vuelvan a ingresar al código base.
- **Solución Aplicada:** Implementado test guardián con xUnit y FluentAssertions que inspecciona por reflexión todos los tipos públicos del ensamblado de `ConversorPDF.XLL`, validando que no existan nombres duplicados (`OrdinalIgnoreCase`) y que ningún método registrable se denomine `De`.

#### 3. Fase 1036 — `DesdeCohorte` queda expuesto como UDF inválido (6 argumentos requeridos)
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgClaveCohorte.cs`
- **Diagnóstico:** `DesdeCohorte(int, int, int, string, string aulaRef = null)` es público y estático en clase pública con tipos compatibles. Excel-DNA 1.8.0 no soporta valores por defecto (`ParameterInfo.IsOptional` no se evalúa) y registra la función exigiendo todos los parámetros en Excel, fallando las invocaciones documentadas con 4/5 parámetros.
- **Solución Aplicada:** Modificada la visibilidad de `DesdeCohorte` a `internal static string DesdeCohorte(...)`, restringiéndolo a key-builder de dominio y excluyéndolo del registro UDF de Excel.

#### 4. Fase 1037 — El test guardián no replica las reglas de Excel-DNA 1.8.0
- **Archivo Afectado:** `ConversorPDF.Tests/Unit/ExcelDnaRegistrationGuardTests.cs`
- **Diagnóstico:** El test guardián inicial filtraba `IsSpecialName`, omitiendo accesores de propiedades `get_`/`set_` estáticos públicos que Excel-DNA 1.8.0 sí registra (7 métodos en Release). Asimismo, omitía los filtros de tipos de Excel-DNA (`type.IsArray`, genéricos abiertos y namespace `"My"`).
- **Solución Aplicada:** Alineadas las reglas en `EsMetodoRegistrableExcelDna180` y `ObtenerMetodosRegistrables` con `AssemblyLoader.GetExcelMethods` y `IsMethodSupported` de Excel-DNA 1.8.0 (removido `IsSpecialName`, agregada exclusión de métodos abstractos/genéricos, arrays, tipos genéricos con parámetros abiertos y namespace `"My"`).

#### 5. Fase 1038 — Trazabilidad y datos incorrectos en documentación
- **Archivos Afectados:** `implementation_plan.md`, `walkthrough.md`, `docs/articles/changes/2026-10-22-colision-nombre-funcion-de.md`
- **Diagnóstico:** Desincronización en la raíz del repositorio omitiendo la Ronda 203; cita incorrecta en artículo alegando excepción en lugar del log real de sobrescritura `Repeated function name: 'De'`; omisión de la visibilidad interna de `DesdeCohorte` y falta de advertencia de error `#NAME?` para fórmulas que invocaban `=De(...)`.
- **Solución Aplicada:** Sincronizados `implementation_plan.md` y `walkthrough.md` en la raíz con la Ronda 203 (Fases 1033 a 1038); corregido el artículo de cambios con los textos y advertencias de compatibilidad precisos.

---

### 157.160 Ronda 204 — Corrección de Solapamientos Espaciales y Extensión de `treeContextoAcademico` (Fases 1039 a 1044)

#### 1. Fase 1039 — Splitter Horizontal y Reorganización de `tabSimulacion`
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.Designer.cs`, `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** El contenedor `tabSimulacion` acumulaba controles apilados con `Dock=Bottom` fijos (82 + 58 + 118 = 258 px), comprimiendo el `grpSeleccion` a 357 px y dejando el árbol en solo 138 px con ~105 px de vacío gris inútil sin posibilidad de ajuste manual por parte del usuario.
- **Solución Aplicada:**
  1. Se introdujo `splitContenedorSim` (`SplitContainer`, `Orientation=Horizontal`, `Dock=Fill`, `FixedPanel=Panel2`, `SplitterWidth=6`) como único hijo de `tabSimulacion`.
  2. En `Panel1` se alojan `grpSeleccion` (`Dock=Fill`, `MinimumSize=(0,240)`) y `pnlConfiguracionInferior` (`Dock=Bottom`).
  3. En `Panel2` se alojan `pnlBotonesSim` (`Dock=Top`) y `txtLogAG` (`Dock=Fill`).
  4. En `OnLoad`, se calibra dinámicamente `SplitterDistance = Math.Max(280, Height - 300)` otorgando al árbol ≥ 240 px iniciales expandibles hasta ~345 px mediante el divisor.

#### 2. Fase 1040 — `treeContextoAcademico` con `Dock=Fill` y Toolbar Desacoplada sin Colisiones
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.Designer.cs`, `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** `treeContextoAcademico` utilizaba `Anchor` cuádruple con coordenadas fijas (10, 112, 1072, 138), perdiendo altura relativa. En `pnlArbolToolbar`, `lblConteoCombinaciones` en X=738 con `AutoSize=true` crecía hasta ~928 px solapando por 39 px a `btnGenerarAulas` (X=889) al pintarse encima por z-order.
- **Solución Aplicada:**
  1. `treeContextoAcademico` pasa a `Dock=Fill` con `BorderStyle=FixedSingle`, `HideSelection=false`, `ShowNodeToolTips=true`, `ItemHeight=20` y `Scrollable=true`.
  2. En `pnlArbolToolbar` (alto 34, `Dock=Top`), los botones de perfil se agruparon en `pnlToolbarIzq` (`FlowLayoutPanel`, `Dock=Left`, `WrapContents=false`, `AutoSize=true`).
  3. `btnGenerarAulas` se fijó con `Dock=Right` (150 px).
  4. `lblConteoCombinaciones` se configuró con `Dock=Fill`, `AutoSize=false`, `AutoEllipsis=true`, `TextAlign=MiddleLeft` y `toolTipConteo` para visualización completa sin colisión.

#### 3. Fase 1041 — Cabeceras Nivel 6 Desacopladas y Cálculo Proporcional Unificado
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.Designer.cs`, `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** Las cabeceras de Nivel 6 (`lblColJornada`, `lblColDias`, `lblColHorasDia`) residían en la segunda fila de `pnlArbolToolbar` con Y=36 y X fijas (253, 478, 893), permanecían siempre visibles aún sin nodos de modalidad desplegados, y no coincidían con los 11 offsets mágicos hardcodeados en `DrawNode`, `ProcesarClickDimensionNivel6` y `MouseMove`.
- **Solución Aplicada:**
  1. Se extrajeron las cabeceras a un panel independiente `pnlCabecerasNivel6` (`Dock=Top`, alto 20, `Visible=false`) con separadores verticales de 1 px (`sepNivel6A`, `sepNivel6B`).
  2. Se creó el struct `AgColumnasNivel6` y método `CalcularColumnasNivel6()` como fuente única de verdad con reparto proporcional (Jornada 30%, Días 48%, Horas 22%) sobre el ancho disponible y anchos mínimos garantizados (62, 74, 64 px).
  3. Se unificaron `DrawNode`, `ProcesarClickDimensionNivel6`, `MouseMove` y `ActualizarCabecerasNivel6()` consumiendo el struct, fijando `pnlCabecerasNivel6.Visible = HayNodoNivel6Visible()`.

#### 4. Fase 1042 — Reorganización de Barra de Acciones y Depuración de Emojis
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.Designer.cs`, `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** `pnlAccionesSim` tenía ancho rígido de 700 px dejando 89 px muertos y solo 7 px entre `chkModoDebug` y los botones. Con fuentes del sistema sin glifo emoji (`🔍`, `▶`) en Segoe UI 9 o bajo DPI 125 %, los textos se desbordaban generando solapes visuales sobre la barra de progreso.
- **Solución Aplicada:**
  1. En `pnlBotonesSim` (`Dock=Top` en `Panel2`, alto 58), se creó `pnlProgreso` (`Dock=Left`, ancho 380) conteniendo `progressBarAG` y `lblStatusAG` (`AutoSize=false`, `AutoEllipsis=true`).
  2. `chkModoDebug` se fijó con `Dock=Right` (180 px, `AutoEllipsis=true`).
  3. Las acciones se alojaron en `flowAcciones` (`FlowLayoutPanel`, `Dock=Right`, `FlowDirection=RightToLeft`, `AutoSize=true`).
  4. Se normalizaron los textos eliminando emojis: `"Diagnóstico Pre-vuelo"`, `"Optimizar Horarios"`, `"Cancelar"`.

#### 5. Fase 1043 — Estandarización Tipográfica, DPI Base y Suscripciones de Redimensión
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.Designer.cs`, `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** El formulario carecía de propiedad `Font` explícita heredando MS Sans Serif 8.25 mientras los controles usaban Segoe UI 9, provocando distorsión en la métrica `AutoScaleDimensions = (6F, 13F)` al escalar a 125%/150% DPI. Asimismo, colapsar nodos o mover el splitter no refrescaba las cabeceras ni el cursor.
- **Solución Aplicada:**
  1. Fijado en el formulario `Font = new Font("Segoe UI", 9F)`, `AutoScaleDimensions = (7F, 15F)`, `ClientSize = (1120, 720)` y `MinimumSize = (1120, 700)`.
  2. Suscritos eventos `treeContextoAcademico.AfterCollapse`, `treeContextoAcademico.ClientSizeChanged` y `splitContenedorSim.SplitterMoved` a `ActualizarCabecerasNivel6()`.
  3. `ObtenerXBaseNivel6` enriquecido con salvaguarda `Math.Max(..., ClientSize.Width * 0.09)` para pantallas de alta densidad.

#### 6. Fase 1044 — Sincronización Documental y Trazabilidad en Base de Conocimiento
- **Archivos Afectados:** `docs/articles/changes/2026-ag-horarios-academicos.md`, `implementation_plan.md`, `walkthrough.md`
- **Diagnóstico:** Necesidad de reflejar las mejoras de arquitectura espacial de UI en la documentación oficial del motor genético (§8.2) y sincronizar la trazabilidad de la Ronda 204 en los registros de cambios del repositorio.
- **Solución Aplicada:**
  1. Actualizada la subsección §8.2 de `docs/articles/changes/2026-ag-horarios-academicos.md` documentando el modelo de docking responsivo, el SplitContainer y la distribución paramétrica de Nivel 6.
  2. Consolidadas las Fases 1039 a 1044 en `implementation_plan.md` y `walkthrough.md`.

---

### 157.161 Ronda 205 — Corrección de Contenedor de Componentes Nulo, Calibración de Splitter, Anti-Reentrada y Separadores (Fases 1045 a 1051)

#### 1. Fase 1045 — Instanciación de `components` en `InitializeComponent`
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.Designer.cs`
- **Diagnóstico:** `this.toolTipConteo = new System.Windows.Forms.ToolTip(this.components);` con `components == null` lanzaba `ArgumentNullException` en tiempo de ejecución al construir el formulario, impidiendo su apertura y bloqueando el Módulo S.
- **Solución Aplicada:** Inicializado `this.components = new System.ComponentModel.Container();` al inicio de `InitializeComponent()`, asegurando la liberación adecuada mediante `Dispose()`.

#### 2. Fase 1046 — Calibración de Splitter y Corrección de `grpSeleccion.MinimumSize`
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`, `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.Designer.cs`
- **Diagnóstico:** `Height - 300` reservaba espacio excesivo para el panel inferior (294 px) dejando el árbol en solo 174 px útiles; asimismo, `grpSeleccion.MinimumSize = (0, 240)` provocaba que a tamaño mínimo (1120x700) el grupo desbordara 7 px sobre `pnlConfiguracionInferior` reintroduciendo el solape.
- **Solución Aplicada:**
  1. Configurado `splitContenedorSim.Panel1MinSize = 300` y `splitContenedorSim.Panel2MinSize = 120`.
  2. Calibrado `SplitterDistance = alto - altoPanel2 - anchoSeparador` (objetivo: 229 px útiles para el árbol).
  3. Reducido en Designer `grpSeleccion.MinimumSize = new Size(0, 190)`.

#### 3. Fase 1047 — Protección Anti-Reentrada en `ActualizarCabecerasNivel6`
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** Asignar `pnlCabecerasNivel6.Visible` alteraba el alto disponible del árbol disparando `ClientSizeChanged`, lo que reentraba en `ActualizarCabecerasNivel6` durante el ciclo de layout produciendo parpadeo visible y ejecuciones redundantes.
- **Solución Aplicada:** Incorporada bandera de guardia `private bool _alineandoCabeceras;` con bloque `try/finally` y asignación condicionada `if (pnlCabecerasNivel6.Visible != hayVisible)`.

#### 4. Fase 1048 — Unificación a Recorrido Único de Búsqueda de Nivel 6
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** `ActualizarCabecerasNivel6` ejecutaba dos recorridos recursivos O(n) por cada redimensión del control (`HayNodoNivel6Visible` y `ObtenerXBaseNivel6`).
- **Solución Aplicada:** Localizado el nodo `TreeNode n6` una sola vez en `ActualizarCabecerasNivel6` y propagado como argumento hacia `CalcularColumnasNivel6(n6)` y `ObtenerXBaseNivel6(n6)`.

#### 5. Fase 1049 — Compensación de Borde de 1 px para `BorderStyle.FixedSingle`
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** El árbol dibuja en coordenadas cliente internas mientras que las cabeceras se ubican en el panel exterior, generando un desfase horizontal de 1 px por el borde `FixedSingle`.
- **Solución Aplicada:** Incorporada la constante `const int BordeArbol = 1;` restando 1 px en el posicionamiento horizontal `Left` de etiquetas y separadores.

#### 6. Fase 1050 — Restauración de Separador de Toolbar en `pnlToolbarIzq`
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.Designer.cs`
- **Diagnóstico:** La toolbar izquierda carecía de separación visual entre los botones de administración de perfiles y los botones de navegación del árbol tras el retiro de paneles runtime.
- **Solución Aplicada:** Declarado e instanciado `sepToolbar1` (1x20 px) en Designer e insertado entre `btnEliminarArbol` y `btnExpandirArbol`.

#### 7. Fase 1051 — Corrección de Citas de Evidencia y Sincronización Documental
- **Archivos Afectados:** `walkthrough.md`, `implementation_plan.md`
- **Diagnóstico:** Discrepancia en las líneas citadas para `ActualizarConteoCombinaciones` (3410-3439 vs 3326-3350) y necesidad de alinear la métrica de extensión del árbol al objetivo de 229 px.
- **Solución Aplicada:** Sincronizadas las evidencias de código y consolidados los registros de cambio en los artefactos de auditoría.

---

## 🚀 Ronda 206 — Sistema de Carriles Disjuntos y Oleadas Secuenciales (Fases 1052 a 1065)

### Propósito y Objetivos
Garantizar la viabilidad, calidad pedagógica y consistencia normativa de los horarios generados por el motor AG:
1. **Invariante I1:** Persistencia obligatoria de `Jornada`, `HorasDia` y `DiasVentana` desde la selección del árbol hacia `ag_aula` y `ag_aula_asignatura` mediante UPSERT.
2. **Invariantes I2 e I3:** Carriles disjuntos de 2 días ($\max(1, |\beta|/2)$) con rotación modular `carril = idxAsignatura mod carriles` por orden canónico `(orden_curricular, id)`.
3. **Invariante I4:** Oleadas temporales secuenciales encadenadas por carril (`oleada = idxAsignatura div carriles`), acotando sub-ventanas `[FechaInicio, FechaFin]` sin solape.
4. **Invariante I5:** Compuerta de pre-vuelo de capacidad verificando que la ventana contiene ranuras suficientes para la carga horaria.
5. **Invariante I6 / D4:** Blindaje de 4.0h en pregrado en todos los niveles (8.0h reservado estrictamente a posgrado con `HorasDia == 8`).
6. **Invariante I7:** Erradicación de la degradación silenciosa a 1 día; planes infactibles se marcan como vacantes `SIN_DISPONIBILIDAD`.
7. **Invariante I8:** Compuerta de seguridad en exportación y promoción ante `HardViolations > 0`.

### Fases Ejecutadas
- **Fase 1052:** Sincronización de dimensiones de decisión en `ag_aula` (`UPDATE`) y `ag_aula_asignatura` (`ON CONFLICT DO UPDATE`) en `AgFormGeneradorHorarios.cs`.
- **Fase 1053:** Mapeador `AgSeleccionMapper.cs` validando `Jornada`, `HorasDia` (4 u 8) y `DiasVentana` conservando el código canónico.
- **Fase 1054:** Modelo matemático y de dominio `AgCarriles.cs` con partición disjunta, rotación modular y aislamiento de días residuales.
- **Fase 1055:** POCO inmutable `AgPlanOleada.cs` modelando sub-ventanas, carril y oleada por asignatura.
- **Fase 1056:** Defensa en profundidad en `AgExpansorSesiones.cs` y regla dura `R11` en `AgValidadorReglas.cs` validando duración según nivel académico.
- **Fase 1057:** Servicio determinista `AgPlanificadorOleadas.cs` planificando sub-ventanas temporales y asignación de carriles.
- **Fase 1058:** Integración plan-first en `AgFabricaIndividuos.cs`, orden canónico `(OrdenCurricular, Id)`, supresión de degradación a 1 día y filtro estricto de bloques de 4h.
- **Fase 1059:** Adaptación de reglas `R5` (concurrencia semanal $\le$ carriles) y `R6` (secuencia y solape por carril) en `AgValidadorReglas.cs`.
- **Fase 1060:** Compuerta de seguridad ante `HardViolations > 0` en `ExportarResultadosExcel` y `ExportarProyeccionEspejo` de `AgFormGeneradorHorarios.cs`.
- **Fase 1061:** Migración V74 (`MigrationV74_SembrarBloques4hFaltantes.cs`) asegurando bloques de 4.0h para días 1 a 6 y registro en `SqliteHelper.cs`.
- **Fase 1062:** Mutación constructiva en `AgOperadorMutacion.cs` preservando duración de nivel y días del carril asignado.
- **Fase 1063:** Batería de pruebas unitarias `AgCarrilesTests.cs`, `AgPlanificadorOleadasTests.cs`, `AgSeleccionMapperTests.cs` y `MigrationV74Tests.cs` en `ConversorPDF.Tests`.
- **Fase 1064:** Artículo de documentación de cambios `docs/articles/changes/2026-10-22-ag-carriles-oleadas.md`.
- **Fase 1065:** Sincronización de artefactos de auditoría (`implementation_plan.md`, `walkthrough.md`) y análisis estático de calidad.

---

## 🔧 Ronda 207 — Aislamiento de Log en Pruebas y Descarte de Falso Positivo V66 (Fases 1066 a 1070)

### Propósito y Objetivos
1. **Descarte Concluyente del Falso Positivo V66:** Confirmar que los mensajes `[V66] Bloque vinculaciones falló.` acompañados de `[TEST-V66] fallo intencional` en `log_20261022.txt` corresponden a la ejecución de pruebas unitarias (`MigrationV66Tests.Up_WhenBlockFails_ShouldRollbackAndThrow`) y **no a un fallo en producción**. La tabla `vinculaciones` no tiene restricciones UNIQUE/CHECK sobre `modalidadVinculacion` en BD productiva y V66 corrió 4 veces OK el mismo día.
2. **Aislamiento de Registro en Pruebas:** Evitar que las suites de prueba contaminen `%LOCALAPPDATA%\ConversorPDF\log\` mediante redirección automática a `%TEMP%\ConversorPDF.Tests\log`.
3. **Cero Impacto en Producción:** `LoggerHelper` mantiene su ruta por defecto para la aplicación normal; el redireccionamiento se activa exclusivamente al cargar el ensamblado de pruebas mediante un inicializador de módulo.

### Fases Ejecutadas
- **Fase 1066 — Redirección configurable de la salida de log (`LoggerHelper.cs`):** Campo `_logDirectory` mutable y propiedad `internal static string DirectorioLog { get; set; }` para redirigir la carpeta de logs dinámicamente.
- **Fase 1067 — Shim de `ModuleInitializerAttribute` para net48 (`ModuleInitializerAttribute.cs`):** Definición de atributo bajo `System.Runtime.CompilerServices` con guard `#if !NET5_0_OR_GREATER` para habilitar inicializadores de módulo en C# 9+ sobre .NET Framework 4.8.
- **Fase 1068 — Bootstrap de pruebas con redirección a `%TEMP%` (`TestLogBootstrap.cs`):** Inicializador de módulo que desvía automáticamente `LoggerHelper.DirectorioLog` al cargar `ConversorPDF.Tests`.
- **Fase 1069 — Prueba de no regresión de la redirección (`LoggerHelperRedirectionTests.cs`):** Test unitario que valida la escritura en carpeta aislada y restauración limpia del estado.
- **Fase 1070 — Trazabilidad documental y clarificación en pruebas:** Comentario aclaratorio en `MigrationV66Tests.cs`, artículo técnico `docs/articles/changes/2026-10-22-aislamiento-log-pruebas.md` y sincronización de `implementation_plan.md` y `walkthrough.md`.

---

## 🔧 Ronda 208 — Carriles y Oleadas: Fugas de 8 h, Oleada Sobrescrita por Cruce y Carril Obsoleto (Fases 1071 a 1082)

### Propósito y Objetivos
1. **Erradicación Concluyente de 8 h en Pregrado:** Eliminar el `fallbackBloqueId` en fábrica que asignaba bloques arbitrarios (de 8 h o ajenos al día); colocar genes sin bloque adecuado en cuarentena (`IdBloqueHorario = 0`, `Incompleto`). Alinear franja TARDE en V74 a 13:00-17:00 coincidente con V62 y `AgFranjaJornada`.
2. **Preservación de Sub-Ventana de Oleada en Cruce:** Asegurar que el descendiente mantenga su plan temporal (`FechaInicioEstimada`, `FechaFinEstimada`), heredando del progenitor sólo cuando esté vacío, evitando solapes espurios en el mismo carril.
3. **Rotación Forzosa del Carril ante Cambio de Ventana:** Derivar siempre `dias_patron` desde la ventana vigente en la selección del árbol sin rehusar patrones obsoletos desalineados.
4. **Normalización Estricta de Carril en `DiasKDe`:** Validar que los patrones guarden la propiedad de carril disjunto de 2 días consecutivos; de lo contrario, normalizar al carril canónico de la ventana.
5. **Eliminación de Falsos Positivos Inter-Aula en Cohortes:** Discriminar por `IdAula` en R6 y R13 para permitir que múltiples aulas de la misma cohorte (presencial y virtual) coexistan sin colisión de oleadas.
6. **Aislamiento de Asignaturas Infactibles y Sanción Dura R7:** Propagar `R7_CoberturaHorariaIncompleta` dura si una asignatura carece de plan de oleada factible (I5), evitando que el AG acepte individuos vacantes como óptimos.
7. **Cálculo de Ranuras con Descuento de Festivos:** Descontar días feriados en `AgPlanificadorOleadas` para detectar infactibilidad I5 de forma temprana y exacta.
8. **Feedback Bloqueante en UI ante Configuración Incompleta:** Advertir al usuario con `MessageBox` modal si existen hojas del árbol sin jornada, ventana o duración configuradas al generar aulas.

### Fases Ejecutadas
- **Fase 1071 — Cuarentena estricta en fábrica y erradicación de fallbackBloqueId:** Reemplazo de fallback por `IdBloqueHorario = 0`, `EsVacante = 1`, `Estado = AgEstadoGen.Incompleto` en `AgFabricaIndividuos.cs:522-536`.
- **Fase 1072 — Protección de sub-ventana de oleada en operador de cruce:** Asignación condicional `if (!dst.FechaInicioEstimada.HasValue)` en `AgOperadorCruce.cs:143-146`.
- **Fase 1073 — Derivación modular y resiembra del carril sobre ventana vigente:** Generación invariable de `diasPatron` con `AgCarriles.DelCarril` en `AgFormGeneradorHorarios.cs:4195-4220`.
- **Fase 1074 — Validación de carril disjunto consecutivo en `DiasKDe`:** Validación `esCarril` (2 días, gap 1) en `AgPlanificadorPatronSemanal.cs:52-64`.
- **Fase 1075 — Discriminación por `IdAula` en R6 y R13:** Agrupación por `${x.IdAula}|${DiasPatron}` en `AgValidadorReglas.cs:374, 442`.
- **Fase 1076 — Validación de duración real del bloque en regla R11:** Resolución de duración desde `contexto.BloquesPorId[IdBloqueHorario]` en `AgValidadorReglas.cs:425-433`.
- **Fase 1077 — Aislamiento de asignaturas infactibles y sanción dura R7:** Extracción de `asig=` de error I5 en fábrica y adición de `R7_CoberturaHorariaIncompleta` en `AgValidadorReglas.cs:471-480`.
- **Fase 1078 — Descuento de festivos en cálculo de capacidad de oleadas:** Conteo de `diasEfectivos` en `AgPlanificadorOleadas.cs:120-130`.
- **Fase 1079 — Franja TARDE 4h canónica 13:00-17:00 en V74:** Ajuste de hora de inicio y fin en `MigrationV74_SembrarBloques4hFaltantes.cs:33-37`.
- **Fase 1080 — Eliminación de `<InternalsVisibleTo Include>` inerte en csproj legacy:** Limpieza de ítem inerte en `ConversorPDF.XLL.csproj:685-687`.
- **Fase 1081 — Detección bloqueante de combinaciones inválidas/incompletas en UI:** Parámetro `descartes` en `ObtenerCombinacionesSeleccionadas` y compuerta en `GenerarEncabezadosAulasAsync` de `AgFormGeneradorHorarios.cs:3390, 4022`.
- **Fase 1082 — Trazabilidad documental y suites de pruebas de no regresión:** Pruebas unitarias de F1071 a F1079 en `AgFabricaIndividuosTests.cs`, `AgOperadorCruceTests.cs`, `AgPlanificadorPatronSemanalTests.cs`, `AgValidadorReglasTests.cs`, `AgPlanificadorOleadasTests.cs`, `MigrationV74Tests.cs` y sincronización de artefactos.

---

## 🔧 Ronda 209 — F1071 y F1074 Incompletas: 8 h y Carril Inválido Persisten por Rutas Alternas (Fases 1083 a 1089)

### Propósito y Objetivos
1. **Erradicación Concluyente y Total de `fallbackBloqueId` (F1071 incompleta):** Sustituir la asignación residual de `fallbackBloqueId` en la rama de día único de `AgFabricaIndividuos.cs` por cuarentena estricta (`IdBloqueHorario = 0`, `EsVacante = 1`, `Estado = Incompleto`), y eliminar completamente la variable `fallbackBloqueId`.
2. **Validación Estricta de Carril en Sobrecarga Textual de `DiasKDe` (F1074 incompleta):** Asegurar que el overload textual `DiasKDe(string, string, AgAula, int)` en `AgPlanificadorPatronSemanal.cs` valide la consecutividad de 2 días disjuntos (`esCarril`) antes de devolver el patrón; de lo contrario, resolver el carril canónico de la ventana.
3. **Desacoplamiento de Motivo de Vacante en Regla Dura R7:** Introducir `AgMotivoVacante.SIN_PLAN_OLEADA = 6` para diferenciar la infactibilidad de oleada (I5) de la ausencia de bloques en catálogo (`SIN_DISPONIBILIDAD`), evitando que la falta de bloques de 4h bloquee la corrida completa bajo una etiqueta errónea.
4. **Filtro Estricto de Duración en Fallback del Cruce:** Restringir el último recurso de asignación de bloques en `AgOperadorCruce.cs` para que coincida exactamente con `unidadCruce` (4h en pregrado), impidiendo que se asignen bloques de 8h a asignaturas de pregrado.
5. **Calibración Robusta de Concurrencia R5 para Cohortes Multi-Aula:** Calcular `maxCarriles` como el máximo entre todas las aulas activas de la cohorte en `AgValidadorReglas.cs`, eliminando falsos positivos cuando conviven aulas presenciales (2 carriles) y virtuales (1 carril).
6. **Fuente Única de Verdad para Capacidad de Oleadas:** Delegar la validación de capacidad en `AgPlanificadorOleadas.cs` directamente a `AgPlanificadorPatronSemanal.Construir`, eliminando algoritmos duplicados de conteo de feriados.
7. **Sincronización Documental y Pruebas Unitarias:** Documentar la corrección integral en `implementation_plan.md` y `walkthrough.md`, y añadir pruebas de no regresión para F1083-F1088.

### Fases Ejecutadas
- **Fase 1083 — Cuarentena estricta en rama de día único y erradicación total de fallbackBloqueId:** Reemplazo de fallback por cuarentena (`IdBloqueHorario = 0`, `EsVacante = 1`, `Incompleto`, `HorasSemanales = horasEfectivas`) y eliminación de la variable en `AgFabricaIndividuos.cs:65, 414-427`.
- **Fase 1084 — Validación de carril disjunto en sobrecarga textual de `DiasKDe`:** Evaluación de `esCarril` (Count == 2 && gap == 1) en `AgPlanificadorPatronSemanal.cs:111-122`.
- **Fase 1085 — Desacoplamiento de motivo vacante `SIN_PLAN_OLEADA` y corrección de R7:** Inclusión de `AgMotivoVacante.SIN_PLAN_OLEADA = 6` en `AgMotivoVacante.cs:14`, asignación en `AgFabricaIndividuos.cs:288-298`, evaluación en `AgValidadorReglas.cs:482-490`, y atención en `AgFormGeneradorHorarios.cs:2285, 2327`.
- **Fase 1086 — Filtro de duración en asignación de bloque de último recurso en cruce:** Verificación estricta de `Math.Abs(b.DuracionHoras - unidadCruce) < 0.001` en `AgOperadorCruce.cs:193-211`.
- **Fase 1087 — Cálculo de `maxCarriles` usando el máximo de las aulas de la cohorte en R5:** Agrupación y cálculo con `.Max()` en `AgValidadorReglas.cs:325-331`.
- **Fase 1088 — Delegación de capacidad de oleada a `AgPlanificadorPatronSemanal.Construir`:** Reemplazo de cálculo manual por llamada unificada en `AgPlanificadorOleadas.cs:118-130`.
- **Fase 1089 — Trazabilidad documental y pruebas de no regresión:** Actualización de `walkthrough.md` y suite de pruebas unitarias (`AgFabricaIndividuosTests.cs`, `AgPlanificadorPatronSemanalTests.cs`, `AgValidadorReglasTests.cs`, `AgOperadorCruceTests.cs`, `AgPlanificadorOleadasTests.cs`, `AgMotivoVacanteTests.cs`).

---

## 🔧 Ronda 210 — Plan de Oleadas Recalculado por Individuo, Aborto Total del Plan y `SIN_PLAN_OLEADA` no Inmune a `Sellar` (Fases 1090 a 1096)

### Propósito y Objetivos
1. **Memoización del Plan de Oleadas (Fase 1090):** Evitar la regresión de rendimiento O(individuos × asignaturas) memoizando los planes deterministas por aula y parámetros de ventana en `_cachePlanes` de `AgFabricaIndividuos.cs` e inyectando `_planificadorPatron` en el constructor de `AgPlanificadorOleadas.cs`.
2. **Inmunidad de `SIN_PLAN_OLEADA` en `Sellar` (Fase 1091):** Evitar que la normalización de grupo en `AgCoherenciaGrupo.cs` elimine el motivo `SIN_PLAN_OLEADA` cuando existe al menos un gen no vacante (e.g. cruce mixto), garantizando que la regla dura R7 y el diagnóstico en UI no se pierdan.
3. **Planificación Parcial ante Asignaturas Infactibles (Fase 1092):** Reemplazar el aborto total del plan por la acumulación de infactibles en `AgPlanOleadasResultado`, permitiendo que las asignaturas factibles conserven sus carriles y oleadas y aislando únicamente las que no caben en su oleada.
4. **Bloqueo Estricto de Aulas sin Ventana (Fase 1093):** Bloquear aulas con `dias_ventana` vacíos con `[I1_VENTANA_VACIA]` en `AgValidadorVentana.ValidarFactibilidad`, advertir en `AgPlanificadorPatronSemanal.DiasKDe` y marcar la cohorte como infactible en la fábrica en vez de planificar con `{2,3}` silencioso.
5. **Cálculo Canónico de Carriles en Advertencias (Fase 1094):** Sustituir `Math.Ceiling(count / 2.0)` por `AgCarriles.Cantidad(ventanaIsoCfg)` en `AgValidadorVentana.cs`.
6. **Capacidad Teórica Real en `AgPlanOleada` (Fase 1095):** Corregir `CapacidadRanuras` para reflejar la holgura teórica real (`semanasOleada × diasCarril`) en lugar de reportar tautológicamente las sesiones requeridas.
7. **Desacoplamiento de Motivo `SIN_BLOQUE_NIVEL` (Fase 1096):** Introducir `AgMotivoVacante.SIN_BLOQUE_NIVEL = 7` en `AgMotivoVacante.cs`, asignarlo en `AgFabricaIndividuos.cs` ante ausencia de bloque horario de 4h, y mapear su tratamiento contextual en `AgFormGeneradorHorarios.cs`.

### Fases Ejecutadas
- **Fase 1090 — Memoización de planes de carril/oleada e inyección de planificador:** Diccionario `_cachePlanes` en `AgFabricaIndividuos.cs:20, 220` e inyección de `_planificadorPatron` en `AgPlanificadorOleadas.cs:17-21`.
- **Fase 1091 — Inmunidad de `SIN_PLAN_OLEADA` en `AgCoherenciaGrupo.Sellar`:** Guarda prioritaria en `AgCoherenciaGrupo.cs:234-245` que preserva vacante y motivo `SIN_PLAN_OLEADA`.
- **Fase 1092 — Planificación parcial con `AgPlanOleadasResultado`:** Modelo `AgPlanOleadasResultado.cs`, retorno de planes factibles + lista de infactibles en `AgPlanificadorOleadas.cs:115-152` y consumo selectivo en `AgFabricaIndividuos.cs:235-245`.
- **Fase 1093 — Detección bloqueante de aula sin ventana y traza en fallback:** Bloqueo `[I1_VENTANA_VACIA]` en `AgValidadorVentana.cs:209-214`, traza en `AgPlanificadorPatronSemanal.cs:90-94` y cuarentena en `AgFabricaIndividuos.cs:253-260`.
- **Fase 1094 — Corrección de cálculo de carriles con `AgCarriles.Cantidad` en validador de ventana:** Unificación en `AgValidadorVentana.cs:178`.
- **Fase 1095 — Capacidad teórica real en modelo de oleada:** Asignación de `semanasPorOla[oleada] * diasPatron.Count` en `AgPlanificadorOleadas.cs:136`.
- **Fase 1096 — Desacoplamiento de motivo vacante `SIN_BLOQUE_NIVEL`:** Creación de `AgMotivoVacante.SIN_BLOQUE_NIVEL = 7`, asignación en `AgFabricaIndividuos.cs:437, 546` y atención en `AgFormGeneradorHorarios.cs:2287, 2325`.

---

## 🔧 Ronda 211 — Errores de Compilación por Tipo No Registrado y `using` Faltante; Cuarentena I1 Efectiva (Fases 1097 a 1102)

### Propósito y Objetivos
1. **Restauración de Compilación en Proyecto Legacy (Fase 1097):** Registrar explícitamente `<Compile Include="Academic\Domain\Models\Genetica\AgPlanOleadasResultado.cs" />` en `ConversorPDF.XLL.csproj` para resolver error bloqueante CS0246.
2. **Importación de Namespace Funcional (Fase 1098):** Incorporar `using CSharpFunctionalExtensions;` en `AgFabricaIndividuos.cs` para resolver el tipo `Result<T>` en `_cachePlanes`.
3. **Efectividad Operativa de Cuarentena I1 (Fase 1099):** Asegurar que cuando una cohorte carece de plan de oleadas (`cohorteSinPlan = planInfactible && asigsInfactibles.Count == 0`), sus genes se fijen con `IdBloqueHorario = 0`, `DiasPatron = string.Empty` y `Estado = Incompleto`, impidiendo que `AgMaterializadorSesiones` materialice sesiones espurias con el centinela.
4. **Desduplicación de Sanción Dura R14 (Fase 1100):** Excluir del chequeo genérico `R14_BloqueHorarioNoAsignado` a los genes con motivo tipificado `SIN_BLOQUE_NIVEL`.
5. **Observabilidad y Verificación Real de Memoización (Fase 1101):** Incorporar contador interno `Construcciones` en `AgPlanificadorOleadas.cs` y validar en prueba unitaria que se ejecute exactamente 1 vez para individuos de la misma corrida.
6. **Aislamiento de Caché por Corrida y Período (Fase 1102):** Limpiar `_cachePlanes` al detectar cambio en `idCorrida` e incluir `idCorrida` e `IdPeriodo` en la clave compuesta de la memoización.

### Fases Ejecutadas
- **Fase 1097 — Registro de `AgPlanOleadasResultado.cs` en csproj:** Inclusión en `ConversorPDF.XLL.csproj:653`.
- **Fase 1098 — Directiva `using CSharpFunctionalExtensions;` en fábrica:** Importación de namespace en `AgFabricaIndividuos.cs:8`.
- **Fase 1099 — Cuarentena efectiva I1 en construcción de genes:** Anulación de bloque (`IdBloqueHorario = 0`, `DiasPatron = ""`, `Incompleto`) ante `cohorteSinPlan` en `AgFabricaIndividuos.cs:312, 455, 595`.
- **Fase 1100 — Exclusión de `SIN_BLOQUE_NIVEL` en regla R14:** Guarda en `AgValidadorReglas.cs:172-178` para no duplicar sanción.
- **Fase 1101 — Contador de construcciones en planificador de oleadas y prueba estricta:** Propiedad `Construcciones` en `AgPlanificadorOleadas.cs:19, 34` y aserción en `AgFabricaIndividuosTests.cs:330`.
- **Fase 1102 — Invalidación de caché por corrida y clave extendida:** Limpieza de `_cachePlanes` ante nueva corrida y clave con periodo en `AgFabricaIndividuos.cs:23, 53, 230`.

---

## 🔧 Ronda 212 — `SIN_BLOQUE_NIVEL` sin Sanción Dura, Preservación en Sellar y Corrección de Factibilidad I5 (Fases 1103 a 1107)

### Propósito y Objetivos
1. **Sanción Dura para Genes `SIN_BLOQUE_NIVEL` (Fase 1103):** Garantizar que un gen con carga horaria y motivo `SIN_BLOQUE_NIVEL` emita la regla dura `R7_CoberturaHorariaIncompleta`, evitando que el individuo eluda la penalización y la asignatura desaparezca silenciosamente en materialización.
2. **Preservación de Motivos No Asignables en `Sellar` (Fase 1104):** Extender la guarda de sellado en `AgCoherenciaGrupo.cs` para proteger tanto `SIN_PLAN_OLEADA` como `SIN_BLOQUE_NIVEL` ante cruces mixtos, asegurando que ningún gen sin bloque sea revertido erróneamente a docente real.
3. **Cálculo Real de Demanda en Pre-Vuelo I5 (Fase 1105):** Eliminar la asignación ficticia de 64h ante `HorasObjetivo == 0` en `AgValidadorVentana.cs`, resolviendo la intensidad desde el catálogo de asignaturas y excluyendo materias desconocidas (<= 0) con advertencia `[I5_HORAS_DESCONOCIDAS]` para no bloquear corridas viables.
4. **Desduplicación de `SIN_PLAN_OLEADA` en R14 (Fase 1106):** Excluir `SIN_PLAN_OLEADA` de `R14_BloqueHorarioNoAsignado` en `AgValidadorReglas.cs`, evitando duplicidad con la sanción dura `R7` previamente tipificada.
5. **Cobertura Integral de Pruebas Unitarias (Fase 1107):** Añadir pruebas en `AgValidadorReglasTests.cs`, `AgCoherenciaGrupoTests.cs` y `AgValidadorVentanaTests.cs` para certificar la emisión de R7 dura, la exclusión de R14, la preservación en `Sellar` y la no-infactibilidad ante `HorasObjetivo == 0`.

### Fases Ejecutadas
- **Fase 1103 — Regla dura R7 para motivo `SIN_BLOQUE_NIVEL`:** Bucle en `AgValidadorReglas.cs:494-504` que emite `AgReglaViolacion.R7_CoberturaHorariaIncompleta` dura.
- **Fase 1104 — Preservación de `SIN_BLOQUE_NIVEL` y `SIN_PLAN_OLEADA` en `Sellar`:** Generalización con `motivoNoAsignable` en `AgCoherenciaGrupo.cs:234-245`.
- **Fase 1105 — Demanda real en validador de ventana I5 sin default de 64h:** Resolución en `AgValidadorVentana.cs:203, 222-235` y emisión de advertencia `[I5_HORAS_DESCONOCIDAS]`.
- **Fase 1106 — Exclusión de `SIN_PLAN_OLEADA` en verificación R14:** Guarda combinada en `AgValidadorReglas.cs:172-178`.
- **Fase 1107 — Cobertura de pruebas unitarias para `SIN_BLOQUE_NIVEL`, `Sellar` e I5:** Nuevas pruebas en `AgValidadorReglasTests.cs:610-670`, `AgCoherenciaGrupoTests.cs:252-290` y `AgValidadorVentanaTests.cs:211-232`.

---

## 🔧 Ronda 213 — Fechas de Oleadas en Vinculaciones/Promoción, Min/Max de Grupo, Conteo I5 y Contador Atómico (Fases 1108 a 1112)

### Propósito y Objetivos
1. **Preservación de Fechas de Oleada en Vinculaciones (Fase 1108):** Evitar que `MapGrupo` y `AgMaterializadorSesiones` sustituyan las fechas estimadas de la sub-ventana por las fechas completas del semestre cuando `IdPeriodo > 0`, manteniendo la visibilidad de oleadas en `ag_vinculaciones` y `ag_promocion`.
2. **Cálculo de Rango Min/Max de Fechas por Grupo (Fase 1109):** En grupos multi-día, calcular `FechaInicio` y `FechaFin` a partir del mínimo y máximo de las fechas estimadas de los genes que componen el grupo, en lugar de adoptar únicamente los valores del primer gen (`g0`).
3. **Conteo Coherente de Materias Computadas en Bloqueo I5 (Fase 1110):** Reportar explícitamente en el mensaje de error `[I5_INFACTIBLE]` el número de materias efectivamente computadas en la demanda frente al total asignado (`({materiasComputadas} materias computadas de {asigsAula.Count})`).
4. **Sincronización Atómica del Contador de Construcciones (Fase 1111):** Sincronizar el contador `Construcciones` en `AgPlanificadorOleadas.cs` mediante `Interlocked.Increment` y `Volatile.Read` para garantizar estabilidad bajo generación paralela o multihilo.
5. **Cobertura de Pruebas Unitarias para Fechas de Oleada e I5 (Fase 1112):** Añadir pruebas en `AgGenToAgVinculacionMapperTests`, `AgMaterializadorSesionesTests` y `AgValidadorVentanaTests` para certificar la conservación de fechas de oleada, el cálculo de Min/Max y el conteo de materias computadas.

### Fases Ejecutadas
- **Fase 1108 — Preservación de sub-ventana de oleada en mapeador y materializador:** Fechas de genes preservadas en `AgGenToAgVinculacionMapper.cs:73-82` y acotadas sin sobreescritura en `AgMaterializadorSesiones.cs:99-108`.
- **Fase 1109 — Rango Min/Max en `MapGrupo`:** `fiGrupo` y `ffGrupo` calculados mediante `Min()` y `Max()` sobre `genes` en `AgGenToAgVinculacionMapper.cs:74-75`.
- **Fase 1110 — Reporte de materias computadas en advertencia/bloqueo I5:** Contador `materiasComputadas` y mensaje actualizado en `AgValidadorVentana.cs:227, 238, 244`.
- **Fase 1111 — Contador `Construcciones` sincronizado en planificador:** `_construcciones` con `Volatile.Read` e `Interlocked.Increment` en `AgPlanificadorOleadas.cs:18-20, 39`.
- **Fase 1112 — Pruebas unitarias de oleadas en vinculación e I5:** Pruebas en `AgGenToAgVinculacionMapperTests.cs:60-129`, `AgMaterializadorSesionesTests.cs:320-384` y `AgValidadorVentanaTests.cs:231-254`.

---

## 🔧 Ronda 214 — Pre-vuelo UI con Duración Efectiva por Nivel, Invariante Dura de Nivel (F1114) y Parseo Invariante en Sub-Ventana (Fases 1113 a 1116)

### Propósito y Objetivos
1. **Validación Pre-Vuelo con Duración Efectiva por Nivel (Fase 1113):** Resolver `horasEfectivasV` vía `AgUnidadesSemanales.HorasSesionAula` según el nivel del programa en `AgFormGeneradorHorarios.cs`, impidiendo que una selección de 8h en Pregrado bloquee falsamente la corrida cuando el catálogo no posee bloques de 8h.
2. **Invariante Dura de Nivel — Ningún Gen de PREGRADO con Bloque de 8h (Fase 1114 corregida):** Establecer una invariante global en el punto de convergencia de operadores (`AgCoherenciaGrupo.Sellar`), fijar la duración impuesta por el nivel en `AgOperadorMutacion.cs`, bloquear en pre-vuelo (`[D4_PREGRADO_8H]` en `AgValidadorVentana.cs` y aborto con mensaje claro en `AgFormGeneradorHorarios.cs`), y certificar con pruebas unitarias en todas las capas.
3. **Parseo Culturalmente Invariante en Cotas de Fechas (Fase 1115):** Utilizar `DateTime.TryParseExact` con `CultureInfo.InvariantCulture` en `AgMaterializadorSesiones.cs` para evitar omisiones de cotas o fallos de parseo bajo culturas no gregorianas, advirtiendo ante cadenas inválidas.
4. **Desacoplamiento de Patrón en Pre-Vuelo de Aula (Fase 1116):** Pasar `string.Empty` como `diasPatron` en la validación pre-vuelo de aula en `AgFormGeneradorHorarios.cs`, dado que los patrones específicos por asignatura se validan a nivel de motor.

### Fases Ejecutadas
- **Fase 1113 — Pre-vuelo UI con horas efectivas de nivel:** Resolución de `horasEfectivasV` en `BtnGenerarAulas_Click` y `BtnOptimizar_Click` de `AgFormGeneradorHorarios.cs:1113-1126, 1360-1373`.
- **Fase 1114 — Invariante dura de nivel (1114.a a 1114.d):**
  - *1114.a (Sellar):* Normalización de `HorasSemanales` a `uNivel`, reasignación de bloque del mismo día o cuarentena con `SIN_BLOQUE_NIVEL` en `AgCoherenciaGrupo.cs:56-96`.
  - *1114.b (Mutación):* Duración objetivo impuesta por nivel del aula en `AgOperadorMutacion.cs:110-125`.
  - *1114.c (Bloqueo Pre-vuelo):* Bloqueo `[D4_PREGRADO_8H]` en `AgValidadorVentana.cs:222-230` y aborto bloqueante en `AgFormGeneradorHorarios.cs:1117-1125, 1373-1381`.
  - *1114.d (Pruebas):* Cobertura en `AgCoherenciaGrupoTests.cs:288-356`, `AgOperadorMutacionTests.cs:192-228`, `AgValidadorVentanaTests.cs:254-276` y `AgFabricaIndividuosTests.cs:405-432`.
- **Fase 1115 — Parseo `InvariantCulture` en cota de sub-ventana:** `DateTime.TryParseExact` con advertencia de log en `AgMaterializadorSesiones.cs:100-123` y prueba en `AgMaterializadorSesionesTests.cs:387-463`.
- **Fase 1116 — Desacoplamiento de patrón en pre-vuelo:** Argumento `string.Empty` como `diasPatron` en `AgFormGeneradorHorarios.cs:1128, 1384`.

---

## 🔧 Ronda 215 — Robustecimiento de Invariante de Nivel: Protección de POSGRADO, Resguardo de Genes sin Duración e Indexación O(1) de Bloques (Fases 1117 a 1121)

### Propósito y Objetivos
1. **Resguardo de Genes con Duración Desconocida y Protección de Fechas (Fase 1117):** Excluir genes con `HorasSemanales <= 0` de la invariante en `AgCoherenciaGrupo.Sellar` y limitar la corrección estrictamente a casos de sobre-duración (`g.HorasSemanales > uNivel + 0.001`), evitando borrar `FechasGen` ni mandar genes válidos a vacante `SIN_BLOQUE_NIVEL`.
2. **Corrección de Unidad por Nivel sin Degradar POSGRADO (Fase 1118):** En `AgCoherenciaGrupo.Sellar`, hacer que `UnidadNivel` retorne `-1.0` si el aula o programa no son resolubles, en lugar del fallback a `4.0` que degradaba erróneamente asignaturas legítimas de POSGRADO (8h) o en aulas no indexadas.
3. **Cobertura de No-Regresión de la Invariante (Fase 1119):** Añadir pruebas unitarias específicas en `AgCoherenciaGrupoTests.cs` verificando que POSGRADO conserva 8h, que genes con `HorasSemanales = 0` no pierden fechas ni estado, y que aulas no indexadas no son alteradas.
4. **Causa de Fallo Específica en Pre-Vuelo (Fase 1120):** Incorporar `HorasDiaInvalidoNivel = 13` en `AgCausaFallo.cs` y registrarlo en las trazas de fallo de pre-vuelo P07 y P08 en `AgFormGeneradorHorarios.cs`.
5. **Optimización O(1) de Bloques en Sellar (Fase 1121):** Construir un diccionario indexado de bloques activos por `(DiaSemanaNumero, DuracionHoras)` al inicio de `Sellar`, sustituyendo la búsqueda lineal O(Genes × Bloques) repetida en cada llamada de cruce, mutación y materialización.

### Fases Ejecutadas
- **Fase 1117 — Protección contra disparo de invariante con HorasSemanales <= 0:** Cláusulas de guarda en `AgCoherenciaGrupo.cs:72-74` evitando borrar `FechasGen` o alterar genes sin sobre-duración.
- **Fase 1118 — Retorno -1.0 en UnidadNivel ante aulas no resolubles:** Prevención de asunción indebida de PREGRADO en `AgCoherenciaGrupo.cs:68`.
- **Fase 1119 — Pruebas unitarias de no-regresión de invariante:** Implementación de `Sellar_PosgradoConBloque8h_NoDegradaDuracion`, `Sellar_GenSinHorasSemanales_NoModificaFechasNiEstado` y `Sellar_AulaNoIndexada_NoAlteraDuracion` en `AgCoherenciaGrupoTests.cs:360-464`.
- **Fase 1120 — Enum AgCausaFallo.HorasDiaInvalidoNivel:** Nuevo código en `AgCausaFallo.cs:50` y referenciado en `AgFormGeneradorHorarios.cs:1122, 1385`.
- **Fase 1121 — Indexación O(1) de bloques en Sellar:** Diccionario `indiceBloques` agrupado y búsqueda directa en `AgCoherenciaGrupo.cs:53-57, 80-81`.

---

## 🔧 Ronda 216 — Validación Pre-Vuelo por Sub-Ventana de Oleada, Capacidad Canónica de Carril y Caché de Bloques en Contexto (Fases 1122 a 1126)

### Propósito y Objetivos
1. **Validación de Sub-Ventana de Oleada por Asignatura (Fase 1122):** Invocar `AgPlanificadorOleadas.Construir` dentro de `ValidarFactibilidad` en `AgValidadorVentana.cs`, detectando asignaturas cuya carga horaria exceda la sub-ventana asignada a su oleada (`[I5_INFACTIBLE_OLEADA]`), evitando que demandas individuales inviables pasen el pre-vuelo.
2. **Capacidad Canónica de Carril — 1 Sesión/Día (Fase 1123):** Fijar `sesionesPorDia = 1` en `CapacidadDe` de `AgValidadorVentana.cs` en correspondencia con el modelo de carriles disjuntos (R4), impidiendo la sobreestimación artificial de capacidad a 2 sesiones/día ante selecciones de 8h/día.
3. **Cálculo Canónico de Carriles en Pre-Vuelo (Fase 1124):** Utilizar `AgCarriles.Cantidad(ventanaIso)` en lugar de `Math.Ceiling(count / 2.0)` en `CapacidadDe` de `AgValidadorVentana.cs`, alineando el reporte diagnóstico de multicarril con la partición real del motor.
4. **Caché Estable de Bloques en AgContexto (Fase 1125):** Implementar `TryGetBloque(int dia, double dur, out AgBloqueHorario b)` con diccionario memoizado en `AgContexto.cs` e invalidación en `ConstruirIndices()`, eliminando la re-creación repetitiva del diccionario de bloques en `AgCoherenciaGrupo.Sellar`.
5. **Cobertura de No-Regresión de Oleada y Capacidad (Fase 1126):** Diseñar pruebas unitarias en `AgValidadorVentanaTests.cs` que certifiquen el bloqueo ante asignaturas que exceden su oleada y verifiquen que la capacidad no se duplica.

### Fases Ejecutadas
- **Fase 1122 — Validación de sub-ventana por oleada en pre-vuelo:** Bloqueo `[I5_INFACTIBLE_OLEADA]` en `AgValidadorVentana.cs:255-276`.
- **Fase 1123 — Capacidad canónica de 1 sesión/día/carril:** `sesionesPorDia = 1` en `AgValidadorVentana.cs:122`.
- **Fase 1124 — Carriles canónicos con AgCarriles.Cantidad:** Corrección de `Math.Ceiling` en `AgValidadorVentana.cs:129`.
- **Fase 1125 — Caché indexado de bloques en AgContexto:** Campo `_indiceBloquesDiaDur`, método `TryGetBloque` en `AgContexto.cs:150-169`, e integración en `AgCoherenciaGrupo.cs:81-89`.
- **Fase 1126 — Cobertura de pruebas unitarias de oleada y carril:** Pruebas `ValidarFactibilidad_AsignaturaExcedeSuOleada_BloqueaI5` y `ValidarFactibilidad_Pregrado8h_NoDuplicaCapacidadPorDia` en `AgValidadorVentanaTests.cs:277-343`.

---

## 🔧 Ronda 217 — Coherencia de Jornada en Reasignación de Duración, Invalidación Explícita de Caché de Bloques y Alcance Curricular en Pre-Vuelo (Fases 1127 a 1130)

### Propósito y Objetivos
1. **Coherencia Jornada ↔ Franja en Corrección de Duración (Fase 1127):** Evitar colapsos de franja en `AgContexto.TryGetBloque` alindexar listas de bloques por `(DiaSemanaNumero, DuracionHoras)` y filtrar por la franja canónica de la jornada del aula (`AgFranjaJornada.Resolve`), garantizando que aulas de jornada TARDE reciban bloques de franja TARDE tras normalizar 8h→4h.
2. **Invalidación Explícita de Caché ante Mutación de Bloques (Fase 1128):** Exponer `InvalidarIndiceBloques()` en `AgContexto.cs` e invocarlo en `ConstruirIndices()` y en los puntos de recarga de bloques en `AgFormGeneradorHorarios.cs` (P05), previniendo falsos `SIN_BLOQUE_NIVEL` por catálogos obsoletos.
3. **Alineación de Alcance Curricular en Pre-Vuelo de Oleadas (Fase 1129):** Filtrar `itemsPlan` en `AgValidadorVentana.cs` según coincidencia de programa y semestre del aula (`AgFormatoAsignatura.SemestreRomano`), impidiendo bloqueos falsos `[I5_INFACTIBLE_OLEADA]` por materias fuera del alcance que el motor nunca planificará.
4. **Cobertura de No-Regresión de Franja, Caché y Alcance Curricular (Fase 1130):** Añadir pruebas unitarias verificando la asignación de franja según jornada, la actualización del catálogo tras invalidación y la exclusión de asignaturas de otro semestre en pre-vuelo.

### Fases Ejecutadas
- **Fase 1127 — Coherencia de jornada y franja preferida en TryGetBloque:** Índice a `List<AgBloqueHorario>` en `AgContexto.cs:153-195` y resolución con `AgFranjaJornada.Resolve` en `AgCoherenciaGrupo.cs:79-88`.
- **Fase 1128 — InvalidarIndiceBloques ante recarga de bloques:** Método `InvalidarIndiceBloques()` en `AgContexto.cs:157` y llamada en `AgFormGeneradorHorarios.cs:762`.
- **Fase 1129 — Filtro de itemsPlan por programa y semestre del aula:** Cláusula `Where` curricular en `AgValidadorVentana.cs:261-265`.
- **Fase 1130 — Cobertura de pruebas unitarias de jornada, caché y alcance:** Pruebas `Sellar_PregradoTarde8h_ReasignaBloqueDeFranjaTarde` y `Contexto_TrasInvalidarIndiceBloques_UsaCatalogoVigente` en `AgCoherenciaGrupoTests.cs:463-524`, y `ValidarFactibilidad_AulaConAsignaturaDeOtroSemestre_NoBloquea` en `AgValidadorVentanaTests.cs:342-372`.

---

## 🔧 Ronda 218 — Coherencia de Jornada en Cruce, Mutación y Regla R11 con Unificación de Selección de Bloques (Fases 1131 a 1135)

### Propósito y Objetivos
1. **Coherencia de Jornada en Operador de Cruce (Fase 1131):** Evitar la reasignación arbitraria al primer bloque del día en `AgOperadorCruce.cs` al seleccionar bloques compatibles con la franja canónica de la jornada del aula (`AgFranjaJornada.Resolve`), garantizando que aulas de jornada TARDE reciban bloques de franja TARDE.
2. **Coherencia de Jornada en Operador de Mutación (Fase 1132):** Filtrar bloques candidatos en mutación mono-día y multi-día de `AgOperadorMutacion.cs` según la jornada del aula, impidiendo que genes de aulas NOCHE o TARDE migren a franjas distintas.
3. **Validación Dura R11 para Discrepancias Franja ∉ Jornada (Fase 1133):** Extender la regla dura R11 en `AgValidadorReglas.cs` para validar que la franja del bloque asignado a cada gen corresponda estrictamente a la jornada del aula (`AgFranjaJornada.Resolve`), bloqueando individuos inconsistentes.
4. **Cobertura de Pruebas Unitarias de Jornada en Operadores y Validador (Fase 1134):** Agregar pruebas unitarias en `AgOperadorCruceTests.cs`, `AgOperadorMutacionTests.cs` y `AgValidadorReglasTests.cs` asegurando que el cruce, la mutación y la regla R11 respeten y certifiquen la jornada del aula.
5. **Punto Único de Selección de Bloque Horario (Fase 1135):** Crear `AgSeleccionBloque.cs` como punto único para resolver día, duración y franja por jornada, eliminando la duplicación en fábrica, cruce, mutación y coherencia.

### Fases Ejecutadas
- **Fase 1131 — Selección de bloque por jornada en cruce:** Uso de `AgSeleccionBloque.ParaDia` en `AgOperadorCruce.cs:182-205`.
- **Fase 1132 — Filtro de franja por jornada en mutación:** Resolución de `franjaObjetivo` y filtro de candidatos en `AgOperadorMutacion.cs:126-175`.
- **Fase 1133 — Regla R11 con validación estricta de franja vs jornada:** Emisión de violación dura R11 en `AgValidadorReglas.cs:445-459`.
- **Fase 1134 — Cobertura de pruebas unitarias de jornada:** Pruebas `Cruzar_AulaTarde_AsignaBloqueDeFranjaTarde` (`AgOperadorCruceTests.cs:213-241`), `Mutacion_AulaTarde_NoMigraAManana` (`AgOperadorMutacionTests.cs:227-260`) y `Validador_GenDeFranjaFueraDeJornada_EmiteR11` (`AgValidadorReglasTests.cs:672-708`).
- **Fase 1135 — Unificación de selección en AgSeleccionBloque:** Clase `AgSeleccionBloque.cs` creada, registrada en `ConversorPDF.XLL.csproj` e integrada en `AgOperadorCruce.cs`, `AgCoherenciaGrupo.cs`.

---

## 🔧 Ronda 219 — Formato Canónico DIA|horas|horario en Columna K del Exportador Excel (Fases 1136 a 1138)

### Propósito y Objetivos
1. **Formato Canónico en Columna K de Excel (Fase 1136):** Sustituir la emisión simple de nombres de días `"MARTES/MIERCOLES"` en `SemestresExcelExporter.cs` por la cadena canónica `DIA|horas|horario` (ej. `"MARTES|4|08:00-12:00, MIERCOLES|4|08:00-12:00"`), permitiendo verificar la unidad horaria de 4h y los días del carril.
2. **Overload en DiasHorasSemanaFormatter y Limpieza de TODO (Fases 1136 & 1137):** Crear el método `DesdeSesionListItems(IEnumerable<SesionListItem>)` en `DiasHorasSemanaFormatter.cs` para el modelo desnormalizado del exportador masivo, retirando comentarios pendientes de cableado.
3. **Cobertura de Pruebas Unitarias para Columna K (Fase 1138):** Crear `DiasHorasSemanaFormatterTests.cs` en `ConversorPDF.Tests` con cobertura para pares de carril, deduplicación de sesiones semanales, casos sin día y colecciones nulas.
4. **Evaluación de Botón Cargar en Módulo R (Fase 1139):** Confirmar preservación de compuertas de seguridad existentes en `AgFormGeneradorHorarios.cs:2086-2097`, posponiendo la refactorización de flujo por tratarse de alcance de producto y no de defecto.

### Fases Ejecutadas
- **Fase 1136 — Cableado de DiasHorasSemanaFormatter en SemestresExcelExporter:** Overload `DesdeSesionListItems` en `DiasHorasSemanaFormatter.cs:57-93` y asignación en `SemestresExcelExporter.cs:377-380`.
- **Fase 1137 — Retiro de TODO y preservación de granularidad en grilla:** Eliminado `TODO` en `DiasHorasSemanaFormatter.cs` y confirmada separación de columnas en `FormSesionesAcademicas.cs`.
- **Fase 1138 — Pruebas unitarias de DiasHorasSemanaFormatter:** 5 pruebas en `DiasHorasSemanaFormatterTests.cs`.

---

## 🔧 Ronda 220 — Implementación F1-F4 compilación-bloqueada y saneamiento V76 inefectivo en filas legacy (Fases 1195 a 1201)

### Propósito y Objetivos
1. **Corrección de Compilación AgMensajeDimensionesAula (Fase 1195):** Corregir import de namespace para `AgAula` (`ConversorPDF.XLL.Academic.Domain.Models.Genetica`) en `AgMensajeDimensionesAula.cs` y suite de pruebas `AgMensajeDimensionesAulaTests.cs`.
2. **Saneamiento Determinista V76 para Filas Legacy y Múltiples Árboles (Fases 1196 & 1197):** Flexibilizar `ag_aula` con subunidad vacía en el JOIN contra `ag_arbol_nodo`, y resolver ambigüedad multivariable seleccionando deterministamente el nodo con `MAX(idarbol)` que posea dimensiones válidas.
3. **Aislamiento por Subunidad y Marcado de Suciedad en Sincronización (Fase 1198):** Incorporar parámetro `subunidad` en `AgAulaDimensionesSync.SincronizarAsync`, filtrar `WHERE (@Sub = '' OR COALESCE(subunidad,'') = @Sub OR COALESCE(subunidad,'') = '')` y notificar `SyncContext.Scheduler?.MarkDirty()`.
4. **Tope en Modal, Normalización de Saltos de Línea y Telemetría Estructurada (Fase 1199):** Limitar modal a 5 aulas, normalizar saltos `\n` por `" ¶ "` en logs y enriquecer telemetría de traza en pasos P07 y P08.
5. **Reintento de Migración V76 ante Fallos (Fase 1200):** Eliminar captura silenciosa en `MigrationV76_SanearDimensionesAgAula.cs` para permitir que `SchemaMigrator` registre el hueco de versión y reintente en el siguiente arranque.
6. **Evaluación de Duplicación y Obsolecencia de Métodos (Fase 1201):** Ratificar marca `[Obsolete]` preexistente en `GenerarEncabezadoAulaAsync` y preservar simetría de etapas P07 y P08 sin introducir abstracciones artificiales.

### Fases Ejecutadas
- **Fase 1195 — Corrección de namespace en AgMensajeDimensionesAula:** Namespace `ConversorPDF.XLL.Academic.Domain.Models.Genetica` en `AgMensajeDimensionesAula.cs` y `AgMensajeDimensionesAulaTests.cs`.
- **Fase 1196 — Tolerancia a subunidad vacía en V76:** Cláusula `(COALESCE(x.subunidad,'') = COALESCE(a.subunidad,'') OR TRIM(COALESCE(a.subunidad,'')) = '')` en `MigrationV76_SanearDimensionesAgAula.cs`.
- **Fase 1197 — Determinismo por MAX(idarbol) en V76:** Subconsulta correlacionada `MAX(x.idarbol)` eliminando `GROUP BY a.id` en `MigrationV76_SanearDimensionesAgAula.cs`.
- **Fase 1198 — Sincronización con subunidad y scheduler dirty:** Parámetro `subunidad`, cláusula `@Sub` y `Scheduler?.MarkDirty()` en `AgAulaDimensionesSync.cs`, actualización de llamada en `AgFormGeneradorHorarios.cs:105` y pruebas en `AgAulaDimensionesSyncTests.cs`.
- **Fase 1199 — Paginación de modal y saneamiento de log en P07/P08:** Límite de 5 aulas en modal, reemplazo de `\n` por `" ¶ "` en log y telemetría de `AulasSinDimensiones` y `AulasConErrorVentana` en `AgFormGeneradorHorarios.cs:1189-1205, 1483-1499`.
- **Fase 1200 — Relanzamiento de excepción en V76:** `LoggerHelper.Error` y `throw` en catch de `MigrationV76_SanearDimensionesAgAula.cs:115-120`.
- **Fase 1201 — Evaluación crítica de duplicación P07/P08 y API:** Retirada como refactor innecesario; preservadas identidades de traza independientes P07/P08 y ratificado `[Obsolete]` en `GenerarEncabezadoAulaAsync`.
- **Fase 1202 — Registro de archivos nuevos en ConversorPDF.XLL.csproj:** Inclusión de `AgMensajeDimensionesAula.cs`, `AgAulaDimensionesSync.cs` y `MigrationV76_SanearDimensionesAgAula.cs` en `ConversorPDF.XLL.csproj:660-662`, desbloqueando la compilación en Debug/Release x64 y la ejecución de la suite de pruebas.

---

## 🔧 Ronda 223 — Planificador de oleadas v3 por bandas de días (Fases 1224 a 1235)

### Propósito y Objetivos
1. **Partición de Ventana en Bandas Disjuntas de Días (Fases 1224 & 1225):** Reemplazar reparto de carriles de 2 ranuras por partición contigua de días en `AgBandasDias` y validar determinismo y propiedades matemáticas en `AgBandasDiasTests`.
2. **Planificador de Oleadas v3 Atómico por Semanas (Fases 1226 & 1227):** Eliminar `olas = ceil(n/carriles)` y sincronización por rebanadas globales en `AgPlanificadorOleadas`, sustituyéndolas por pipelines secuenciales por banda y búsqueda de asignación exhaustiva/BFD; extender `AgPlanOleadasResultado` con campos de partición y holgura.
3. **Escalera de Capacidad y Escalado Asíncrono de Aulas (Fases 1228 & 1229):** Definir `AgEscaleraVentanas` para sugerencias monótonas dentro del catálogo (`VIE-SAB` → `JUE-SAB`) e implementar `AgEscaladorVentanaAula` para dry-run de déficit y aplicación transaccional en `ag_aula` y `ag_arbol_nodo`.
4. **Validación de Factibilidad Accionable y Detección de Holgura Baja (Fase 1230):** Enriquecer `[I5_INFACTIBLE]` en `AgValidadorVentana` con ventana sugerida de catálogo, emitir advertencia `[I5_HOLGURA_BAJA]` al superar 88% de ocupación y canalizar patrones declarados.
5. **Huella de Patrones en Clave de Caché y Write-Back por Banda (Fases 1231 & 1232):** Incorporar huella de patrones en la clave de memoización de `AgFabricaIndividuos` y sembrar `dias_patron` por banda en `AgFormGeneradorHorarios.cs:4544` para prevenir contradicciones con R8.
6. **Diálogo Interactivo Pre-Vuelo y Suites de Prueba (Fases 1233 a 1235):** Integrar diálogo de propuesta de escalado en diagnóstico pre-vuelo de UI, actualizar suites canónicas T1-T5 en `AgPlanificadorOleadasTests` y documentar especificación formal y limitaciones.

### Estado: PLANIFICADO — NO EJECUTADO (retirado en Ronda 223)

Artefactos ausentes verificados con `Test-Path`: `AgBandasDias.cs`, `AgEscaleraVentanas.cs`,
`AgEscaladorVentanaAula.cs`, `AgBandasDiasTests.cs`, `docs/articles/changes/2026-10-23-ag-oleadas-banda-dias.md`.
Sin registros `<Compile Include=...>` en `ConversorPDF.XLL.csproj`.

### Fases Pendientes (1224 a 1235)
- **Fase 1224 — AgBandasDias:** pendiente.
- **Fase 1225 — Suite de AgBandasDias:** pendiente.
- **Fase 1226 — AgPlanificadorOleadas v3:** pendiente.
- **Fase 1227 — Resultado extendido:** pendiente.
- **Fase 1228 — AgEscaleraVentanas + tests:** pendiente.
- **Fase 1229 — AgEscaladorVentanaAula + tests:** pendiente.
- **Fase 1230 — Validador accionable y holgura baja:** pendiente.
- **Fase 1231 — Fábrica con huella de patrón:** pendiente.
- **Fase 1232 — Write-back de dias_patron por banda:** pendiente.
- **Fase 1233 — Cableado pre-vuelo:** pendiente.
- **Fase 1234 — Suites canónicas de prueba:** pendiente.
- **Fase 1235 — Documentación y verificación:** pendiente.

---

## 🔧 Ronda 223 — Corrección de trazabilidad y 5 defectos reales de motor (Fases 1236 a 1241)

### Propósito y Objetivos
1. **Saneamiento de Trazabilidad Fantasma (Fase 1236):** Desmarcar las 12 fases 1224-1235 como ejecutadas en `implementation_plan.md` y `walkthrough.md`.
2. **Desbloqueo de Cura de Dimensiones con horas_dia NULL (Fase 1237):** Proteger `horas_dia` con `COALESCE(horas_dia, -1) <> @H` en `AgAulaDimensionesSync.cs`.
3. **Aislamiento Estricto por Subunidad en Sync (Fase 1238):** Evitar que subunidad vacía barra otras subunidades mediante `COALESCE(NULLIF(TRIM(subunidad), ''), @Sub) = @Sub`.
4. **Desacople de Factibilidad de la Cota de Oleadas (Fase 1239):** Reemplazar `W < olas` por cálculo de capacidad total de ranuras y acotar `oleada` a los límites del array en `AgPlanificadorOleadas.cs`.
5. **Siembra de Patrón de Días por Subconjunto Estricto (Fase 1240):** Asignar patrón de ≤2 días por alternancia modular en `AgFormGeneradorHorarios.cs:4544` para no forzar ventana completa en `ag_aula_asignatura`.
6. **Huella de Patrones Declarados en Clave de Memoización (Fase 1241):** Agregar `huellaPatrones` a `clavePlan` en `AgFabricaIndividuos.cs` para invalidar caché al reasignar patrones.

### Fases Ejecutadas
- **Fase 1236 — Corrección de trazabilidad:** `implementation_plan.md` y `walkthrough.md` rebajados a PLANIFICADO.
- **Fase 1237 — horas_dia NULL en sync:** `COALESCE(horas_dia, -1) <> @H` en `AgAulaDimensionesSync.cs:51`.
- **Fase 1238 — Aislamiento por subunidad:** `COALESCE(NULLIF(TRIM(subunidad), ''), @Sub) = @Sub` en `AgAulaDimensionesSync.cs:48`.
- **Fase 1239 — Factibilidad por capacidad de ranuras:** Validación de `demandaTotal > capacidadTotal` y clamp de `oleada` en `AgPlanificadorOleadas.cs:110-128, 215`.
- **Fase 1240 — Siembra por banda estricta:** Sembrado modular de días por subconjunto (<=2 días) en `AgFormGeneradorHorarios.cs:4544`.
- **Fase 1241 — Huella de patrones en caché:** `huellaPatrones` integrada en `clavePlan` en `AgFabricaIndividuos.cs:225-240`.

---

## 🔧 Ronda 224 — Corrección de regresiones en oleadas, siembra y pruebas (Fases 1242 a 1246)

### Propósito y Objetivos
1. **Saneamiento de Test Obsoleto (Fase 1242):** Actualizar aserción de `I5_SIN_SEMANAS_POR_OLEADA` a `I5_SIN_CAPACIDAD` en `AgPlanificadorOleadasTests.cs`.
2. **Prevención de Solape Ola × Carril (Fase 1243):** Detectar acumulación de demanda sobrepasando oferta en `(Ola, Carril)` en `AgPlanificadorOleadas.cs` y marcar infactibilidad accionable.
3. **Suspensión de Siembra de Bandas hasta v3 (Fase 1244):** Preservar siembra por carriles detrás del flag `USAR_BANDAS_DIAS = false` en `AgFormGeneradorHorarios.cs:4540-4560`.
4. **Mensajes Precisos de Semanas Hábiles y Capacidad (Fase 1245):** Aislar `I5_SIN_SEMANAS_HABILES` cuando `W == 0` y corregir texto de ranuras/semana en `I5_SIN_CAPACIDAD`.
5. **Protección Anti-Clobber en Sync y Alineación de Evidencia (Fase 1246):** Incorporar `COALESCE(NULLIF(@J, ''), jornada)` y `CASE WHEN @H > 0` en `AgAulaDimensionesSync.cs:45-47` y alinear referencias en `walkthrough.md`.

### Fases Ejecutadas
- **Fase 1242 — Test suite alineada a I5_SIN_CAPACIDAD:** Corrección de aserción en `AgPlanificadorOleadasTests.cs:84`.
- **Fase 1243 — Guard anti-solape ola × carril:** `cargaOlaCarril` y saturación en `AgPlanificadorOleadas.cs:228-255`.
- **Fase 1244 — Diferimiento de siembra por banda:** Flag `USAR_BANDAS_DIAS = false` en `AgFormGeneradorHorarios.cs:4544-4556`.
- **Fase 1245 — Mensajes de capacidad precisos:** `I5_SIN_SEMANAS_HABILES` e `I5_SIN_CAPACIDAD` en `AgPlanificadorOleadas.cs:125-135`.

---

## 🔧 Ronda 225 — Planificador de oleadas v3 por bandas de días y correcciones (Fases 1224 a 1235 + 1247/1248)

### Propósito y Objetivos
1. **Predicado Efectivo en Sync (Fase 1248):** Comparar contra valores efectivos en `AgAulaDimensionesSync.cs` (`COALESCE(jornada, '') <> COALESCE(NULLIF(@J, ''), COALESCE(jornada, ''))`, etc.).
2. **Partición de Ventana en Bandas Disjuntas de Días (Fases 1224 & 1225):** Crear `AgBandasDias.cs` con resolución determinista (exhaustiva/BFD) y validar en `AgBandasDiasTests.cs`.
3. **Contrato de Resultado Extendido (Fase 1227):** Exponer `Particion`, `VentanaSugerida`, `PatronesRenormalizados` y `OcupacionPorcentaje` en `AgPlanOleadasResultado.cs`.
4. **Escalera de Catálogo y Persistencia Asíncrona (Fases 1228 & 1229):** Crear `AgEscaleraVentanas.cs` y `AgEscaladorVentanaAula.cs` para sugerencias idempotentes y aplicación transaccional en `ag_aula` y `ag_arbol_nodo`.
5. **Motor v3 por Bandas de Días y Saturación por Exceso (Fases 1226 & 1247):** Implementar partición por bandas en `AgPlanificadorOleadas.cs`, asignando `Carril = banda` y `Oleada = 0`, marcando solo el excedente en saturación.
6. **Validación Accionable y Holgura Baja (Fase 1230):** Enriquecer `[I5_INFACTIBLE]` y `[I5_INFACTIBLE_OLEADA]` con ventana sugerida de `AgEscaleraVentanas`, y emitir `[I5_HOLGURA_BAJA]` si ocupación ≥ 88%.
7. **Huella de Patrones (Fase 1231):** Confirmada cubierta por Fase 1241 (`huellaPatrones` en `AgFabricaIndividuos.cs`).
8. **Write-Back por Banda (Fase 1232):** Activar flag `USAR_BANDAS_DIAS = true` en `AgFormGeneradorHorarios.cs:4551`.
9. **Diálogo Interactivo Pre-Vuelo (Fase 1233):** Insertar consulta de confirmación modal de escalado en pre-vuelo de `AgFormGeneradorHorarios.cs`.
10. **Suites Canónicas T1–T5 (Fase 1234):** Actualizar tests unitarios en `AgPlanificadorOleadasTests.cs` para reflejar bandas paralelas y validar casos 60/64 sesiones.
11. **Documentación Formal de Cambios (Fase 1235):** Crear `docs/articles/changes/2026-10-23-ag-oleadas-banda-dias.md`.

### Fases Ejecutadas
- **Fase 1248 — Predicado efectivo en sync:** Implementado en `AgAulaDimensionesSync.cs:51-53`.
- **Fase 1224 — AgBandasDias:** Implementado en `AgBandasDias.cs`, registrado en `ConversorPDF.XLL.csproj:653`.
- **Fase 1225 — AgBandasDiasTests:** Implementado en `AgBandasDiasTests.cs`.
- **Fase 1227 — Resultado extendido:** `AgPlanOleadasResultado.cs` actualizado con propiedades complementarias.
- **Fase 1228 — AgEscaleraVentanas:** Implementado en `AgEscaleraVentanas.cs`, registrado en `ConversorPDF.XLL.csproj:654`.
- **Fase 1229 — AgEscaladorVentanaAula:** Implementado en `AgEscaladorVentanaAula.cs`, registrado en `ConversorPDF.XLL.csproj:655`.
- **Fase 1226 — Motor v3 por bandas de días:** Reemplazado `Construir` en `AgPlanificadorOleadas.cs` con asignación determinista por banda.
- **Fase 1247 — Saturación no-cascada por banda:** Integrado en bucle de evaluación de bandas de `AgPlanificadorOleadas.cs`.
- **Fase 1230 — Validador con sugerencia y holgura baja:** `AgValidadorVentana.cs:295-344`.
- **Fase 1231 — Clave memoización patrones:** Cubierta por Fase 1241 en `AgFabricaIndividuos.cs:234-242`.
- **Fase 1232 — Write-back por banda:** `USAR_BANDAS_DIAS = true` en `AgFormGeneradorHorarios.cs:4551`.
- **Fase 1233 — Diálogo pre-vuelo de escalado:** Insertado en pre-vuelo de `AgFormGeneradorHorarios.cs:1235`.
- **Fase 1234 — Suites canónicas T1-T5:** Actualizado `AgPlanificadorOleadasTests.cs` (pruebas de bandas y casos T1-T5).
- **Fase 1235 — Documentación v3:** Creado `docs/articles/changes/2026-10-23-ag-oleadas-banda-dias.md`.

---

## 🔧 Ronda 226 — Auditoría de v3: bandas, escalera, escalador (Fases 1251 a 1262)

### Propósito y Objetivos
1. **Patrón de Días como Restricción (Fase 1251):** `AgBandasDias.cs` filtra candidatos mediante `ParticionAlojaPatrones` en dos pasadas (estricta y relajada).
2. **Cableado de diasPatronPorAsignatura (Fase 1252):** Conectar diccionario de patrones en `AgFabricaIndividuos.cs:245` y `AgValidadorVentana.cs:324`.
3. **Fuente Única de Saturación (Fase 1253):** Llenar `saturadas` directamente desde `asign.ExcesoPorBanda` en `AgPlanificadorOleadas.cs`.
4. **Limpieza de Variables Muertas (Fase 1254):** Eliminar `carriles` y `olas` en `AgPlanificadorOleadas.cs:86-87` (CS0219).
5. **Alineación de Suites de Prueba (Fase 1255):** Actualizar aserciones v2 a semántica v3 en `AgBandasDiasTests.cs` y `AgPlanificadorOleadasTests.cs`.
6. **Infactibilidad por Violación de Patrón R8 (Fase 1256):** Reportar infactibilidad accionable ante patrones que no caben en la banda asignada.
7. **Salvaguarda de Ventana Vacía (Fase 1257):** Retorno seguro en fallback de `AgBandasDias.Resolver` cuando `ventana` está vacía (evita `IndexOutOfRangeException`).
8. **Filtro de Árbol en Nodo (Fase 1258):** Condición `(@IdArbol <= 0 OR idarbol = @IdArbol)` en `AgEscaladorVentanaAula.cs` para no contaminar otros árboles.
9. **Filtro de Demanda por Programa y Semestre (Fase 1259):** Alinear `DemandaAula` en `AgEscaladorVentanaAula.cs` con el alcance del validador.
10. **Suites Unitarias de Escalera y Escalador (Fase 1260):** Crear `AgEscaleraVentanasTests.cs` y `AgEscaladorVentanaAulaTests.cs`.
11. **Refactorización Pre-Vuelo y Revalidación (Fase 1261):** Identificador ASCII `decisionesEscalado`, instancia única de escalador y revalidación en `AgFormGeneradorHorarios.cs`.
12. **Eliminación de Escalón Lateral (Fase 1262):** Remover `MAR-VIE → LUN-JUE` en catálogo de `AgEscaleraVentanas.cs`.

### Fases Ejecutadas
- **Fase 1251 — Patrón como restricción:** Dos pasadas con `ParticionAlojaPatrones` en `AgBandasDias.cs`.
- **Fase 1252 — Cableado de diasPatron:** Conectado en `AgFabricaIndividuos.cs` y `AgValidadorVentana.cs`.
- **Fase 1253 — Saturación por ExcesoPorBanda:** `saturadas` derivado del solver en `AgPlanificadorOleadas.cs`.
- **Fase 1254 — Eliminación de CS0219:** Eliminadas variables `carriles` y `olas` en `AgPlanificadorOleadas.cs`.
- **Fase 1255 — Tests v3 actualizados:** `AgBandasDiasTests.cs` y `AgPlanificadorOleadasTests.cs`.
- **Fase 1256 — Infactibilidad R8 accionable:** Detección de patrón no alojado en banda en `AgPlanificadorOleadas.cs`.
- **Fase 1257 — Fallback seguro de ventana vacía:** Retorno vacío antes de `demandaBanda` en `AgBandasDias.cs`.
- **Fase 1258 — Filtro idarbol en sqlNodo:** Parámetro `@IdArbol` en `AgEscaladorVentanaAula.cs`.
- **Fase 1259 — Filtro de demanda por aula:** Asignaturas filtradas por programa/semestre en `AgEscaladorVentanaAula.cs`.
- **Fase 1260 — Tests unitarios escalera y escalador:** Creados `AgEscaleraVentanasTests.cs` y `AgEscaladorVentanaAulaTests.cs`.
- **Fase 1261 — Pre-vuelo saneado con revalidación:** Identificadores limpios y `ValidarFactibilidad` tras escalado en `AgFormGeneradorHorarios.cs`.
- **Fase 1262 — Catálogo escalera sin lateral:** `MAR-VIE` escala directamente a `1,2,3,4,5` en `AgEscaleraVentanas.cs`.

---

## 🔧 Ronda 227 — Auditoría de compilación y regresiones de v3 (Fases 1263 a 1270)

### Propósito y Objetivos
1. **Restitución de Llave de Cierre en Ranuras (Fase 1263):** Reparar CS1513 en `AgPlanificadorOleadas.cs:84` restituyendo llave de cierre del bloque de ranuras de compatibilidad.
2. **Corrección de Acceso a InformeFactibilidad (Fase 1264):** Reemplazar llamadas a propiedades inexistentes por `EsFactible` y `BloqueosCriticos` en `AgFormGeneradorHorarios.cs:1263-1267` (CS1061).
3. **Exigencia de Factibilidad en Pasada Estricta (Fase 1265):** Descartar planes con `excesoTotal > 0` en pasada estricta de `AgBandasDias.Evaluar` para no bloquear relajación.
4. **Coherencia de Test de Patrón Común (Fase 1266):** Alinear aserción en `AgPlanificadorOleadasTests.cs` reconociendo que el patrón exige contención en banda.
5. **Capacidad Efectiva de Banda por Ocurrencia y Festivos (Fase 1267):** Pasar `capacidadDia` desde `AgPlanificadorOleadas.cs` a `AgBandasDias.Resolver` y `Evaluar`.
6. **Claridad de Árbol en Pre-Vuelo (Fase 1268):** Documentar propagación de `dias_ventana` desde pre-vuelo a `ag_aula` y nodos de árbol asociados.
7. **Reporte Específico R8 en Validador (Fase 1269):** Emitir `[R8_DiaFueraDePatron]` diferenciado de `[I5_INFACTIBLE_OLEADA]` en `AgValidadorVentana.cs:348-351`.
8. **Limpieza de Métodos Muertos v2 (Fase 1270):** Remover `ContarDias` y `CabeEnSemanas` de `AgPlanificadorOleadas.cs`.

### Fases Ejecutadas
- **Fase 1263 — Llave restituida en ranuras:** Llave cerrada en `AgPlanificadorOleadas.cs:84`.
- **Fase 1264 — Contrato InformeFactibilidad:** `EsFactible` y `BloqueosCriticos` en `AgFormGeneradorHorarios.cs:1263-1267`.
- **Fase 1265 — Factibilidad estricta en bandas:** `if (exigirPatrones && excesoTotal > 0) return;` en `AgBandasDias.cs:238`.
- **Fase 1266 — Test patrón común v3:** `Construir_CuandoPatronComun_AmbasAsignaturasCompartenBanda` en `AgPlanificadorOleadasTests.cs:334-345`.
- **Fase 1267 — CapacidadDia por festivos:** `capacidadDia` en `AgBandasDias.cs` y `AgPlanificadorOleadas.cs`.
- **Fase 1268 — Documentación pre-vuelo árbol:** Comentario aclaratorio en `AgFormGeneradorHorarios.cs:1252`.
- **Fase 1269 — Reporte R8 en validador:** `[R8_DiaFueraDePatron]` en `AgValidadorVentana.cs:348-351`.
- **Fase 1270 — Limpieza v2:** Eliminados métodos residuales en `AgPlanificadorOleadas.cs`.

---

## 🔧 Ronda 228 — Corrección de infactibles falsos y saneamiento de v3 (Fases 1271 a 1275)

### Propósito y Objetivos
1. **Reintento en Ventana Completa (Fase 1271):** En `AgPlanificadorOleadas.cs`, si la sub-ventana truncada (`iniBanda..finBanda`) no alberga la demanda requerida debido a festivos, reintentar sobre la ventana completa (`fechaFin.Date`) antes de declarar infactible.
2. **Aislamiento Estricto de Árbol en Pre-Vuelo (Fase 1272):** Condicionar la ejecución de `sqlNodo` a `idArbol > 0` en `AgEscaladorVentanaAula.cs` para evitar actualizar nodos de otros árboles cuando `idArbol == 0`.
3. **Capacidad Efectiva y Ocupación con Festivos (Fase 1273):** Calcular `capacidadDia` antes de la guarda global `demandaTotal > capacidadTotal` en `AgPlanificadorOleadas.cs`, y usar `capacidadDia.Values.Sum()` para `slotsTotales` en `AgBandasDias.cs:275`.
4. **Revalidación Asíncrona fuera de UI (Fase 1274):** Ejecutar `ValidarFactibilidad` dentro de `Task.Run` con `ConfigureAwait(true)` en `AgFormGeneradorHorarios.cs:1264`.
5. **Retiro de Restos Muertos de v2 (Fase 1275):** Eliminar la colección y bucle inalcanzable de `saturadas`, la rama `saturadas.Contains` en `detalleInfactibles`, y el método privado no referenciado `ContarRanuras` en `AgPlanificadorOleadas.cs`.

### Fases Ejecutadas
- **Fase 1271 — Reintento en ventana completa:** Implementado en `AgPlanificadorOleadas.cs:179-184`.
- **Fase 1272 — Pre-vuelo no toca otros árboles:** `if (idArbol > 0)` en `AgEscaladorVentanaAula.cs:122` y doc en `AgFormGeneradorHorarios.cs:1252`.
- **Fase 1273 — Capacidad efectiva pre-guard y ocupación:** `capacidadDia` en `AgPlanificadorOleadas.cs:128-142` y `AgBandasDias.cs:275-278`.
- **Fase 1274 — Revalidación asíncrona:** `await Task.Run` en `AgFormGeneradorHorarios.cs:1264`.
- **Fase 1275 — Limpieza saturadas y ContarRanuras:** Eliminados de `AgPlanificadorOleadas.cs`.

---

## 🔧 Ronda 229 — Sincronización de perfil activo y coherencia de reporte de oleadas (Fases 1276 a 1278)

### Propósito y Objetivos
1. **Sincronización de Perfil Activo en Pre-Vuelo (Fase 1276):** Pasar `idArbolActivo` desde `cmbArboles.SelectedItem` en `AgFormGeneradorHorarios.cs:1253` para actualizar `ag_arbol_nodo` del perfil activo sin afectar árboles ajenos.
2. **Coherencia de Semanas y Capacidad tras Reintento (Fase 1277):** Si la banda se amplió a la ventana completa en `AgPlanificadorOleadas.cs`, reportar `semanasReportadas = W` y `capacidadReportada = RanurasEfectivas(...)` en `AgPlanOleada`, restituyendo la función `RanurasEfectivas`.
3. **Detalle de Infactibilidad con Ranuras Efectivas (Fase 1278):** Reportar ranuras efectivas (festivos descontados) mediante `RanurasEfectivas` en el mensaje de error de `AgPlanificadorOleadas.cs:209-211`.

### Fases Ejecutadas
- **Fase 1276 — Sincronización idArbolActivo:** `idArbolActivo` en `AgFormGeneradorHorarios.cs:1252-1254` y comentario en `AgEscaladorVentanaAula.cs:122`.
- **Fase 1277 — Coherencia de SemanasOleada/CapacidadRanuras:** `semanasReportadas`, `capacidadReportada` y `RanurasEfectivas` en `AgPlanificadorOleadas.cs:184, 216-228, 240-249`.
- **Fase 1278 — Ranuras efectivas en detalle de error:** Mensaje contextual en `AgPlanificadorOleadas.cs:210-212`.

---

## 🔧 Ronda 230: Recuperación de Última Optimización AG y Persistencia Resiliente de Pernocta (Fases 1279 a 1283)

### Propósito y Objetivos
1. **Soporte de Columna isPernoctaUltimoDia en Staging (Fase 1279):** Crear e implementar `MigrationV77_AddPernoctaAgVinculaciones` para asegurar la columna `isPernoctaUltimoDia INTEGER NOT NULL DEFAULT 0` en `vinculaciones` y `ag_vinculaciones`, evitando excepciones SQL en el repositorio académico en modo AG. Registrar en `SqliteHelper.cs` y `ConversorPDF.XLL.csproj`.
2. **Sustitución de Botones de Exportación en UI (Fase 1280):** En `AgFormGeneradorHorarios.Designer.cs`, remover `btnExportarProyeccion` y `btnExportarSimulacion`, e incorporar `btnCargarUltimaOptimizacion` inmediatamente a la izquierda de `btnPromoverProduccion`.
3. **Carga y Rehidratación de Última Optimización (Fase 1281):** En `AgFormGeneradorHorarios.cs`, cablear el botón a `CargarUltimaOptimizacionAsync()`, consultando `_scheduleRepository.GetUltimaCorridaAsync` y `VinculacionQueryRepository.GetSesionesAsync`, rehidratando grids, vacantes, labels de métricas y eliminando 4 métodos muertos de exportación.
4. **Endurecimiento Transaccional en FormSesionesAcademicas (Fase 1282):** Validar `SqliteResult.IsFailure` en inserción/actualización de vinculación y sesiones para prevenir reasignación de IdVinculacion inválido, y dinamizar tabla de consulta en eliminación con `FROM {_modo.TablaVinculaciones}`.
5. **Suites de Pruebas Unitarias (Fase 1283):** Validar en `MigrationV77Tests.cs` los metadatos, idempotencia y adición de columna, y en `AcademicRepositoryAgModoPernoctaTests.cs` la resolución del error SQL y la persistencia de pernocta en `ag_vinculaciones`.

### Fases Ejecutadas
- **Fase 1279 — Migración V77:** Implementada en `MigrationV77_AddPernoctaAgVinculaciones.cs`, registrada en `SqliteHelper.cs:378` y `ConversorPDF.XLL.csproj:668`.
- **Fase 1280 — Control UI btnCargarUltimaOptimizacion:** Configurado en `AgFormGeneradorHorarios.Designer.cs:90, 930, 973-982, 1284`.
- **Fase 1281 — CargarUltimaOptimizacionAsync:** Implementado en `AgFormGeneradorHorarios.cs:354, 2061-2212` con remoción de exportadores obsoletos.
- **Fase 1282 — Guardado y borrado resiliente:** En `FormSesionesAcademicas.cs:1330-1348, 2590`.
- **Fase 1283 — Suites de pruebas unitarias:** Creados `MigrationV77Tests.cs` y `AcademicRepositoryAgModoPernoctaTests.cs`.

---

## 🔧 Ronda 231: Consolidación de Carga de Staging y Saneamiento de Estado (Fases 1284 a 1288)

### Propósito y Objetivos
1. **Desduplicación de Carga de Resultados (Fase 1284):** Hacer que `CargarUltimaOptimizacionAsync` delegue en `CargarModuloResultadosAsync(corrida.Id)` eliminando ~110 líneas duplicadas y garantizando la población de violaciones y vacantes auditables.
2. **Saneamiento de Estado de Sesión ante Corrida Vacía (Fase 1285):** Restablecer a cero/nulo `_idCorridaCargada` y `_ultimoResultado`, y restaurar etiquetas de modo y colores de advertencia cuando la corrida no posee sesiones en staging.
3. **Desacople Diagnóstico de Fallos SQL (Fase 1286):** En `CargarModuloResultadosAsync`, separar la verificación de `sesionesResult.IsFailure` con log de error específico de la condición `sesionesResult.Value.Count == 0`.
4. **Evaluación de Pernocta en Vistas AG (Fase 1287):** Exponer `isPernoctaUltimoDia` en `vw_AgSesionesDetalle` y `vw_AgVinculacionesDetalle` mediante nueva migración V78, condicionada a decisión de alcance previa por su impacto en 5 suites de pruebas.
5. **Ampliación de Dimensiones del Botón de Carga (Fase 1288):** Ajustar ancho a 195px (`Location = (535, 10)`, `Size = (195, 32)`) para prevenir truncamiento tipográfico bajo escalado DPI > 100%.

### Fases Ejecutadas
- **Fase 1284 — Desduplicación de Carga:** Delegación en `CargarModuloResultadosAsync` en `AgFormGeneradorHorarios.cs:2088-2105`.
- **Fase 1285 — Saneamiento de Estado Vacío:** Limpieza de `_idCorridaCargada`, `_ultimoResultado` y labels en `AgFormGeneradorHorarios.cs:2093-2102`.
- **Fase 1286 — Desacople de IsFailure:** Logging de fallo SQL aislado de corrida vacía en `AgFormGeneradorHorarios.cs:1917-1926`.
- **Fase 1287 — Pernocta en Vistas AG:** Retirada / diferida a ronda dedicada por requerir nueva migración V78 y actualización de 5 suites de fixtures.
- **Fase 1288 — Ampliación de Botón UI:** Dimensiones `(535, 10)` / `(195, 32)` en `AgFormGeneradorHorarios.Designer.cs:976-978`.

---

## Ronda 232 — Contrato Explícito de Carga, Saneamiento de Pantalla Mixta y Documentación XML (Fases 1289 a 1291)

### Propósito y Objetivos
1. **Contrato Explícito de Carga (Fase 1289):** Sustituir la detección de fallo indirecta basada en `DataSource == null` por un contrato directo `Task<bool>` en `CargarModuloResultadosAsync`. Consumir el valor retornado en `CargarUltimaOptimizacionAsync` para evitar falsos éxitos cuando la corrida no posee sesiones o falla la rehidratación.
2. **Saneamiento de Pantalla Mixta ante Fallo de Staging (Fase 1290):** Al detectar `sesionesResult.IsFailure` en `CargarModuloResultadosAsync`, limpiar inmediatamente el grid de horarios, lista/grid de vacantes, etiquetas de métricas (`lblMejorFitness`, `lblHardViolaciones`, `lblVacantes`) y deshabilitar `btnPromoverProduccion` para prevenir mezclas de datos de corridas anteriores.
3. **Actualización de Documentación XML (Fase 1291):** Actualizar el comentario XML de `CargarUltimaOptimizacionAsync` para reflejar con precisión la delegación a `CargarModuloResultadosAsync` y el origen en staging (`ag_sesiones` + `ag_vinculaciones`) en vez de la vista de producción.
4. **Fase 1287:** Permanece diferida a nueva migración V78 y actualización transversal de 5 suites de pruebas.

### Fases Ejecutadas
- **Fase 1289 — Contrato Explícito Task<bool>:** `CargarModuloResultadosAsync` retorna `Task<bool>` y `CargarUltimaOptimizacionAsync` evalúa `!cargada` en `AgFormGeneradorHorarios.cs:1847, 1884, 2064, 2110-2123`.
- **Fase 1290 — Saneamiento ante Fallo de Staging:** Limpieza completa de UI y vacantes en rama `sesionesResult.IsFailure` en `AgFormGeneradorHorarios.cs:1917-1930`.
- **Fase 1291 — Actualización de Documentación XML:** Corrección de documentación en `AgFormGeneradorHorarios.cs:2066-2070`.

---

## 🔧 Ronda 233: Corrección Quirúrgica de Franja Horaria de Sábado y Defensa de Invariante I7 (Fases 1292 a 1296)

### Propósito y Objetivos
1. **Tests Primero (Fase 1292):** Validar en pruebas unitarias que: (a) En sábado 4h se asignen bloques de MANANA y TARDE; (b) En sábado 8h posgrado se conserve franja única; (c) L-V conserve franja única; (d) El genotipo resultante no emita violaciones duras en R4 ni R14.
2. **Corrección de Ranuras en Fábrica (Fase 1293):** En `AgFabricaIndividuos.cs`, condicionar `permitidasRanura.Count > 1` (`franjaExactaObligatoria = true`) para exigir la franja exacta de la ranura, evitando el colapso a MANANA en sábado 4h.
3. **Sellar Franjas de Sábado (Fase 1294):** En `AgCoherenciaGrupo.cs:Sellar`, reasignar el 2.º gen del sábado a la franja alterna si comparte franja con el 1.º, previniendo regresiones por cruce/mutación.
4. **Deduplicación en Columna K (Fase 1295):** En `AgMaterializadorSesiones.cs:DiaHorasSemanalesDe`, agrupar por `(Día, HoraInicio, HoraFin)` en vez de `IdBloqueHorario` para emitir `SABADO|4|08:00 A 12:00, SABADO|4|13:00 A 17:00`.
5. **Invariante I7 de Sesiones Duplicadas (Fase 1296):** En `AgExpansorSesiones.cs` y `AgMaterializadorSesiones.cs`, rechazar y poner en cuarentena grupos o individuos con sesiones duplicadas en `(FechaSesion, HorarioSesion)`.

### Fases Ejecutadas
- **Fase 1292 — Cobertura de Pruebas Unitarias:** Implementados tests en `AgFabricaIndividuosTests.cs` y `AgValidadorReglasTests.cs`.
- **Fase 1293 — Exigencia de Franja Exacta en Ranuras:** En `AgFabricaIndividuos.cs:429-438`, incorporada `franjaExactaObligatoria` cuando `permitidasRanura.Count > 1`.
- **Fase 1294 — Sellar Franja Alterna en Sábado:** En `AgCoherenciaGrupo.cs:108-154`, reasignación determinista a franja alterna en sábado.
- **Fase 1295 — Deduplicación de Columna K por Horario:** En `AgMaterializadorSesiones.cs:136-175`, agrupación por `(DiaSemana, HoraInicio, HoraFin)`.
- **Fase 1296 — Invariante I7:** En `AgExpansorSesiones.cs:187-206` y `AgMaterializadorSesiones.cs:258-272`.

---

## 🔧 Ronda 234: R8 Permite Desdoble de Sábado y Saneamiento de Fallbacks y Tests (Fases 1297 a 1302)

### Propósito y Objetivos
1. **Corrección de Identificador Enum en Tests (Fase 1297):** Reemplazar `R14_SolapeHorarioBloque` por el miembro canónico `R14_BloqueSolapadoPorIntervalo` en `AgValidadorReglasTests.cs:765`.
2. **R8 Evalúa Unicidad por Par (Día, Franja) (Fase 1298):** En `AgValidadorReglas.cs:580-592`, evaluar pares `(DiaSemana, Franja)` en vez de días puros para que sábado 4h con MANANA y TARDE no reciba violación dura espuria.
3. **Reproducción Efectiva en Test de Fábrica (Fase 1299):** En `AgFabricaIndividuosTests.cs:436-468`, añadir segunda asignatura para alcanzar carril 1 ([6:MANANA, 6:TARDE]) y asertar `HardViolations == 0`.
4. **Fallback Legacy ante Falta de Bloque Exacto (Fase 1300):** En `AgFabricaIndividuos.cs:440-454`, si `bloqueRanura == null` con `franjaExactaObligatoria`, aplicar fallback legacy y advertencia en vez de enviar a cuarentena I6.
5. **Tracking de Franjas en Sellar (Fase 1301):** En `AgCoherenciaGrupo.cs:108-153`, gestionar `franjasOcupadas` (HashSet) y registrar `WarningLimitado` si no hay bloque alterno disponible.
6. **Descarte de Duplicados en Expansor (Fase 1302):** En `AgExpansorSesiones.cs:187-206`, advertir y deduplicar sesiones duplicadas en `ExpandirIndividuo` sin cancelar todo el individuo.

### Fases Ejecutadas
- **Fase 1297 — Enum R14:** Corregido en `AgValidadorReglasTests.cs:765`.
- **Fase 1298 — R8 con Pares (Día, Franja):** Implementado en `AgValidadorReglas.cs:580-592`.
- **Fase 1299 — Test de Fábrica de Sábado:** Actualizado en `AgFabricaIndividuosTests.cs:436-468`.
- **Fase 1300 — Fallback Legacy en Fábrica:** Implementado en `AgFabricaIndividuos.cs:440-454`.
- **Fase 1301 — Sellar con HashSet de Franjas:** Refactorizado en `AgCoherenciaGrupo.cs:108-153`.
- **Fase 1302 — I7 con Deduplicación en Expansor:** Modificado en `AgExpansorSesiones.cs:187-206`.

---

## 🔧 Ronda 235: Coherencia Defensiva en Materialización y Saneamiento de Pool Docente (Fases 1303 a 1306)

### Propósito y Objetivos
1. **Habilitación de Docente en Test de Validador (Fase 1303):** En `AgValidadorReglasTests.cs:722`, registrar docente habilitado con comodín de semestre (`contexto.RegistrarDocenteHabilitado(10, 1, 100, "")`) para satisfacer R5 y permitir `HardViolations == 0`.
2. **Ejecución Defensiva Incondicional de Sellar en Staging (Fase 1304):** En `AgMaterializadorSesiones.cs:63-70`, invocar `AgCoherenciaGrupo.Sellar` para todo grupo con `genes.Count > 1`, asegurando que el desdoble de franjas de sábado actúe en staging aun cuando `Σ HorasGen == esperado`.
3. **Reserva de Franjas para Genes Fijados (Fase 1305):** En `AgCoherenciaGrupo.cs:117-128`, registrar previamente en `franjasOcupadas` las franjas de genes bloqueados (`Estado == Fijado`) en sábado para no reasignar sobre ellas.
4. **Clarificación Diagnóstica de Fallback Legacy (Fase 1306):** En `AgFabricaIndividuos.cs:465-472`, condicionar el mensaje del log a si `bloqueRanura != null` o si el gen terminará en cuarentena I6.

### Fases Ejecutadas
- **Fase 1303 — Habilitación de Docente en Test:** Implementado en `AgValidadorReglasTests.cs:722`.
- **Fase 1304 — Sellar Incondicional en Materialización:** Actualizado en `AgMaterializadorSesiones.cs:63-70`.
- **Fase 1305 — Reserva de Franjas Fijadas en Sellar:** Añadido en `AgCoherenciaGrupo.cs:117-128`.
- **Fase 1306 — Clarificación de Log en Fallback:** Corregido en `AgFabricaIndividuos.cs:465-472`.

---

## 🔧 Ronda 236: Desdoble de Sábado con Genes Fijados, Catálogo Desconectado y Ajuste de Equivalencias (Fases 1307 a 1310)

### Propósito y Objetivos
1. **Desdoble de Sábado con Genes Fijados (Fase 1307):** En `AgCoherenciaGrupo.cs:108-117`, calcular `genesSabadoTotal` considerando todos los genes del día 6 (incluyendo `Fijado`), y activar el desdoble cuando `genesSabadoTotal > 1 && genesSabado.Count > 0`. Test unitario en `AgCoherenciaGrupoTests.cs`.
2. **Resolución Robusta de Franja en Sellar (Fase 1308):** En `AgCoherenciaGrupo.cs:117-154`, crear helper `FranjaResuelta` con fallback a `ctx.Bloques`, evitar mutaciones a ciegas si el bloque no es resoluble (`null`), y emitir log diferenciado `SabadoSinFranjaLibre` vs `SabadoSinAlterna`.
3. **Restricción de franjaExactaObligatoria a Sábado (Fase 1309):** En `AgFabricaIndividuos.cs:430-438`, condicionar `diaIso == 6 && permitidasRanura.Count > 1` para no romper equivalencias legítimas `TARDE/NOCHE` de 8h en L–V.
4. **Diagnóstico Explícito sin Drop en Expansor I7 (Fase 1310):** En `AgExpansorSesiones.cs:187-210`, registrar `LoggerHelper.WarningLimitado` detallando genes en colisión sin descartar sesiones, evitando déficit de horas I2.

### Fases Ejecutadas
- **Fase 1307 — Gate de Desdoble con Fijados:** Actualizado en `AgCoherenciaGrupo.cs:108-117` y test en `AgCoherenciaGrupoTests.cs:525-550`.
- **Fase 1308 — FranjaResuelta y Log Diferenciado:** Implementado en `AgCoherenciaGrupo.cs:117-154`.
- **Fase 1309 — franjaExactaObligatoria Acotada a Sábado:** Modificado en `AgFabricaIndividuos.cs:430-438`.
- **Fase 1310 — Diagnóstico I7 sin Pérdida de Horas:** Ajustado en `AgExpansorSesiones.cs:187-210`.

---

## 🔧 Ronda 238: Cierre Falso de Ronda 237 — 0/15 Fases Aplicadas y 9 Defectos de Diseño (Fases 1326 a 1334)

### Contexto y Diagnóstico
- **Ronda 237 Reabierta:** Las fases 1311 a 1325 no fueron aplicadas en disco (0/12 archivos nuevos existen, 0 migraciones registradas). Permanecen abiertas en el backlog de implementación.
- **Defectos de Especificación:** Se identificaron 9 defectos críticos en la propuesta de diseño original que requerían corrección antes de la codificación en disco.

### Fases de Corrección de Especificación (Fases 1326 a 1334)
- **Fase 1326 — Falso Cierre con Evidencia Inexistente:** Reabrir fases 1311–1325. Establecer gate estricto de verificación `Test-Path` + `Select-String` en disco por fase. Estado: `ESPECIFICADA`.
- **Fase 1327 — Reflow de `pnlFiltros` en 2 Filas (84 px):** Rediseñar `pnlFiltros` a 84 px de alto en dos filas. Estado: `ESPECIFICADA`.
- **Fase 1328 — Integridad de `PoolDocenteKeys` y Genes Fijados:** Validar que genes fijados no usen docentes excluidos del alcance para prevenir violaciones duras R5; preservar centinela VACANTE en `AplicarAlcanceTerceros`. Estado: `ESPECIFICADA`.
- **Fase 1329 — Exclusión Explícita de Tercero Centinela VACANTE:** Filtrar centinela en consulta SQL (`id <> @IdVacante`) y en `AgTerceroFiltro` en memoria. Estado: `ESPECIFICADA`.
- **Fase 1330 — Paridad en Filtro de Subunidad:** Incluir `(subunidad = @Sub OR subunidad = '' OR subunidad IS NULL)` en consultas de terceros. Estado: `ESPECIFICADA`.
- **Fase 1331 — Supresión de Eventos en Inicialización:** Incorporar bandera `_suprimiendoEventos` durante carga de catálogos y ejecutar `RestaurarFiltrosPersistidosAsync` antes de `PoblarArbolContextoAcademico`. Estado: `ESPECIFICADA`.
- **Fase 1332 — Persistencia Síncrona en `OnFormClosing`:** Reemplazar llamadas asíncronas en el cierre por `GuardarFiltrosPersistidosSync` defensivo antes del apagado del scheduler. Estado: `ESPECIFICADA`.
- **Fase 1333 — Normalización e Índice en `ag_filtro_seleccion`:** Definir `subunidad TEXT NOT NULL DEFAULT ''` e índice único compuesto `ux_ag_filtro_seleccion(subunidad, tipo, id_ref)`. Estado: `ESPECIFICADA`.
- **Fase 1334 — API de `CheckedComboBox` Compatible:** Diseñar `CheckedComboBox` con método `Cargar<T>`, `Seleccionados<T>`, `Primario<T>` y `VacioEsTodos = true`. Estado: `ESPECIFICADA`.

---

## 🔧 Ronda 239: Ronda 238 Registrada Solo en Documentación — 9/9 Fases "ACEPTADAS" con 0 Líneas de Código (Fases 1335 a 1341)

### Contexto y Diagnóstico
- **Ausencia de Código Físico:** Verificación en disco confirma 0 líneas de código implementadas para las rondas 237 y 238.
- **Régimen de Estados:** Se formalizan los estados operativos: `ESPECIFICADA` (documentada sin código), `IMPLEMENTADA` (código en disco con `Test-Path = True` y `Select-String`), `VERIFICADA` (compilación y pruebas exitosas).
- **Parches de Integración:** Se especifican las fases 1335 a 1341 para resolver desactualización de csproj, radio de impacto de `CheckedComboBox`, truncamiento del árbol, persistencia síncrona sin deadlocks y corrección de atribuciones.

### Fases Aplicadas y Verificadas
| Fase | Estado | Evidencia en Disco |
|---|---|---|
| — | Ninguna fase aplicada en código aún | — |

### Backlog Abierto (Implementación Física Pendiente)
| Fase | Nombre | Estado | Archivo Destino |
|---|---|---|---|
| 1311 | `AgTipoProgramaCatalog` | ESPECIFICADA | `Academic\Domain\Models\AgTipoProgramaCatalog.cs` |
| 1312 / 1334 / 1337 | `CheckedComboBox` (UI + API + Compatibilidad) | ESPECIFICADA | `Academic\UI\Controls\CheckedComboBox.cs` |
| 1313 / 1327 / 1338 | Designer + Eventos + Reflow (84px, grp 326px) | ESPECIFICADA | `Academic\UI\Forms\AgFormGeneradorHorarios.Designer.cs` |
| 1314 | `AgFiltroArbolResolver` + Accesores | ESPECIFICADA | `Academic\Application\Services\AgFiltroArbolResolver.cs` |
| 1315 / 1333 | Migración V78 + `AgFiltroSeleccionRepository` | ESPECIFICADA | `Academic\Infrastructure\Persistence\Migrations\MigrationV78_AddAgFiltroSeleccion.cs` |
| 1316 / 1331 / 1332 / 1339 | Restauración, Supresión y Cierre Síncrono | ESPECIFICADA | `Academic\UI\Forms\AgFormGeneradorHorarios.cs` |
| 1317 | Horas por Nivel (4h Pregrado / 8h Posgrado) | ESPECIFICADA | `Academic\Application\Mappers\AgNodoArbolKeyMapper.cs` |
| 1318 / 1329 / 1330 / 1340 | `AgFormTerceros` (UI + SQL saneado) | ESPECIFICADA | `Academic\UI\Forms\AgFormTerceros.cs` |
| 1319 / 1329 | DTO + Lógica Pura Terceros + Filtro Vacante | ESPECIFICADA | `Academic\Application\DTOs\AgTerceroSeleccionDto.cs` |
| 1320 / 1328 | Alcance Terceros + Guarda Genes Fijados | ESPECIFICADA | `Academic\Domain\Models\Genetica\AgContexto.cs` |
| 1321 | Botones Generador y Resumen Filtros | ESPECIFICADA | `Academic\UI\Forms\AgFormGeneradorHorarios.cs` |
| 1322 | Suites de Pruebas Unitarias | ESPECIFICADA | `ConversorPDF.Tests\...` |
| 1323 / 1336 | Actualización `<Compile Include>` e Integración csproj | ESPECIFICADA | `ConversorPDF.XLL.csproj` |
| 1324 / 1335 / 1341 | Sincronización Documental y Reglas de Gate | ESPECIFICADA | `implementation_plan.md`, `walkthrough.md` |
| 1325 | Procedimiento de Aceptación Manual | ESPECIFICADA | Protocolo Operativo |
| 1342 | `SelectedIndex` read/write y carga tipada sin DataSource | ESPECIFICADA | `Academic\UI\Controls\CheckedComboBox.cs`, `AgFormGeneradorHorarios.cs` |
| 1343 | `ModoSeleccion` (Unico/Multiple) y disparo de evento | ESPECIFICADA | `Academic\UI\Controls\CheckedComboBox.cs`, `AgFormGeneradorHorarios.cs` |
| 1344 | `pnlFiltros` 66px y `SplitterDistance = 390` (árbol 189px) | ESPECIFICADA | `Academic\UI\Forms\AgFormGeneradorHorarios.Designer.cs` |
| 1345 | `EncryptedSQLiteContext.Open()` para transacción Dapper | ESPECIFICADA | `Academic\Infrastructure\Persistence\Repositories\AgFiltroSeleccionRepository.cs` |
| 1346 | Registro de `MigrationV78_AddAgFiltroSeleccion` en `SqliteHelper` | ESPECIFICADA | `Core\Helpers\SqliteHelper.cs` |
| 1347 | Saneamiento de régimen de estados (ESPECIFICADA / IMPLEMENTADA) | ESPECIFICADA | `implementation_plan.md`, `walkthrough.md` |
| 1348 | Conexión compartida sin `using` en repo (`Func<IDbConnection>`) | ESPECIFICADA | `Academic\Infrastructure\Persistence\Repositories\AgFiltroSeleccionRepository.cs` |
| 1349 | `CheckedComboBox` definición completa unificada | ESPECIFICADA | `Academic\UI\Controls\CheckedComboBox.cs` |
| 1350 | Coordenadas exactas en `pnlFiltros` sin solapamiento de labels | ESPECIFICADA | `Academic\UI\Forms\AgFormGeneradorHorarios.Designer.cs` |
| 1351 | Declaración de campos e instanciación en Designer | ESPECIFICADA | `Academic\UI\Forms\AgFormGeneradorHorarios.Designer.cs` |
| 1352 | Reconciliación de persistencia y restauración de período único | ESPECIFICADA | `Academic\UI\Forms\AgFormGeneradorHorarios.cs` |
| 1353 | Accesores `ProgramaPrimario()` y `ProgramasSeleccionados()` | ESPECIFICADA | `Academic\UI\Forms\AgFormGeneradorHorarios.cs` |
| 1354 | `PoblarArbolContextoAcademico` consume selecciones de combos | ESPECIFICADA | `Academic\UI\Forms\AgFormGeneradorHorarios.cs` |
| 1355 | Autoridad única en árbol de nivel 6 (combos son pre-filtros) | ESPECIFICADA | `Academic\UI\Forms\AgFormGeneradorHorarios.cs` |
| 1356 | Declaración de `SubunidadActual()` para persistencia/restauración | ESPECIFICADA | `Academic\UI\Forms\AgFormGeneradorHorarios.cs` |
| 1357 | `CheckedComboBox` con superficie visual completa | ESPECIFICADA | `Academic\UI\Controls\CheckedComboBox.cs` |
| 1358 | Suites de pruebas unitarias para lógica nueva (SDK-style) | ESPECIFICADA | `ConversorPDF.Tests\Academic\*.cs` |
| 1359 | Dimensionamiento fijo y AutoEllipsis en etiquetas variables | ESPECIFICADA | `Academic\UI\Forms\AgFormGeneradorHorarios.Designer.cs` |
| 1360 | Preservación de combinaciones marcadas en repoblado de árbol | ESPECIFICADA | `Academic\UI\Forms\AgFormGeneradorHorarios.cs` |
| 1361 | Guarda `_sincronizandoLista` en `CheckedComboBox.ItemCheck` | ESPECIFICADA | `Academic\UI\Controls\CheckedComboBox.cs` |
| 1362 | `AutoClose = true` y cierre por Enter/Esc en desplegable | ESPECIFICADA | `Academic\UI\Controls\CheckedComboBox.cs` |
| 1363 | Consumo de `cmbTipoPrograma` en `FiltrarProgramas` | ESPECIFICADA | `Academic\UI\Forms\AgFormGeneradorHorarios.cs` |
| 1364 | Uso del identificador canónico `toolTipConteo` | ESPECIFICADA | `Academic\UI\Forms\AgFormGeneradorHorarios.cs` |
| 1365 | Snapshot por ruta jerárquica y materialización recursiva en árbol lazy | ESPECIFICADA | `Academic\UI\Forms\AgFormGeneradorHorarios.cs` |
| 1366 | Clave compuesta jerárquica para eliminar colisiones en niveles 4-6 | ESPECIFICADA | `Academic\UI\Forms\AgFormGeneradorHorarios.cs` |
| 1367 | Supresión de `AfterCheck` (`_isUpdatingTreeChecks = true`) en restauración | ESPECIFICADA | `Academic\UI\Forms\AgFormGeneradorHorarios.cs` |
| 1368 | Preservación y restauración de ramas expandidas y scroll de árbol | ESPECIFICADA | `Academic\UI\Forms\AgFormGeneradorHorarios.cs` |

---

## 🔧 Ronda 240: Ronda 239 Registrada Sin Código y 4 Parches Técnicamente Inválidos (Fases 1342 a 1347)

### Contexto y Diagnóstico
- **Ronda 239 Documental:** 0 líneas de código implementadas; 1335–1341 permanecen en estado `ESPECIFICADA`.
- **Invalidación Técnica de Parches Previos:**
  - `SelectedIndexSet` no existe en WinForms ni en los call sites; se requiere propiedad read/write `SelectedIndex`.
  - `SelectedIndexChanged` no tenía punto de emisión; se requiere `ModoSeleccion { Unico, Multiple }` para período único y ventana semestral.
  - `grpSeleccion.Size` es inerte ante `Dock = Fill`; el ajuste real requiere `pnlFiltros` de 66 px y `SplitterDistance = 390` en `splitContenedorSim`.
  - `EncryptedSQLiteContext` no expone propiedad `Connection`; la API canónica es `using (var conn = ctx.Open())`.
  - `MigrationV78_AddAgFiltroSeleccion` carecía de orden de registro en `SqliteHelper.cs:378`.
  - El régimen de estados volvió a violarse con marcas prematuras en la tabla de principios.

---

## 🔧 Ronda 241: Parches 1342–1347 Registrados Sin Código y 6 Defectos en las Correcciones "Aceptadas" (Fases 1348 a 1353)

### Contexto y Diagnóstico
- **Régimen de Estados:** Backlog unificado (1311 a 1353) en estado `ESPECIFICADA`. 0 líneas de código implementadas en disco aún.
- **Defectos Críticos Subsanados en Especificación:**
  - Evitar disponer conexión compartida (`AcademicRepository.cs:25`): usar `Func<IDbConnection>` sin `using` local sobre `Open()`.
  - Definición unificada, coherente y compilable de `CheckedComboBox` (evita duplicidad y símbolos huérfanos).
  - Posicionamiento exacto de etiquetas y controles en `pnlFiltros` de 66 px eliminando colisiones.
  - Declaración e inicialización explícita de controles en `AgFormGeneradorHorarios.Designer.cs` y desvinculación de combos antiguos.
  - Sincronización del modelo de persistencia `ag_filtro_seleccion` con la restricción de período único (`Modo = ModoSeleccion.Unico`).
  - Accesores canónicos `ProgramaPrimario()` y `ProgramasSeleccionados()` para desacoplar los 6 call sites sensibles de `chkProgramas`.

---

## 🔧 Ronda 242: Parches 1348–1353 Registrados Sin Código y 6 Defectos Residuales (Fases 1354 a 1359)

### Contexto y Diagnóstico
- **Régimen de Estados:** Todo el backlog (Fases 1311 a 1359) permanece en `ESPECIFICADA`. 0 líneas de código implementadas en disco aún.
- **Defectos Críticos Subsanados en Especificación:**
  - Consumo efectivo de selecciones en `PoblarArbolContextoAcademico` mediante `FiltrarPeriodos`, `FiltrarSedes` y `FiltrarProgramas`.
  - Unificación de autoridad en la selección de nodos nivel 6 del árbol (combos actúan estrictamente como pre-filtros).
  - Definición explícita de `private string SubunidadActual() => SyncContext.ActiveSubunidad ?? string.Empty;`.
  - Superficie interactiva de `CheckedComboBox` con botón desplegable, `CheckedListBox` alojado en `ToolStripDropDown`.
  - Incorporación de pruebas unitarias SDK-style sin dependencias de compilación manual.
  - Dimensionamiento fijo (280×18) y `AutoEllipsis = true` con `ToolTip` en etiquetas dinámicas del panel de filtros.

---

## 🔧 Ronda 243: Fases 1354–1359 Solo Documentadas — 5 Defectos Residuales (Fases 1360 a 1364)

