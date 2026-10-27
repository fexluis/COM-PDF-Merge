# Motor de Optimización de Horarios Académicos mediante Algoritmos Genéticos (AG v7)

## 1. Visión General
El motor genético académico de **ConversorPDF** optimiza la programación horaria y asignación docente para las asignaturas de las distintas subunidades académicas. Opera bajo una ventana temporal semestral acotada ($\alpha$), validando restricciones duras (R1-R7, W1) y blandas (S1-S5), con integración al pool docente, cálculo automático de viáticos desacoplado, y promoción no destructiva mediante un esquema *staging* (`ag_corridas`, `ag_individuos`, `ag_genes`, `ag_sesiones`).

---

## 2. Definición del Genoma y Cromosoma

### Estructura del Gen (`AgGen`)
Representa la asignación indivisible de una asignatura en un bloque horario dentro de un aula específica:
- `Id`: Identificador único del gen.
- `IdAsignatura`: Asignatura a programar.
- `IdTercero`: Docente habilitado en el pool.
- `IdAula`: Aula física asignada para la corrida.
- `IdBloqueHorario`: Bloque temporal (día de la semana, hora inicio, hora fin).
- `FechaInicioEstimada` y `FechaFinEstimada`: Rango temporal acotado $\alpha$ de la asignatura.
- `HorasSemanales` y `TotalHorasAsignatura`: Carga horaria requerida.
- `EsFijo`: Flag binario (1 = asignación inmutable pre-existente, 0 = mutable).

### Estructura del Individuo (`AgIndividuo`)
Cromosoma que agrupa una colección completa de genes para la subunidad y aula.
- `Genes`: Lista de genes que cubren las asignaturas requeridas.
- `Fitness`: Aptitud calculada mediante la función objetivo multitérmino.
- `HardViolations`: Número de violaciones de restricciones duras.
- `SoftViolations`: Puntaje acumulado de penalizaciones blandas.
- `EsFactible`: `true` si y solo si `HardViolations == 0`.

---

## 3. Función de Aptitud (Fitness) Formal

La aptitud de cada solución se evalúa mediante la formulación:
$$F = F_1 - F_2 + 2F_3 + F_4 + 2F_5$$

Donde:
- **$F_1$ (Cobertura):** Cantidad de genes/cursos válidamente representados en el individuo.
- **$F_2$ (Penalizaciones):** $\text{HardViolations} \times \text{PenalidadHardBase} + \text{SoftViolations}$.
- **$F_3$ (Ventana Semestral):** Bonificación por asignaciones con rango estimado $[F_{\text{inicio}}, F_{\text{fin}}]$ estrictamente contenido en la ventana semestral $\alpha$.
- **$F_4$ (Afinidad Pool y Residencia):** Bonificación cuando la residencia del docente coincide con la ciudad de la sede (minimizando viáticos de desplazamiento).
- **$F_5$ (Preservación):** Bonificación por mantener intactos los genes fijados previamente (`EsFijo == 1`).

---

## 4. Restricciones y Reglas del Sistema

### Restricciones Duras (Hard Constraints)
1. **R1 (Superposición Docente):** Un docente no puede tener dos clases simultáneas en el mismo bloque horario.
2. **R2 (Superposición Aula):** Un aula no puede albergar dos asignaturas simultáneamente en el mismo bloque horario.
3. **R3 (Superposición Grupo):** Un grupo estudiantil (`Programa + Semestre + Grupo`) no puede tener dos clases en simultáneo.
4. **R4 (Compatibilidad Aula):** El aula asignada debe estar en estado `Disponible`.
5. **R5 (Habilitación Pool):** El docente asignado debe estar activo y habilitado en la tabla `terceros_asignaturas` para la asignatura y semestre.
6. **R6 (Consistencia de Horas):** La suma de horas de las sesiones expandidas debe satisfacer `TotalHorasAsignatura`.
7. **R7 (Ocupación Externa):** El docente no puede tener sesiones aprobadas o preexistentes en otra subunidad en las mismas fechas y franjas.
8. **W1 (Ventana Semestral $\alpha$):** Ninguna sesión ni fecha estimada puede situarse fuera de $[F_{\text{inicio}}, F_{\text{fin}}]$ del semestre.

### Restricciones Blandas (Soft Constraints)
1. **S1 (Preferencia Horaria):** Penalización por bloques en días u horas no preferidas por el docente.
2. **S2 (Desplazamientos Inter-Sede):** Penalización si un docente debe trasladarse entre sedes distintas en el mismo día.
3. **S3 (Huecos Horarios):** Penalización por ventanas inactivas excesivas entre clases del mismo día.
4. **S4 (Distribución de Días):** Incentivar distribución uniforme a lo largo de la semana.
5. **S5 (Balance de Carga):** Evitar sobrecargar a un docente con más horas de las recomendadas.

---

## 5. Operadores Genéticos y Reparación

1. **Selección:** Selección por Torneo (`AgOperadorSeleccion`) con presión de selección configurable ($k = 3$).
2. **Cruce (Order Crossover - OX):** Cruce por permutación/segmentos preservando genes fijados y vacantes. Incorpora reparación en cascada mediante `AgAsignadorTercero` con búsqueda de docentes alternos habilitados y diagnóstico de vacantes (`EsVacante = 1`, `MotivoVacante`) (`AgOperadorCruce`).
3. **Mutación:** Mutación guiada con exploración de docentes habilitados en el pool, franjas horarias y bloques horarios activos (`AgOperadorMutacion`).
4. **Elitismo y Métricas:** Preservación de los mejores $N$ individuos a la siguiente generación con cálculo unificado de penalizaciones en `AgMetricas.Recalcular(individuo, contexto)`.

---

## 6. Persistencia, Mappers y Promoción (Staging a Producción)

- **Mappers Desacoplados:**
  - `AgGenToAgVinculacionMapper`: Mapeo puro de genes hacia cabeceras de vinculación en staging.
  - `AgAulaToAgVinculacionMapper`: Enriquecimiento y herencia de atributos de aula sobre la vinculación (modalidad, fechas, periodo).
  - `AgVinculacionToVinculacionMapper`: Mapeo estricto de vinculaciones de staging hacia la entidad de producción `Vinculacion`.

- **Tablas de Staging y Trazabilidad:**
  - `ag_corridas`: Metadatos de la corrida, configuración de hiperparámetros y ventana $\alpha$.
  - `ag_individuos`: Cromosomas generados en cada generación.
  - `ag_genes`: Asignaciones específicas de cada individuo con trazabilidad de vacantes y pertenencia a ventana $\alpha$.
  - `ag_vinculaciones`: Cabeceras de vinculación en staging con columnas espejo (`grupo`, `dia_horas_semanales`, `modalidadVinculacion`).
  - `ag_vinculacion_gen`: Tabla puente N:M para trazabilidad de cabeceras de vinculación y genes.
  - `ag_sesiones`: Sesiones detalladas expandidas fecha a fecha con `id_aula` explícito.
  - `ag_salvedades`: Registro de excepciones y justificaciones cuando una restricción es omitida manualmente.
  - `ag_bitacora`: Registro histórico de convergencia ($F_{\text{mejor}}$, $F_{\text{promedio}}$, $F_{\text{peor}}$, y `duracion_ms` por generación).
  - `ag_preferencia_docente`: Preferencias horarias de docentes por día (1=LUNES..7=DOMINGO) y franja.

- **Vistas Idempotentes de Staging:**
  - `vw_AgSesionesDetalle`: Vista espejo de sesiones detalladas con alias PascalCase completos.
  - `vw_AgVinculacionesDetalle`: Vista espejo de cabeceras de vinculación para consumo transparente en formularios WinForms.

- **Alineación de Esquema Defensiva:**
  - `MigrationV49_AddAgHorariosTables` ejecuta un bloque de actualización con `ALTER TABLE ... ADD COLUMN ...` idempotente que asegura compatibilidad hacia atrás en bases de datos creadas con versiones preliminares del esquema AG.

- **Promoción a Producción (`AgPromotorVinculaciones`):**
  - Valida compuerta estricta: bloquea promoción si existen vacantes (`EsVacante == 1` o `IdTercero <= 0`), aula no aprobada (`AgEstadoAula.Aprobado`) o violaciones de restricciones duras.
  - Promueve vinculaciones usando `AgVinculacionToVinculacionMapper` y sesiones fecha a fecha de forma atómica bajo transacción.
  - Implementa idempotencia en re-promoción: si `idagvinculacion` ya existe en `ag_promocion`, actualiza la vinculación de producción (`UpdateVinculacion`) y sus sesiones (`DeleteSesionesByVinculacion` + re-inserción) evitando registros duplicados.
  - Registra una fila por vinculación en `ag_promocion` vinculando `idagvinculacion` con `idvinculacion`.
  - Actualiza el estado del individuo a `AgEstadoIndividuo.Promocionado`.

---

## 7. Ciclo de Vida de Estados y Bloques Canónicos

### Estados de Gen (`AgEstadoGen`)
- `Valido (0)`: Asignación viable sin conflictos detectados.
- `Conflicto / Conflictivo (1)`: Asignación que colisiona con restricciones duras.
- `Incompleto (2)`: Gen marcado como vacante (`EsVacante == 1`, sin docente asignable).
- `Fijado (3)`: Asignación predeterminada e inmutable (`EsFijo == 1`).

### Estados de Individuo (`AgEstadoIndividuo`)
- `NoEvaluado (0)`: Estado inicial tras instanciación.
- `Factible (1)`: Evaluado con cero violaciones duras (`HardViolations == 0`).
- `Infactible (2)`: Posee una o más violaciones duras no resueltas.
- `Optimo (3)`: Mejor individuo global de la corrida con cero restricciones duras.
- `Promocionado (4)`: Individuo materializado y promovido formalmente a la base de producción.

### Diagnóstico de Vacantes (`AgMotivoVacante`)
- `SinCandidatos (1)`: El pool docente para la asignatura está vacío o sin registros activos.
- `SinDisponibilidad (2)`: Existen docentes en el pool pero todos presentan solapes horarios en la franja.
- `NoCumpleRequisitos (3)`: Existen docentes en el pool pero no están habilitados para el programa o semestre específico.
- `FueraDeVentanaSemestre (4)`: La fecha de la sesión tentativa excede la ventana temporal acotada del semestre.

### Bloques Horarios Canónicos (γ)
Definidos estrictamente para la semana académica Martes a Domingo:
- **Martes a Viernes (Días 2..5):** 1 bloque diario (18:00–22:00, 4.0h, franja `NOCHE`).
- **Sábado (Día 6):** 2 bloques diarios (08:00–12:00, 4.0h, franja `MANANA` y 13:00–17:00, 4.0h, franja `TARDE`).
- **Domingo (Día 7):** 1 bloque diario (08:00–12:00, 4.0h, franja `MANANA`).

---

## 8. Rediseño de la Ventana Temporal Semestral y Jerarquía de 6 Niveles (Fases 151 - 170)

### 8.1 Resolución de Ventana en Cascada (Fases 151 - 155)
La resolución de la ventana temporal semestral ($\alpha$) opera bajo una jerarquía estricta de prioridades en `AgResolutorVentanaSemestre`:
1. **Prioridad 1 — Fechas canónicas en `periodos` (`AgOrigenVentana.Periodos = 5`):** Lectura de `periodo_inicio` y `periodo_fin` en la tabla `periodos`. Si existen y `periodo_fin >= periodo_inicio`, se adopta inmediatamente.
2. **Prioridad 2 — Festivos presenciales (`AgOrigenVentana.Presencial = 1`):** Rango abarcado por registros en `festivos` con `tipofestivo = 'PRESENCIAL'`.
3. **Prioridad 3 — Rango manual (`AgOrigenVentana.Manual = 4`):** Fechas manuales suministradas por el usuario en el formulario.
4. **Prioridad 4 — Bloqueo informativo:** Error bloqueante descriptivo si no se configuraron fechas válidas.

La migración **V50** (`MigrationV50_AddPeriodoVentana`) incorpora las columnas `periodo_inicio TEXT` y `periodo_fin TEXT` a `periodos` mediante `ALTER TABLE` condicional y ejecuta un `UPDATE` de backfill para periodos preexistentes (IDs 1, 2 y 3).

### 8.2 Árbol Jerárquico de 6 Niveles y Tupla de Contexto $S$ (Fases 156 - 164)
- **Estructura del Árbol:** `Periodo (1) -> Sede (2) -> Programa (3) -> Semestre (4) -> Grupo (5) -> Modalidad (6)`.
- **Lazy Loading y Claves:** Carga bajo demanda en `treeContextoAcademico_BeforeExpand` utilizando `AgNodoArbolKey` para almacenar el tipo de nodo y las claves jerárquicas asociadas sin consultas prematuras a la base de datos.
- **Servicio de Selección Tri-Estado (`AgTreeViewCheckService`):** Manejo recursivo de selección ascendente (padres indeterminados/marcados) y descendente (hijos marcados/desmarcados) con conteo de combinaciones en tiempo real.
- **Expansión y Control de Cardinalidad (`AgSeleccionAulaExpandidor`):** Genera la lista plana de `AgContextoSeleccionDto`, deduplica mediante `AulaRef`, y aplica un tope de salvaguarda de $\le 500$ combinaciones para prevenir explosión combinatoria.
- **Tupla $S$ con Modalidad:** `AulaRef` incorpora la modalidad en mayúsculas (`{IdPeriodo}|{IdSede}|{IdPrograma}|{Semestre}|{Grupo}|{MODALIDAD}`), erradicando colisiones entre grupos presenciales y virtuales homónimos.
- **Arquitectura Espacial y Docking Responsivo (Ronda 204 - Fases 1039 a 1044):**
  - Reorganización modular con `SplitContainer` horizontal (`splitContenedorSim`): Panel superior expandible para selección académica y parámetros, panel inferior redimensionable para bitácora diagnóstica y acciones.
  - El control `treeContextoAcademico` adopta `Dock = Fill` dentro de `grpSeleccion`, alcanzando un área visual útil de $\ge 240$ px por defecto (hasta ~345 px mediante splitter) y erradicando espacios vacíos.
  - Distribución matemática unificada de columnas de Nivel 6 (`AgColumnasNivel6`): reparto proporcional (Jornada 30%, Días Ventana 48%, Horas/Día 22%) con anchos mínimos garantizados (62/74/64 px) sincronizando en un solo punto el dibujo (`DrawNode`), el área de clic (`ProcesarClickDimensionNivel6`), el cursor (`MouseMove`) y las cabeceras (`ActualizarCabecerasNivel6`).
  - Cabeceras desacopladas en contenedor dedicado `pnlCabecerasNivel6` con visibilidad condicional (`HayNodoNivel6Visible()`), mostrándose únicamente cuando existen ramas de modalidad desplegadas en pantalla.

### 8.3 Correcciones de Esquema, Estabilidad y Factorías (Fases 165 - 169)
- **Fase 165 (Bloqueante):** Corrección de la consulta `sqlAsig` en `AgFormGeneradorHorarios` para alinearse exclusivamente a las columnas existentes en la tabla `asignaturas` (`id, idprograma, semestre, num_nivel, cod_unidad, cod_materia, uni_teorica, metodologia, abreviatura, nucleo, nombre, creditos, int_horaria, tipoasignatura, tipocredito, prerequisito, subunidad`), resolviendo la excepción `no such column: codigo`.
- **Fase 166:** Blindaje de `TotalHorasAsignatura` en `AgFabricaIndividuos`, resolviendo en cascada `HorasObjetivo -> IntHoraria -> Horas -> Creditos*16 -> 64.0`, asegurando que el total de horas sea siempre $> 0$.
- **Fase 167:** Corrección del filtro de semestre en `sqlPool` de `AgFormGeneradorHorarios`, vinculando `semestreRama` para acotar adecuadamente el pool docente del nivel seleccionado.
- **Fase 168:** Implementación de la suite completa de pruebas unitarias para el rediseño:
  - `AgResolutorVentanaSemestreTests`: Validación de cascada de prioridades (`PERIODOS > FESTIVOS > MANUAL > Failure`).
  - `MigrationV50Tests`: Verificación de `ALTER TABLE` y backfill de fechas en tablas preexistentes.
  - `AgAulaFactoryTests`: Verificación de mapeo DTO a `AgAula` y modalidad en `AulaRef`.
  - `AgSeleccionAulaExpandidorTests`: Expansión, deduplicación y límite de combinaciones.
  - `AgTreeViewCheckServiceTests`: Propagación de estados en el árbol de 6 niveles.
  - `PeriodoTests`: Validación de entidad y método `TryObtenerVentana`.
- **Fase 169:** Endurecimiento de `AgAulaFactory` con valores predeterminados seguros (`Capacidad >= 30/999`, `TipoAula = 'ESTANDAR'/'VIRTUAL'`, `CaracteristicasJson = '{}'`, `Codigo/Nombre` no nulos) para satisfacer estrictamente las restricciones `NOT NULL` de la tabla `ag_aula`.

### 8.4 Fase 170 — Veredicto y Protocolo de Pase a Build
- **Veredicto de Build:** Favorable. La solución cumple con todas las restricciones de arquitectura limpia, .NET Framework 4.8, C# 7.3 y convención de tipo único por archivo.
- **Protocolo de Validación en Entorno Seguro:**
  1. Ejecutar compilación MSBuild del proyecto `ConversorPDF.XLL` y de la suite `ConversorPDF.Tests`.
  2. Ejecutar la suite completa de pruebas unitarias xUnit (`AgResolutorVentanaSemestreTests`, `MigrationV50Tests`, etc.).
  3. Ejecutar smoke test sobre una **copia aislada** del vault de datos antes de la promoción definitiva a producción.

### 8.5 Perfiles de Selección de Árbol y Sincronía Canónica (Fases 235 - 244)
- **Fase 235:** Carga obligatoria de `_periodos` en el catálogo inicial de `AgFormGeneradorHorarios` con orden canónico por activo e ID, previniendo combos vacíos.
- **Fase 236:** Migración `MigrationV53_AddAgArbolTables` para almacenar `ag_arbol` y `ag_arbol_nodo`, preservando parámetros evolutivos y atributos de aula por cada combinación.
- **Fase 237:** Entidades de dominio `AgArbol` y `AgArbolNodo`, y puerto de persistencia `IAgArbolRepository`.
- **Fase 238:** Repositorio transaccional `AgArbolRepository` con operaciones atómicas de guardado, sobrescritura, carga y eliminación en cascada.
- **Fase 239:** Mapeador `AgArbolMapper` y servicio `IAgArbolService` con validaciones de negocio (nombre obligatorio, $\ge 1$ nodo seleccionado).
- **Fase 240:** Interfaz de usuario con diálogo modal `FormNombreArbol` y controles de barra de herramientas (`cmbArboles`, `btnNuevoArbol`, `btnCargarArbol`, `btnSobrescribirArbol`, `btnEliminarArbol`).
- **Fase 241:** Algoritmo `AsegurarSubarbolAsync` para restaurar ramas y selecciones del árbol jerárquico bajo demanda, omitiendo o alertando huérfanos.
- **Fase 243:** Cobertura de pruebas unitarias con `AgArbolRepositoryTests`, `AgArbolServiceTests` y `MigrationV53SchemaTests`.
- **Fase 244:** Auditoría de exportación CSV con 16 columnas canónicas (`IdSesion;Fecha;Dia;Horario;Horas;Periodo;Sede;Programa;Semestre;Grupo;Modalidad;Asignatura;Docente;Aula;AulaRef;Viatica`).

### 8.6 Robustez de Perfiles, Ventana Temporal e Integridad Transaccional (Fases 245 - 252)
- **Fase 245:** Corrección de la secuencia de restauración en `CargarPerfilArbolSeleccionadoAsync`: `ActualizarVentanaSemestralAsync` se invoca estrictamente después de restaurar `chkManualVentana` y los selectores de fecha, garantizando la reproducibilidad de la ventana manual $\alpha$.
- **Fase 246:** Sincronización en caliente de `_ventanaResuelta` antes de guardar o sobrescribir perfiles con fechas manuales activas (`chkManualVentana = true`).
- **Fase 247:** Detección de filas afectadas en `AgArbolRepository.GuardarAsync` durante `UPDATE`: aborto con `tx.Rollback()` ante IDs inexistentes para erradicar la creación de nodos huérfanos en `ag_arbol_nodo`.
- **Fase 248:** Unificación del predicado SQL para unicidad y sobrescritura de perfiles: `LOWER(TRIM(nombre))` y `COALESCE(subunidad, '')`, resolviendo duplicados por casing o espacios.
- **Fase 249:** Deduplicación estricta de colecciones en `AgArbolService.GuardarPerfilAsync` por `AulaRef` para blindar el índice único `uq_ag_arbol_nodo(idarbol, aula_ref)`.
- **Fase 250:** Propagación explícita de `EsVirtual = 1` en `ObtenerCombinacionesSeleccionadas` cuando la modalidad seleccionada en el árbol es `"Virtual"`.
- **Fase 251:** Inclusión de nodo placeholder `"Cargando..."` en `AsegurarSubarbolAsync` para nodos creados manualmente, preservando la capacidad de expansión diferida en `CargarHijosLazy`.
- **Fase 252:** Compuertas de validación temprana en `GuardarNuevoPerfilArbolAsync` y `SobrescribirPerfilArbolAsync` ante períodos inválidos o ventanas temporales sin resolver.

### 8.7 Eliminación de Carreras UI, Manejo Seguro de Eventos Async y Preservación de Dominio (Fases 253 - 256)
- **Fase 253:** Introducción de la bandera `_restaurandoPerfil` en `AgFormGeneradorHorarios` para suprimir la ejecución desordenada de múltiples manejadores asíncronos concurrentes durante la carga de perfiles, garantizando una única resolución final de `_ventanaResuelta`.
- **Fase 254:** Creación del manejador centralizado `OnVentanaChanged` con protección `try / catch` y registro en `LoggerHelper.Error` para prevenir excepciones no controladas en eventos `async void`.
- **Fase 255:** Verificación de unicidad y conectividad en `AgArbolService.GuardarPerfilAsync` mediante `ExisteNombreAsync`, reportando tempranamente fallos del repositorio antes de construir entidades.
- **Fase 256:** Asignación fiel del atributo `Habilitado = n.Habilitado > 0 ? 1 : 0` en `AgArbolRepository.GuardarAsync`, respetando el valor definido en la entidad de dominio `AgArbolNodo`.

### 8.8 Control de Sobrescritura de Perfiles y Estandarización de Migración V53 (Fases 257 - 260)
- **Fase 257:** Formalización de la semántica de `ExisteNombreAsync` en `AgArbolService.GuardarPerfilAsync`, donde `dup.Value == true` ratifica la sobrescritura transaccional intencional manejada por el repositorio.
- **Fase 258:** Prevención de sobrescritura accidental en `GuardarNuevoPerfilArbolAsync` mediante advertencia interactiva `MessageBoxButtons.YesNo` si el nombre ingresado coincide con un perfil preexistente.
- **Fase 259:** Adopción explícita de `IMigrationTolerante` con `Aditiva => false` en `MigrationV53_AddAgArbolTables`, homologando la convención de `MigrationV52` y validando metadatos en pruebas unitarias.
- **Fase 260:** Asignación explícita de `Capacidad = 999` para aulas de modalidad virtual en `ObtenerCombinacionesSeleccionadas`.

### 8.9 Corrección Crítica de Esquema en Generación de Aulas y Validación Operativa (Fases 262 - 263)
- **Fase 262:** Corrección de nombres de columnas en `GenerarEncabezadosAulasAsync` para la tabla `ag_aula_asignatura`. Se reemplazaron identificadores erróneos (`idaula`, `idasignatura`, `activo`) por las columnas canónicas de `MigrationV49` (`id_aula`, `id_asignatura`, `exclusivo`, `prioridad`, `horas_objetivo`, `orden_curricular`, `requerida`, `subunidad`), eliminando la excepción `no such column: idaula` que abortaba la transacción y vaciaba `ag_aula`.
- **Fase 263:** Definición del protocolo de verificación de integridad relacional: validación de asignaturas activas por programa/semestre, persistencia exacta de `AulaRef` en `ag_aula`, vinculación no vacía en `ag_aula_asignatura` y ausencia de rollbacks en logs de ejecución.

### 8.10 Corrección de Filtro Inexistente en Tabla Asignaturas (Fases 264 - 265)
- **Fase 264:** Supresión del predicado `AND activo = 1` en la consulta `sqlAsigs` dentro de `GenerarEncabezadosAulasAsync`. Dado que la tabla `asignaturas` no contiene columna `activo` en su esquema canónico (`AcademicSchema.sql`), el filtro provocaba `no such column: activo` y abortaba la transacción de generación de aulas.
- **Fase 265:** Protocolo de validación operativa en copia de respaldo para conteo de combinaciones en perfiles y vinculación efectiva de asignaturas dependientes.

### 8.11 Fallback Defensivo en Vinculación de Asignaturas (Fases 266 - 267)
- **Fase 266:** Implementación de consulta de fallback en `GenerarEncabezadosAulasAsync` para tolerar posibles discrepancias de formato entre la tupla seleccionada (ej. semestre romano `'VIII'`) y la persistencia en la base de datos (ej. formato numérico `'8'`). Si la consulta acotada retorna 0 filas, se recuperan automáticamente las asignaturas del programa activo, impidiendo aulas huérfanas en `ag_aula_asignatura`.
- **Fase 267:** Protocolo de verificación de integridad para corridas del AG y estado de pre-vuelo.

### 8.12 Persistencia en Vault Cifrado y Restauración Automática de Perfil (Fases 268 - 269)
- **Fase 268:** Notificación de mutaciones al programador de sincronización (`SyncContext.Scheduler?.MarkDirty()`) desde `AgArbolRepository`, `AgScheduleRepository` y `GenerarEncabezadosAulasAsync`. Resuelve la pérdida crítica de datos al cerrar Excel, asegurando que el temporizador de debounce y el `Flush()` de cierre persistan y re-cifren `ag_arbol`, `ag_arbol_nodo`, `ag_aula` y `ag_aula_asignatura` en el vault `.db.enc`.
- **Fase 269:** Restauración inmediata del perfil activo en el evento `OnShown` de `AgFormGeneradorHorarios` si existen perfiles persistidos en la subunidad actual, cargando automáticamente el árbol jerárquico, la ventana $\alpha$ y los parámetros evolutivos.

### 8.13 Persistencia Forzada en Cierre y Blindaje de Repositorios (Fases 270 - 272)
- **Fase 270:** Sobrescritura de `OnFormClosing` en `AgFormGeneradorHorarios` para forzar `ForcePersistVaultAsync` y `Flush()` sincrónicamente antes de que el Ribbon destruya la sesión compartida. Protege contra cierres repentinos de Excel previos a la expiración del debounce de 5s.
- **Fase 271:** Notificación integral de `NotifyMarkDirty()` en todas las operaciones de escritura de `AgScheduleRepository` (`InsertAulaAsync`, `InsertCorridaAsync`, `UpdateCorridaAsync`, `SaveIndividuoAsync`, `SaveBitacoraAsync`, `GuardarSesionesStagingAsync`).
- **Fase 272:** Definición del protocolo de verificación de persistencia y consistencia relacional post-cierre de Excel.

### 8.14 Fallback en ConstruirContextoAsync y Diagnóstico Accionable de Catálogo (Fases 274 - 276)
- **Fase 274:** Implementación de consultas fallback en `ConstruirContextoAsync` para `Asignaturas` y `PoolDocentes`. Resuelve el bloqueo de pre-vuelo (`El aula o el programa no tiene asignaturas dependientes activas`) cuando la rama del árbol usa una nomenclatura de semestre (ej. `'I'`) distinta a la almacenada en el catálogo `asignaturas` (ej. `'1'`, vacío o nulo). Al retornar 0 filas la consulta filtrada, se ejecuta inmediatamente una consulta por `idprograma` y `subunidad` sin restricción de semestre.
- **Fase 275:** Mejora de diagnósticos para máxima accionabilidad:
  1. En `EjecutarDiagnosticoPreVueloAsync` y `btnOptimizar_Click`, se reemplazó el mensaje genérico por un detalle contextual con los parámetros de búsqueda (`idprograma`, `semestreRama`, `subunidad`), guiando al usuario si debe verificar el catálogo o cargar el pensum.
  2. En `GenerarEncabezadosAulasAsync`, se incorporó el contador explícito `sinAsignaturas`, notificando con advertencia en el diálogo final y en la bitácora si alguna de las aulas quedó huérfana de asignaturas en `ag_aula_asignatura`.
- **Fase 276:** Protocolo de verificación mediante consultas directas SQL en el vault para validar consistencia de `idprograma`, formatos de `semestre`, asignaciones en `ag_aula_asignatura` y pre-vuelo exitoso.

### 8.15 Trazabilidad y Diagnóstico de Esquema en Módulo S (Fases 277 - 283)
- **Fase 277:** Creación del motor de traza diagnóstica estructurada `AgDiagnosticoTraza` (`ConversorPDF.XLL.Academic.Application.Services.Genetica`). Provee temporización relativa continua en milisegundos (`Stopwatch`), numeración secuencial de pasos `[P01..Pnn]`, normalización de sentencias SQL, serialización segura de parámetros, conteos preventivos con detección de umbrales mínimos (`ALERTA (<min)`), volcado exhaustivo del objeto `AgContexto` y vinculación dual hacia `LoggerHelper` y el componente de bitácora UI `txtLogAG`.
- **Fase 278:** Implementación del validador preventivo de esquema SQLite `AgEsquemaDiagnostico`. Examina metadatos mediante `PRAGMA table_info` antes de que la lógica de negocio consuma tablas fundamentales (`ag_aula`, `ag_aula_asignatura`, `asignaturas`, `terceros_asignaturas`, `periodos`, `ag_corrida`, etc.), advirtiendo proactivamente en los logs sobre columnas ausentes antes de lanzar excepciones Dapper o abortar transacciones.
- **Fase 279:** Instrumentación integral paso a paso del Módulo S (`AgFormGeneradorHorarios`):
  - **P01 — Catálogos y Esquema:** Validación de 10 tablas críticas y conteo de periodos, programas y sedes activos.
  - **P02 — Ventana Semestral:** Volcado de parámetros temporales, modo de resolución (manual vs periodos) y días hábiles calculados.
  - **P03 — Árbol Académico:** Métrica de nodos generados por nivel jerárquico y total de hojas `AulaRef`.
  - **P04 — Generación de Aulas:** Combinaciones evaluadas, consultas SQL a `asignaturas`, detección de fallbacks de semestre y conteo de huérfanas `sinAsignaturas`.
  - **P05 — Contexto Genético:** Inspección detallada con SQL, parámetros y conteos de filas para cada dimensión (`ag_aula`, `ag_bloque_horario`, `ag_aula_asignatura`, `asignaturas`, `terceros_asignaturas`, `terceros`, `festivos`, `presenciales`).
  - **P06 — Índices Cruzados:** Volcado completo del contexto y cardinalidad de índices en memoria.
  - **P07 — Diagnóstico Pre-Vuelo:** Registro de factibilidad con tipificación de causa (`AgCausaFallo`) ante cualquier bloqueo.
  - **P08 — Optimización Evolutiva:** Parámetros de corrida, semilla aleatoria y métricas de convergencia y fitness.
  - **P09 — Staging:** Cantidad de individuos y sesiones provisionales materializadas.
  - **P10 — Promoción a Producción:** Transacción de inserción en tablas operativas y sincronización de pools.
  - **P11 — Cierre de Formulario:** Persistencia forzada en vault cifrado (`ForcePersistVaultAsync`) y tiempo de vaciado.
- **Fase 280:** Catálogo tipificado de causas raíz `AgCausaFallo` (`ConversorPDF.XLL.Academic.Domain.Enums`), aislando en su propio archivo las causas de bloqueo (`SinPeriodo`, `SinVentana`, `SinAulas`, `SinAnclajeAula`, `SinAsignaturas`, `SinAulaAsignaturas`, `SinPoolDocente`, `SinBloques`, `VentanaInfactible`, `EsquemaIncompleto`, etc.).
- **Fase 281:** Control de verbosidad dinámico en UI mediante la casilla de verificación `chkModoDebug` en `pnlBotonesSim`, permitiendo al operador alternar la visualización de trazas de nivel `DEBUG` en pantalla sin reiniciar la aplicación ni alterar la persistencia en el archivo físico de logs.
- **Fase 282:** Registro de los nuevos componentes en `ConversorPDF.XLL.csproj` y desarrollo de la suite de pruebas unitarias en `ConversorPDF.Tests` (`AgDiagnosticoTrazaTests` y `AgEsquemaDiagnosticoTests` con SQLite in-memory).
- **Fase 283:** Actualización de la documentación arquitectónica y operativa del motor genético.
- **Fase 284:** Incorporación de la sobrecarga `traza.Causa(AgCausaFallo, detalle)` y la propiedad `UltimaCausa` en `AgDiagnosticoTraza`, permitiendo asentar códigos de causas raíz no acompañados de excepción y eliminando el error de compilación CS1061.
- **Fase 285:** Depuración de columnas evaluadas por `AgEsquemaDiagnostico` en `CargarCatalogosAsync`, reemplazando columnas inexistentes (`nombre` en `sedes` e `id_aula` en `ag_corrida`) por las canónicas del esquema (`ciudad`, `subsede`, `activo` para `sedes` e `idaula` para `ag_corrida`), eliminando falsos positivos en los diagnósticos de arranque.
- **Fase 286:** Verificación y validación estática integral de compilación, sintaxis C# 7.3 sobre .NET Framework 4.8 y concordancia relacional con `AcademicSchema.sql` y `MigrationV49`.

### 8.16 Carga de AulaAsignaturas, Resolución de Sede en Staging y Cobertura Docente (Fases 293 - 296)
- **Fase 293:** Incorporación de la carga relacional de `ctx.AulaAsignaturas` en `ConstruirContextoAsync` para el aula anclada (`IdAulaCorrida`). Pobla el índice O(1) `ctx.AulaAsignaturasPorId` y habilita la verificación temprana en pre-vuelo (`AgCausaFallo.SinAulaAsignaturas`) para aulas huérfanas de asignaturas.
- **Fase 294:** Resolución robusta de `IdSede` en toda la cadena evolutiva y de staging: inclusión de `IdSede` en `AgContexto`, propagación determinista en `AgFabricaIndividuos`, fallback hacia aula/contexto en `AgGenToAgVinculacionMapper` y compuerta fail-fast en `AgMaterializadorSesiones` que impide la persistencia de vinculaciones con `idsede = 0`.
- **Fase 295:** Control preventivo de la regla R5 y vacantes docentes: validación de cobertura de pool docente en pre-vuelo e inicio de optimización (`AgCausaFallo.SinPoolDocente`), prevención de generación de `idDocente = 0` en `AgFabricaIndividuos` y clasificación como bloqueo crítico en `AgValidadorVentana` ante ausencia total de docentes.
- **Fase 296:** Incorporación de telemetría de cobertura (`CoberturaDocente` y `AulasSinAsignaturas`) en el motor de diagnóstico estructurado `AgDiagnosticoTraza`.

### 8.17 Acceso a Modo AG desde Módulo R y Distinción de Staging (Fases 297 - 299)
- **Fase 297:** Incorporación del panel de auditoría staging `pnlAgModo` en la pestaña Módulo R (`tabResultados`) de `AgFormGeneradorHorarios`. Provee acceso directo mediante botones dedicados: `btnAgVinculaciones` ("🔗 Vinculaciones (AG)"), `btnAgVerVinculaciones` ("📋 Detalle (AG)") y `btnAgSesiones` ("✏️ Sesión (AG)"), junto con el indicador visual `lblAgModo` ("Modo AG · staging: ag_vinculaciones / ag_sesiones").
- **Fase 298:** Cableado asíncrono y robusto de eventos desde Módulo R hacia los tres formularios académicos conmutados en `AgModoDatos.Ag`:
  - `AbrirVinculacionesModoAgAsync()`: Inicializa el grafo completo de repositorios (`VinculacionQueryRepository`, lookups y `VinculacionListService`) configurados en modo staging y presenta `FormVinculaciones` modalmente.
  - `AbrirDetalleSesionesModoAgAsync()`: Consulta `vw_ag_sesiones_detalle` con `VinculacionQueryRepository.GetSesionesAsync(...)`, tipando el resultado en `List<SesionListItem>` para inicializar `FormVerVinculaciones`.
  - `AbrirSesionModoAgAsync()`: Resuelve el `id_vinculacion` a partir del `IdSesion` seleccionado en la grilla mediante consulta parametrizada a `ag_sesiones` y abre `FormSesionesAcademicas` en modo staging.
  - Resincronización automática de resultados mediante `CargarModuloResultadosAsync` tras el cierre de cualquier diálogo modal.
- **Fase 299:** Distinción inequívoca de contexto staging en la barra de título de los formularios modales: inclusión de `if (_modo.EsModoAg) this.Text = $"{this.Text} [MODO AG - STAGING]";` en `FormSesionesAcademicas` (homologado con `FormVinculaciones` y `FormVerVinculaciones`).

### 8.18 Bloqueo Estricto por Cobertura Incompleta e Inicialización Visual (Fases 303 - 306)
- **Fase 303:** Compuerta estricta de validación del pool docente en pre-vuelo e inicio de optimización. Bloquea si existen asignaturas sin docentes habilitados en `terceros_asignaturas` (`asigSinPool > 0`), informando detalladamente la lista de asignaturas incompletas e impidiendo la generación de soluciones que incurran en violaciones duras de la regla R5 (`R5_HabilitacionDocentePool`).
- **Fase 304:** Blindaje estricto en `AgFabricaIndividuos`: salvaguarda fail-fast que previene la creación de genes con `IdTercero <= 0`, garantizando que ninguna solución inicial contenga tuplas degeneradas.
- **Fase 305:** Inicialización explícita del indicador `lblAgModo` en el constructor de `AgFormGeneradorHorarios` con texto parametrizado según las tablas canónicas de staging (`ag_vinculaciones` / `ag_sesiones`).
- **Fase 306:** Verificación de la jerarquía de acoplamiento (`Dock = DockStyle.Top`) en `tabResultados`, garantizando la disposición visual armónica sin colisiones entre el encabezado superior, el panel de Modo AG y el contenedor de resultados.

### 8.19 Encapsulamiento de Excepciones y Trazabilidad en Ejecución Evolutiva (Fases 307 - 308)
- **Fase 307:** Envoltorio integral mediante `try/catch` en `AgGeneradorHorarios.EjecutarOptimizacionAsync`. Captura `OperationCanceledException` y cualquier `Exception` no controlada, retornando `Result.Failure<AgResultadoCorrida>` seguro e impidiendo caídas del proceso de escritorio. En la capa de UI (`AgFormGeneradorHorarios.IniciarOptimizacionAsync`), se enriqueció el bloque `catch` con el registro formal `traza.Fallo("P08_EjecutarOptimizacionAsync", ex, AgCausaFallo.ExcepcionNoControlada)`.
- **Fase 308:** Protocolo de validación operativa verificando contención de fallos, trazabilidad estructurada y consistencia visual de las interfaces de optimización y staging.

### 8.20 Alineación Relacional de Esquema y Títulos en Apertura Modo AG (Fases 309 - 312)
- **Fase 309:** Corrección del identificador de columna en la consulta de resolución de vinculación: se alineó a `SELECT idvinculacion FROM ag_sesiones WHERE id = @IdSesion LIMIT 1;` conforme al DDL canónico de `MigrationV49`, eliminando la excepción `no such column: id_vinculacion`.
- **Fase 310:** Unificación del título de ventana en `FormSesionesAcademicas` a `"Sesiones Académicas [MODO AG - STAGING]"`, erradicando sufijos duplicados.
- **Fase 311:** Interpolación canónica de `modo.VistaSesionesDetalle` en las notificaciones informativas de la UI ante ausencia de registros en staging.
- **Fase 312:** Protocolo integral de verificación de integración del flujo Módulo R hacia formularios modales en Modo AG.

### 8.21 Rehidratación de Corridas y Validación Preventiva en Modo AG (Fases 313 - 315)
- **Fase 313:** Validación estricta de `idVinculacion` en `AbrirSesionModoAgAsync`. Impide la invocación de `AbrirEnModoAg(null)` ante filas sin `IdSesion` válido o inexistentes en `ag_sesiones`, informando al usuario y evitando la apertura involuntaria de borradores vacíos.
- **Fase 314:** Rehidratación persistente de soluciones en Módulo R (`_idCorridaCargada` y reconstrucción de `_ultimoResultado`):
  - Inclusión del campo `_idCorridaCargada` para rastrear la corrida visualizada en `AgFormGeneradorHorarios`.
  - Rehidratación automática de `_ultimoResultado` desde SQLite (`GetCorridaByIdAsync` y `GetIndividuoByIdAsync`) al cargar resultados de corridas históricas o tras reiniciar Excel.
  - Corrección de la compuerta en `AbrirDetalleSesionesModoAgAsync`, eliminando bloqueos falsos cuando existen filas en el grid pero `_ultimoResultado` no residía en memoria.
  - Actualización de los callbacks de cierre modal en Modo AG para refrescar la grilla evaluando `_idCorridaCargada > 0`.
  - Habilitación de las compuertas de seguridad ("Aprobar Corrida y Aula" y "Promover a Producción") para operar de forma transparente con corridas recargadas.
- **Fase 315:** Protocolo de validación operativa post-reinicio garantizando auditoría, edición unitaria y promoción estables.

### 8.22 Tratamiento Integral de Vacantes Docentes y Panel de Vacantes (Fases 317 - 324)
- **Fase 317 (Esquema BD):** Migración `MigrationV54_AddVacanteColumns` con tolerancia a fallos agregando las columnas `es_vacante INTEGER DEFAULT 0`, `motivo_vacante TEXT DEFAULT ''`, `candidatos_pool INTEGER DEFAULT 0` a `ag_vinculaciones` y `ag_sesiones`, y `es_vacante`, `motivo_vacante` a `vinculaciones` y `sesiones` en producción.
- **Fase 318 (Dominio):**
  - Creación del enum canónico `AgMotivoVacante`: `Ninguno = 0`, `SIN_CANDIDATOS = 1`, `SIN_DISPONIBILIDAD = 2`, `NO_CUMPLE_REQUISITOS = 3`, `POOL_VACIO = 4`, `FueraDeVentanaSemestre = 5`.
  - Inclusión de las propiedades `EsVacante`, `MotivoVacante` y `CandidatosPool` en `AgGen`, `AgVinculacion`, `AgSesion`, `Vinculacion` y `Sesion`.
  - Creación del DTO de presentación `AgVacanteGridDto`.
- **Fase 319 (Clasificación de Causa en Fábrica):**
  - Implementación de `AgFabricaIndividuos.ClasificarMotivo(contexto, asignatura, fecha, bloque)`.
  - En lugar de abortar mediante excepciones no controladas cuando una asignatura carece de docente habilitado, la fábrica genera genes vacantes con `IdTercero = 0`, `EsVacante = 1` y asigna su motivo específico según disponibilidad del pool.
- **Fase 320 (Relajación de Regla R5 y Métrica de Vacantes):**
  - `AgValidadorReglas` omite la violación dura `R5_HabilitacionDocentePool` cuando el gen está marcado como vacante (`EsVacante == 1` o `IdTercero <= 0`).
  - Incorporación de penalización blanda configurable `softScore += vacantes * PesoVacante` (`PesoVacante = 50.0`).
  - Métrica auditable `Vacantes` en `AgIndividuo`, `AgResultadoCorrida` y `AgCorrida`.
  - Advertencia preventiva y confirmación obligatoria al promover soluciones que contienen vacantes no resueltas.
- **Fase 321 (Pre-vuelo con Diálogo de Consentimiento):**
  - Conversión de compuerta dura en diálogo de consentimiento explícito (`MessageBoxButtons.YesNo`): ante asignaturas sin docente en pool, el usuario puede autorizar la optimización en modo vacante.
- **Fase 322 (Panel de Vacantes en Módulo R):**
  - Nueva pestaña "👤 Panel de Vacantes" en el panel inferior de auditoría con `dgvVacantes`, combo de filtrado por motivo (`cmbMotivoVacante`), contador dinámico y botón de acción contextual `btnAccionVacante`.
  - Botón contextual dinámico:
    - `SIN_CANDIDATOS` o `POOL_VACIO`: "Cargar pool" (abre `FormTerceroAsignaturas` prefiltrado por programa, asignatura y semestre).
    - `SIN_DISPONIBILIDAD`: "Mover franja" (abre `FormSesionesAcademicas` en Modo AG para la vinculación).
    - `NO_CUMPLE_REQUISITOS`: "Revisar requisitos" (abre `FormTerceroAsignaturas` para revisar salvedades/requisitos).
    - Tras resolver en diálogo modal, el Módulo R recalcula y actualiza el contador de vacantes.
- **Fase 323 (Tests Unitarios):**
  - `AgMotivoVacanteTests`: Cobertura exhaustiva de clasificación para cada escenario (`POOL_VACIO`, `SIN_CANDIDATOS`, `NO_CUMPLE_REQUISITOS`, `SIN_DISPONIBILIDAD`) y generación de genes vacantes sin excepciones.
  - `AgValidadorReglasTests`: Verificación de omisión de HardViolations en R5 y cómputo de penalización suave para vacantes.
- **Fase 324 (Documentación y Exportación CSV):**
  - Sincronización del manual técnico y ampliación del exportador CSV con columnas `Vacante`, `Motivo` y `Candidatos`.

### 8.23 Formato Unificado de Asignatura y Semestre Romano (Fases 325 - 328)
- **Fase 325 (Helper de Formato):** Implementación de `AgFormatoAsignatura` con:
  - `SemestreRomano(string)`: Normalización bidireccional y tolerante de semestre (arábigos `"2"` → `"II"`, minúsculas `"ii"` → `"II"`, nulos/vacíos → `"SIN ASIGNAR"`, passthrough para romanos extendidos o texto especial).
  - `Linea(Asignatura a)`: Proyección estandarizada `"{a.Id}: {SemestreRomano(a.Semestre)} → {a.Nombre}"`.
  - `BulletList(IEnumerable<Asignatura>)`: Unión con viñeta `"• "` para presentación en cuadros de diálogo y bitácoras.
  - Propiedad configurable `Flecha` (`"→"` por defecto, configurable a `"->"` para entornos ASCII puro).
- **Fase 326 (Integración en Puntos de Llamada):** Adopción de `AgFormatoAsignatura.BulletList` en:
  - `ValidarFactibilidadVentanaAsync`: Notificación diagnóstica pre-vuelo y diálogo modal `MessageBox.Show`.
  - `IniciarOptimizacionAsync`: Validación pre-vuelo antes de inicializar la población genética.
  - Registro estructurado en telemetría: `traza.Dato("AsignaturasSinDocente", listaSin.Replace("\n", " "))`.
- **Fase 327 (Pruebas Unitarias):** Cobertura exhaustiva en `AgFormatoAsignaturaTests` validando:
  - Casos arábigos, minúsculas, números romanos, vacíos, nulos y passthrough.
  - Formato de línea individual y listas con viñetas.
  - Configuración dinámica del separador ASCII / Unicode.
- **Fase 328 (Documentación y Auditoría):** Registro formal del formato legible en la documentación técnica del módulo académico.

### 8.24 Normalización de Motivos, Promoción Parcial con Justificación y Ajuste de UI (Fases 329 - 332)
- **Fase 329 (Normalización Determinista de `AgMotivoVacante`):**
  - Erradicación de identificadores duplicados con el mismo valor ordinal en `AgMotivoVacante`, eliminando inconsistencias causadas por llamadas no deterministas a `Enum.ToString()`.
  - Estandarización de miembros canónicos en mayúsculas: `Ninguno = 0`, `SIN_CANDIDATOS = 1`, `SIN_DISPONIBILIDAD = 2`, `NO_CUMPLE_REQUISITOS = 3`, `POOL_VACIO = 4`, `FUERA_DE_VENTANA = 5`.
  - Creación de la clase desacoplada `AgMotivoVacanteExtensions` con métodos `Codigo()` (representación textual en mayúsculas) y `Desde(string)` (conversión segura case-insensitive).
  - Adopción de `.Codigo()` en `AgAsignadorTercero`, `AgFabricaIndividuos`, `AgVacanteGridDto` y suites de pruebas unitarias.
- **Fase 330 (Promoción Parcial con Justificación de Vacantes):**
  - Ampliación de la interfaz `IAgPromotorVinculaciones` y servicio `AgPromotorVinculaciones` con la sobrecarga `PromoverAProduccionAsync(..., string justificacionVacante, ...)`.
  - Desbloqueo de la promoción para individuos con asignaturas vacantes siempre que el usuario suministre una justificación formal.
  - Asiento de `es_vacante = 1` y `motivo_vacante` en tablas de producción (`vinculaciones`, `sesiones`), así como `justificacion_vacante` en la tabla de trazabilidad `ag_promocion`.
  - Filtrado de docentes válidos (`sv.IdTercero > 0`) para prevenir la inserción de registros huérfanos o vacíos en la sincronización del pool docente.
- **Fase 331 (Población Dinámica del Combo de Motivos):**
  - Reemplazo de la lista manual hardcodeada de motivos en `AgFormGeneradorHorarios` por enumeración dinámica de `AgMotivoVacante`, incorporando automáticamente nuevos estados como `FUERA_DE_VENTANA`.
  - Acción contextual dedicada para `FUERA_DE_VENTANA`: botón "Ajustar ventana" que conduce al usuario a las herramientas de edición en modo staging.
- **Fase 332 (Verificación Visual y Espacial en Designer):**
  - Auditoría de coordenadas en el encabezado de resultados (`pnlHeaderResultados`): verificación matemática de separación libre (>120 px) entre `lblVacantes` (375,15, ancho 85 px) y `btnExportarSimulacion` (580,10), confirmando la ausencia de superposiciones o solapes.
- **Fase 333 (Acción Contextual y Navegación Dual para `FUERA_DE_VENTANA`):**
  - Cobertura explícita del motivo `FUERA_DE_VENTANA` en el Panel de Vacantes: configuración dinámica del botón contextual como `"Ajustar ventana"`.
  - Despacho de flujo interactivo dual: diálogo que ofrece abrir la vinculación/sesión en Modo AG (`FormSesionesAcademicas`) para retocar fechas individuales, o bien redirigir a la pestaña del Módulo S (`tabSimulacion`) para ampliar o revisar el rango de la ventana semestral activa.

### 8.25 Modelo de Proyección Semestral Multi-Día y Calibración con Horarios Reales (Fases 378 - 390 / 391 - 408)

Consolidación arquitectónica y algorítmica para proyectar fielmente la programación académica semestral reflejada en la matriz empírica (`Horarios.xlsx`: 219 aulas, 119 columnas-día desde 2026-08-11 hasta 2026-12-07, 9.340 horas y 2.324 sesiones distribuidas en 16 patrones semanales).

- **Fase 378 / Fase 399 (Normalización y Parsing de Semestres y Grupos):**
  - Incorporación en `AgFormatoAsignatura` de `SemestreArabigo(string)` y `ParsearSemestreGrupo(string)` con soporte de expresiones regulares y arrays compatibles con C# 7.3.
  - Conversión bidireccional tolerante: romanos (`"II"`) a arábigos (`"2"`), extracción de semestre y sufijo de grupo (ej. `"2-1"` $\rightarrow$ Semestre: `"2"`, Grupo: `"1"`) permitiendo la desambiguación de cohortes paralelas en la misma aula.

- **Fase 379 / Fase 400 (Acotamiento Estricto de Demanda en UI):**
  - Modificación de la consulta `sqlAsig` en `AgFormGeneradorHorarios.ConstruirContextoAsync`: filtrado estricto por `subunidad_id IN (SELECT DISTINCT subunidad_id FROM aulas WHERE id IN @AulasIds)`.
  - Erradicación del fallback global a toda la sede, garantizando que el optimizador solo planifique las asignaturas correspondientes a las subunidades propietarias de las aulas seleccionadas por el usuario.

- **Fase 380 / Fase 401 (Capacidad por Ocurrencias Reales y Corrección de Bloqueo Pre-Vuelo):**
  - Implementación de `AgPlanificadorTemporal.ContarOcurrencias(DayOfWeek, DateTime, DateTime, Func<DateTime, bool>)`: cálculo de ocurrencias de cada día de la semana en la ventana semestral $\alpha$ descontando festivos reales.
  - Validación de capacidad de oferta de aulas (`AgValidadorVentana.ValidarFactibilidadAsync`) basada en ocurrencias exactas (`diasOcurrencias * 4.0`), en reemplazo de aproximaciones lineales de semanas uniformes.
  - Comprobación de factibilidad individual por curso: alerta si la demanda horaria de una asignatura excede las ocurrencias máximas del día de su bloque.
  - **Eliminación del falso positivo `R-DOC` en pre-vuelo:** En la matriz real existen 98 asignaciones válidas donde un mismo docente dicta dos cursos el mismo día en franjas distintas (ej. Sábado Mañana 08:00-12:00 y Sábado Tarde 13:00-17:00). La restricción dura R1 ya previene solapamientos a nivel `(docente, fecha, franja_horaria)`.

- **Fase 381 / Fase 402 (Patrón Semanal Multi-Día y Genoma 1↔N):**
  - Creación de la entidad de dominio `AgPatronSemanal` y el servicio `AgPlanificadorPatronSemanal`.
  - Simulación calendario real (`GenerarFechas` y `Construir`): calcula iterativamente las fechas válidas para cada día del patrón semanal desde la fecha de inicio, determinando la distribución exacta de horas por día y reproduciendo fielmente los 13 histogramas semanales empíricos (ej. patrones 2-2, 2-3, 3-3, 4-4).
  - Ampliación de `AgGen`: propiedades `IndiceDiaPatron`, `PatronSemanal`, `DiasPatron`, `ClaveGrupo`, `HorasGen` y método `SesionesDelGen()`.
  - `AgFabricaIndividuos`: Asignaturas con patrones multi-día generan $N$ genes vinculados bajo la misma `ClaveGrupo` (`$"{asig.Id}|{idAula}|{grupoEfectivo}"`), cada uno con su día, bloque y carga horaria individual `HorasGen`.

- **Fase 382 / Fase 403 (Mapeo Invariante a Cabeceras y Materialización):**
  - Actualización de `AgGenToAgVinculacionMapper`: cálculo de `TotalHoras` sumando `x.HorasGen > 0 ? x.HorasGen : x.TotalHorasAsignatura`, asegurando el cumplimiento estricto del invariante I1 ($\sum \text{HorasGen} = \text{TotalHorasAsignatura}$).
  - `AgExpansorSesiones`: Proyección de fechas y sesiones acotada a `horasObjetivoGen` (`gen.HorasGen`).
  - `AgMaterializadorSesiones`: Agrupación de genes multi-día por `ClaveGrupo` en una única cabecera `ag_vinculaciones` con formato compuesto de horario: `DiaHorasSemanales = "DIA|horas, DIA|horas"`.

- **Fase 383 / Fase 404 (Validación Defensiva de Materialización y Causa de Fallo StagingVacio):**
  - Incorporación del valor `AgCausaFallo.StagingVacio = 12` en el enum canónico.
  - `AgFormGeneradorHorarios`: Verificación explícita de `matRes.IsFailure` tras la optimización.
  - Validación post-materialización: conteo de registros en `ag_vinculaciones` y `ag_sesiones` de la corrida; si no se generaron registros o la suma de horas es 0, se interrumpe el flujo y se notifica al usuario sin proceder a visualización errónea.

- **Fase 384 (Búsqueda Unificada en Pool Docente y Normalización de Semestre):**
  - Modificación de la consulta de carga del pool docente en `AgFormGeneradorHorarios` para extraer `terceros_asignaturas` correspondientes al alcance de subunidades seleccionadas.
  - Ampliación de `AgContexto.EsDocenteHabilitado`: fallback multinivel para resolución de asignaturas en el pool (clave cruda, semestre romano `"II"`, semestre arábigo `"2"`, y comodín genérico `"*"`).

- **Fase 385 (Migración de Esquema SQLite V55):**
  - Creación de `MigrationV55_AddAgGenPatronSemanal`: agregación idempotente de 5 columnas en `ag_gen` (`indice_dia_patron`, `patron_semanal`, `dias_patron`, `clave_grupo`, `horas_gen`), columna `dias_patron` en `ag_aula_asignatura`, y semilla para el bloque lunes tarde (13:00-17:00).
  - Registro de la migración en `SqliteHelper.BuildMigrator()` y en el proyecto `ConversorPDF.XLL.csproj`.

- **Fase 386 (Migración de Esquema SQLite V56):**
  - Creación de `MigrationV56_AjustarInicioPeriodo2026_02`: actualización de `periodos.periodo_inicio = '2026-08-11'` para el periodo `2026-02`, alineando el inicio semestral con el primer martes de clases del calendario institucional de 17 semanas.
  - Registro en el pipeline de migraciones.

- **Fase 387 / Fase 405 (Sincronización Honesta de UI ante Fallos):**
  - Implementación en `AgFormGeneradorHorarios.CargarResultados`: limpieza completa de DataGridViews de resultados (`dgvResultados`, `dgvSesiones`, `dgvVacantes`) y restablecimiento de etiquetas informativas a 0 cuando la corrida finaliza sin registros válidos en staging.

- **Fase 388 / Fase 406 (Encabezados de Aulas con Intensidad Horaria Real):**
  - Refactorización de `AgFormGeneradorHorarios.GenerarEncabezadosAulasAsync`: cálculo de `horas_objetivo` derivado de la intensidad horaria real `int_horaria` de la asignatura y parseo del semestre mediante `AgFormatoAsignatura.ParsearSemestreGrupo`.

- **Fase 389 / Fase 407 (Batería Exhaustiva de Pruebas Unitarias):**
  - `AgPlanificadorPatronSemanalTests`: Cobertura de derivación de fechas, distribución de horas (patrones 2-2, 3-3, 4-4, etc.) y consistencia de sumatorias.
  - `AgFormatoAsignaturaTests`: Cobertura de `SemestreArabigo` y `ParsearSemestreGrupo` para valores estándar, con sufijo de grupo, inválidos o vacíos.
  - `AgValidadorVentanaTests`: Cobertura de cómputo de capacidad por ocurrencias reales descontando festivos.
  - `PoolDocenteUnionTests`: Cobertura de fallback de claves de habilitación docente en `AgContexto`.

- **Fase 390 / Fase 408 (Exportación Espejo 1:1 a Excel de Proyección Semestral):**
  - Incorporación del botón `btnExportarProyeccion` en el diseñador de `AgFormGeneradorHorarios` con estilo visual coherente.
  - Implementación de `ExportarProyeccionEspejo()`: genera una hoja de cálculo con la estructura exacta de `Horarios.xlsx`:
    - Metadatos (Aulas, Asignatura, Docente, Horas Objetivo).
    - 119 columnas-día consecutivas rotuladas con etiqueta de día (`Ma, Mi, J, V, S, D, L`) y fecha (`dd/MM/yyyy`).
    - Matriz de sesiones proyectadas (4 horas por sesión).
    - Columna de verificación `SubTotal` con fórmula Excel `=SUMA(L{fila}:DZ{fila})`.
    - Resumen inferior de validación de horas proyectadas contra horas objetivo.

### 8.26 Ronda 52 — Estabilización, Paridad y Blindaje de Invariantes (Fases 409 - 422)

- **Fase 409 (Corrección CS1061 en Materialización):**
  - Desacoplamiento de la clave de grupo (`grp.Key` string) mediante extracción determinista de `idAulaGrupo = genes[0].IdAula`.
- **Fase 410 (Corrección de DDL en Migración V55):**
  - Reemplazo de `turno` por `franja` en el seed de `ag_bloque_horario`, previniendo excepción de SQLite al arranque.
- **Fase 411 (Derivación e Inserción de `dias_patron`):**
  - Poblamiento dinámico de `ag_aula_asignatura.dias_patron` según carga horaria ($\ge 80$h $\rightarrow$ `"2,3,4,5,6"`, $\ge 64$h $\rightarrow$ `"4,5,6"`, $<64$h $\rightarrow$ `"2,3"`), eliminando cadenas vacías.
- **Fase 412 (Degradación Defensiva a Día Único):**
  - Si la construcción del patrón semanal falla, se degrada a un único día (`diasK[0]`) para evitar multiplicar la carga horaria en múltiples genes.
- **Fase 413 y Fase 415 (Histogramas Escalonados y Semanas Relativas):**
  - Implementación en `AgPlanificadorPatronSemanal` de rampa de sesiones escalonada para reproducir los patrones empíricos observados (`[3,4,4,5]`, `[2,3,3,4,4]`, `[1,2,3,3,3]`, `[1,2,2,2]`). Agrupación por semanas ordinales relativas a `FechaInicio` garantizando ordenamiento temporal invariante al año calendario.
- **Fase 414 (Proyección Cronológica Fiel con `FechasGen`):**
  - Incorporación de `FechasGen` en `AgGen` (con soporte en `Clonar`) y consumo prioritario en `AgExpansorSesiones` para proyectar las sesiones en las fechas calculadas por el planificador semanal.
- **Fase 416 y Fase 417 (Blindaje de Invariantes I1 e I2):**
  - Comprobación matemática en fábrica y compuertas `Result.Failure` en `AgMaterializadorSesiones`:
    - I1: $\sum \text{HorasGen} == \text{TotalHorasAsignatura}$.
    - I2: $\text{TotalSesiones} \times 4.0 == \text{TotalHorasAsignatura}$.
- **Fase 418 y Fase 419 (Coherencia de Alcances y Default Unificado):**
  - `AgValidadorVentana`: Oferta calculada sobre `aulasDisponibles` (mismo alcance curricular de semestre/grupo que la demanda). Default unificado de `diasDeclarados = 2` en validador y fábrica.
- **Fase 420 (Deuda Técnica y Encapsulación):**
  - Limpieza de nombres inline en `AgContexto` y restablecimiento de visibilidad de `PoolPorAsignatura` a `{ get; private set; }`.
- **Fase 421 (Exportación Espejo 1:1 con Seriales OADate):**
  - Título A1 `"PROGRAMACION DOCENTE PERIODO 2026-02"` y seriales `ToOADate()` en columnas-día para interoperabilidad nativa de fórmulas y gráficos en Excel.
- **Fase 422 (Ampliación de Pruebas Unitarias):**
  - Cobertura de los 4 casos empíricos de histogramas escalonados y comprobación de conservación de suma de horas en `ConversorPDF.Tests`.

### 8.27 Ronda 53 — Calibración de Reparto Monótono, Integración de Festivos y Protección I1 (Fases 423 - 429)

- **Fase 423 (Conservación Estricta $\sum \text{sesiones} == S$):**
  - Guarda condicional y validación final `porSemana.Sum() == S` en `AgPlanificadorPatronSemanal`, evitando pérdidas de sesiones por saturación de tope.
- **Fase 424 (Reparto Monótono Creciente por Redondeo Acumulado):**
  - Formulación acumulativa telescópica $\lceil S \times w / W \rceil - \lceil S \times (w-1) / W \rceil$, garantizando orden no decreciente y suma exacta. Relajación de pruebas unitarias a invariantes matemáticos estables.
- **Fase 425 (Preservación de Patrones Uniformes Canónicos):**
  - El escalado solo se activa si el patrón uniforme no cabe en la ventana disponible `W`, preservando el 90% de las filas uniformes del Excel (`[2,2,2,2,2,2]`, `[2,2,2,2]`).
- **Fase 426 (Integración de Festivos en 'SemanasDisponibles'):**
  - Detección semana a semana de días hábiles no festivos en `diasK`, evitando inflación de semanas disponibles en periodos con festivos.
- **Fase 427 (Restricción de Días a 'diasK'):**
  - Acotamiento estricto de días planificados a los declarados en `diasK` y tope $n \le \min(\text{diasK.Count}, 5)$.
- **Fase 428 (Penalización Dura y Vacante ante Inconsistencia I1 en Fábrica):**
  - Penalización explícita con `individuo.HardViolations++` y marcado de `EsVacante = 1` ante discrepancias en la suma de horas de genes.
- **Fase 429 (Normalización Canónica de Horas Semanales en Día Único):**
  - Fijación de `HorasSemanales = 4.0;` e `HorasGen = totalHoras;` en la rama de día único para mantener la sesión elemental de 4 horas.

### 8.28 Ronda 54 — Orden Ascendente, Semanas Exactas y Ajuste Determinista I1 (Fases 430 - 433)

- **Fase 430 (Orden Ascendente Estable en Reparto Escalado):**
  - Ordenamiento determinista `porSemana.Sort()` tras el cálculo telescópico para garantizar monotonicidad no decreciente compatible con la progresión de semanas y validación de rango $[1..nCap]$.
- **Fase 431 (Reporte Estricto de Semanas Disponibles sin Enmascarar 0):**
  - Retorno directo de semanas hábiles en `SemanasDisponibles` (permitiendo 0) e inserción de guarda en `Construir` que falla de inmediato si la ventana no contiene días hábiles para los días K configurados.
- **Fase 432 (Ajuste Constructivo del Residuo en Gen 0):**
  - Compensación exacta de residuo de redondeo en `g0.HorasGen += (totalHoras - sumaGenes)` en `AgFabricaIndividuos`, eliminando incrementos improcedentes a `HardViolations`.
- **Fase 433 (Diagnóstico Accionable ante Insuficiencia de Franjas):**
  - Mensaje detallado ante semanas con festivos que impiden cumplir la cuota semanal, orientando al usuario a ampliar la ventana o agregar franjas.

### 8.29 Ronda 55 — Paridad Pre-Vuelo, Cuantización de Sesiones y Rotación de Franjas (Fases 434 - 437)

- **Fase 434 (Sincronización Pre-Vuelo vs Planificador en Semanas Disponibles):**
  - `AgValidadorVentana` delega el cálculo de semanas disponibles a `AgPlanificadorPatronSemanal.SemanasDisponibles`, considerando festivos y días no hábiles del patrón. Emisión de bloqueo de pre-vuelo directo si semanas disponibles es 0, unificando el cómputo de horas con la jerarquía completa de la fábrica.
- **Fase 435 (Normalización Canónica de Horas a Múltiplos de 4):**
  - En `AgFabricaIndividuos`, normalización preventiva de `totalHoras` a $\lceil \text{totalHoras} / 4.0 \rceil \times 4.0$ con advertencia en `LoggerHelper.Warning` (`[I2]`), garantizando $\text{totalHoras} == \text{sesiones} \times 4.0$ sin residuos negativos ni sesiones fraccionarias.
- **Fase 436 (Eliminación de Código Muerto 'Clamp'):**
  - Remoción de la función auxiliar estática `Clamp` en `AgPlanificadorPatronSemanal`, erradicando advertencias de análisis estático tras la consolidación del escalado.
- **Fase 437 (Rotación Determinista de Franjas Horarias):**
  - En `AgFabricaIndividuos`, selección de franjas por rotación $\text{franjasDia}[(asig.Id + i) \% N]$, eliminando la saturación sistemática de la franja Sábado MAÑANA y reduciendo conflictos R4 aula/horario en la población inicial.

### 8.30 Ronda 56 — Respeto de Patrón Declarado, Fuente Única de Días K y Limpieza de AulasAlcance (Fases 438 - 440)

- **Fase 438 (Respeto de Patrón Declarado en Catálogo):**
  - En `AgFormGeneradorHorarios`, lectura previa de `ag_aula_asignatura.dias_patron` mediante `LEFT JOIN`, preservando las combinaciones reales configuradas por el usuario o importadas desde `Horarios.xlsx` antes de aplicar la heurística fallback.
- **Fase 439 (Fuente Única Canónica de Días K - 'DiasKDe'):**
  - Centralización del parseo de `DiasPatron` en `AgPlanificadorPatronSemanal.DiasKDe(AgAulaAsignatura aa)` con retorno inmutable `IReadOnlyList<int>`, consumido idénticamente por `AgValidadorVentana` y `AgFabricaIndividuos`. Erradicación de discrepancias pre-vuelo vs corrida.
- **Fase 440 (Eliminación de Consulta Redundante de 'AulasAlcance'):**
  - Remoción de la segunda consulta asíncrona a SQLite de `AulasAlcance` en `AgFormGeneradorHorarios`, sincronizándola directamente desde `ctx.Aulas` para telemetría sin latencia adicional.

### 8.31 Ronda 57 — Dapper POCO Mapping en .NET 4.8, Supresión de 64h Ficticias y SELECT DISTINCT (Fases 441 - 443)

- **Fase 441 (Dapper POCO Mapping y 'int_horaria' Anulable):**
  - Creación de la clase auxiliar `AsigSeedRow` en `AgFormGeneradorHorarios`, reemplazando el uso de `ValueTuple` no soportado por Dapper en .NET Framework 4.8 y manejando de forma segura valores `NULL` en `int_horaria`.
- **Fase 442 (Supresión de Asignaturas sin Carga Horaria):**
  - Omisión estricta de asignaturas con `IntHoraria` nula o $\le 0$ en la siembra de `ag_aula_asignatura`, evitando inflar artificialmente la oferta con 64 horas ficticias y falsos patrones de 3 días.
- **Fase 443 (Garantía de Unicidad con 'SELECT DISTINCT'):**
  - Inclusión de `SELECT DISTINCT` en la consulta `sqlAsigs`, eliminando registros duplicados provocados por el `LEFT JOIN` cuando existían filas redundantes en `ag_aula_asignatura`.

### 8.32 Fase 444 — Desacoplamiento de Métrica 'sinHoras' del Contador 'omitidas'

- **Fase 444 (Métricas Independientes de Aulas vs Asignaturas):**
  - En `AgFormGeneradorHorarios`, introducción de la variable independiente `sinHoras` para contabilizar exclusivamente materias sin `int_horaria`, preservando `omitidas` para aulas preexistentes y ofreciendo métricas claras y transparentes en la UI, trazas y archivos de registro.

---

## 9. Plan Unificado v4 — Decisiones Explícitas de Franja/Días/Horas y Unificación Militar (Fases 648 - 659)

### 9.1 Decisiones Ratificadas de Horario y Jornada

| Jornada | Horas/Día | Franja Devuelta | Hora Inicio | Hora Fin | Duración (h) | Observaciones |
|---|---|---|---|---|---|---|
| MAÑANA | 4 | MANANA | 08:00 | 12:00 | 4.0 | Sesión matutina estándar |
| MAÑANA | 8 | MANANA | 08:00 | 17:00 | 8.0 | Sesión de día completo |
| TARDE | 4 | TARDE | 13:00 | 17:00 | 4.0 | Franja vespertina regular |
| TARDE | 8 | TARDE | 14:00 | 22:00 | 8.0 | Jornada continua sin pausa (14:00 a 22:00) |
| NOCHE | 4 | NOCHE | 18:00 | 22:00 | 4.0 | Franja nocturna regular |
| NOCHE | 8 | NOCHE | 14:00 | 22:00 | 8.0 | Comparte ventana con TARDE 8h (2:00 PM a 10:00 PM) |

> **Nota de Diseño**: Para 8 horas se programa **1 sola sesión por día** de 8 horas continuas, garantizando alineación estricta con los registros de producción.

### 9.2 Catálogo Canónico de Ventanas de Días (`AgDiasVentana`)

1. `LUN-JUE`: `1,2,3,4`
2. `MAR-VIE`: `2,3,4,5`
3. `JUE-SAB`: `4,5,6` *(Añadido para cubrir las filas J-V-S)*
4. `VIE-SAB`: `5,6`
5. `VIE-DOM`: `5,6,7`

### 9.3 Unificación a Formato Militar ("HH:mm A HH:mm")

- **Causa raíz resuelta**: `DayHoursDropdown` utilizaba cadenas de 12 horas AM/PM mientras el AG generaba en formato militar de 24 horas, provocando discrepancias de coincidencia y asignaciones `DBNull.Value`.
- **`HorarioMilitarizador`**: Normalizador estricto e idempotente para convertir cualquier representación temporal a `"HH:mm A HH:mm"`.
- **Migraciones V62 y V63**:
  - `MigrationV62_AddAgAulaVentanasYBloquesSemilla`: Expande `ag_aula` y `ag_arbol_nodo` con `jornada`, `horas_dia`, `dias_ventana` y siembra 42 bloques horarios canónicos (6 franjas × 7 días) mediante `WHERE NOT EXISTS`.
  - `MigrationV63_UnificarHorariosMilitar`: Normaliza retrospectivamente `sesiones.horariosesion`, `ag_sesiones.horariosesion` y el tercer componente de `vinculaciones.dia_horas_semanales`.

### 9.4 Protección Genética: Congelamiento Temporal (`EsFijoTemporal`)

- `AgGen.EsFijoTemporal`: Flag independiente de `Estado` que preserva la decisión explícita de aula, franja y bloque tomada por el usuario en el árbol académico.
- `Clonar()`: Replica íntegramente `EsFijoTemporal`.
- `AgOperadorMutacion`: Salta mutación de aula y bloque horario cuando `gen.EsFijoTemporal == true`, pero mantiene activa la optimización y exploración docente en el pool.
- `AgOperadorCruce (OX)`: Preserva el bloque, franja y día heredados cuando `dst.EsFijoTemporal == true`, transfiriendo exclusivamente el docente del progenitor.



