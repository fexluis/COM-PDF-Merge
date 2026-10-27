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
  - Valida compuerta estricta: bloquea promoción si existen vacantes (`EsVacante == 1` o `IdTercero <= 0`) o violaciones de restricciones duras.
  - Promueve vinculaciones usando `AgVinculacionToVinculacionMapper` y sesiones fecha a fecha de forma atómica bajo transacción.
  - Preserva trazabilidad marcando el estado de la corrida como `PROMOCIONADA`.
