<!-- historico congelado - generado por .agents/tools/split-history.ps1 - commit d0ee6fe - no editar -->
# walkthrough - Fases 990-1107 (Rondas 192..205)

> Indice: [../../../walkthrough.md](../../../walkthrough.md) - Escision: 2026-11-16

---
### Evaluación Crítica de las Fases

#### 1. Fase 990 — Materializador elige `Min()` de bloque por día → franja reportada no es la programada
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Al materializar vinculaciones, las sesiones se agrupaban únicamente por `DiaSemana` tomando el bloque de menor ID (`idBloque = Min()`). En carriles multi-franja en un mismo día (ej. SÁBADO AM 08:00-12:00 y PM 14:00-18:00), el bloque PM era reportado con el horario AM si el ID de este último era menor. Como consecuencia, la columna K reportaba un horario falso y, al reabrir el formulario de sesiones, `RebuildSelectedDays` colapsaba a una sola franja al sobreescribir la entrada por nombre del día.
  - *Causa Raíz:* Agregación exclusivamente por día en lugar de agrupar por tupla `(DiaSemana, IdBloqueHorario)`.
  - *Impacto Confirmado:* Columna K distorsionada en carriles AM/PM y pérdida irreversible de franjas horarias al recargar.
  - *Evidencia:* `AgMaterializadorSesiones.cs:115-141`.
- **Acción Aplicada:** Reemplazada la agrupación por `DiaHorasSemanalesDe(IEnumerable<AgGen> gens)` agrupando por `new { g.DiaSemana, g.IdBloqueHorario }` ordenada por día y bloque, emitiendo una entrada canónica individual por cada franja y removiendo la advertencia espuria de múltiples bloques.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgMaterializadorSesiones.cs`.

#### 2. Fase 991 — `DiasHorasSemanaFormatter` emite 2 segmentos (inerte e incompatible con `RebuildSelectedDays`)
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* `DiasHorasSemanaFormatter` producía cadenas de 2 partes (`"DIA|horario"`). Al no contener las horas en la posición intermedia, el deserializador `RebuildSelectedDays` (`FormSesionesAcademicas.cs:2226`) fallaba en `int.TryParse(segments[1])`, descartando todas las entradas silenciosamente y dejando la columna K en blanco al reabrir.
  - *Causa Raíz:* Contrato de 2 partes incompatible con el estándar persistido de 3 partes (`DIA|horas|HORARIO`).
  - *Impacto Confirmado:* Formateador canónico inoperativo e incompatible con la reconstrucción UI.
  - *Evidencia:* `DiasHorasSemanaFormatter.cs:30-36`.
- **Acción Aplicada:** Adecuada la clave al contrato canónico de 3 partes (`$"{dia}|{horas:0.##}|{horario}"`) con fallback de horas (`s.Horas > 0 ? s.Horas : 4.0`), asegurando compatibilidad round-trip y conservando unicidad por franja.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/DiasHorasSemanaFormatter.cs`.

---

### Verificación de Principios (Ronda 191)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F990 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgMaterializadorSesiones.cs:115-141` (`DiaHorasSemanalesDe` por día y bloque) |
| F991 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `DiasHorasSemanaFormatter.cs:31-35` (contrato canónico 3 partes `dia\|horas\|horario`) |

---

## Ronda 192: F990-F991 aplicadas — `fechasesion` enmascarada con `DateTime.MinValue` (Fases 992 y 993)

### Evaluación Crítica de las Fases

#### 1. Fase 992 — Sesión sin fecha se persiste como `0001-01-01` en `ag_sesiones`
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* `FechaSesion = (s.FechaSesion ?? DateTime.MinValue).ToString("yyyy-MM-dd")` escribía `"0001-01-01"` en la columna obligatoria `fechasesion TEXT NOT NULL` de `ag_sesiones`. Esto ocasionaba que `AcademicDateSyncService.cs:35` la descartara (filtra `> 2000-01-01`), `FormVerVinculaciones.cs:863` mostrara la fecha ficticia `01/01/0001` y los viáticos y exportadores ordenaran erróneamente colocando esa sesión de primera.
  - *Causa Raíz:* Uso de centinela para evitar violar `NOT NULL` en vez de rechazar y fallar tempranamente ante datos corruptos o incompletos.
  - *Impacto Confirmado:* Propagación oculta de sesiones sin fecha a staging y reportes distorsionados.
  - *Evidencia:* `AgScheduleRepository.cs:706-728`.
- **Acción Aplicada:** Incorporada validación previa al bucle de persistencia (`sesiones.FirstOrDefault(x => !x.FechaSesion.HasValue || x.FechaSesion.Value.Year < 1900)`), realizando `trans.Rollback()` y retornando `Result.Failure` explicativo, y persistiendo directamente `s.FechaSesion.Value.ToString("yyyy-MM-dd")`.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgScheduleRepository.cs`.

#### 2. Fase 993 — Promoción a producción replica el centinela `DateTime.MinValue`
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Al copiar de staging a la tabla operativa `sesiones` en `PromoverIndividuoAsync`, se utilizaba `FechaSesion = ss.FechaSesion ?? DateTime.MinValue`. Ante una sesión malformada en staging, degradaba silenciosamente y contaminaba la tabla de producción con fechas residuales irreales sin aislamiento por corrida.
  - *Causa Raíz:* Mismo centinela `DateTime.MinValue` replicado en la ruta de promoción.
  - *Impacto Confirmado:* Contaminación de la base de datos operativa con fechas `0001-01-01`.
  - *Evidencia:* `AgPromotorVinculaciones.cs:173`.
- **Acción Aplicada:** Reemplazado el centinela por expresión `throw` explícita (`ss.FechaSesion ?? throw new Exception(...)`), asegurando un fallo ruidoso antes de persistir sesiones sin fecha válida en producción.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPromotorVinculaciones.cs`.

---

### Verificación de Principios (Ronda 192)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F992 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgScheduleRepository.cs:706-716` (guardia de fecha válida y `Value.ToString`) |
| F993 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPromotorVinculaciones.cs:173-175` (`throw` fail-fast en promoción) |

---

## Ronda 193: F994-F997 — colapso de franjas y contrato de días incompatibles (Fases 994 a 998)

### Evaluación Crítica de las Fases

#### 1. Fase 994 — Día con dos franjas (AM/PM) colapsa al reabrir la vinculación
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Al procesar vinculaciones con carriles multi-franja (ej. SÁBADO AM y PM emitidos en F990), `RebuildSelectedDays` indexaba exclusivamente por nombre del día (`result[segments[0]]`), provocando que la segunda franja sobrescribiera la primera y se perdiera la mitad de la carga horaria al persistir.
  - *Causa Raíz:* Desajuste entre el modelo multidimensional de datos y el diccionario mono-franja por día del control `DayHoursDropdown`.
  - *Impacto Confirmado:* Pérdida irreversible de franjas horarias al reabrir y guardar vinculaciones.
  - *Evidencia:* `FormSesionesAcademicas.cs:2223-2244`.
- **Acción Aplicada:** Agregada verificación de duplicados de día en `RebuildSelectedDays`, emitiendo `LoggerHelper.Warning` y reteniendo deterministamente la franja con mayor número de horas.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/FormSesionesAcademicas.cs`.

#### 2. Fase 995 — `int.TryParse` descarta horas no enteras y formato invariante
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* `int.TryParse(segments[1])` descartaba sesiones con horas fraccionarias (`3.5`). Asimismo, `{horas:0.##}` formateado bajo cultura local es-CO emitía comas decimales (`"7,5"`), fragmentando el delimitador de lista `,` y rompiendo la estructura de días.
  - *Causa Raíz:* Parseo entero para datos decimales y formateo dependiente de la cultura local.
  - *Impacto Confirmado:* Descarte silencioso de sesiones y fragmentación corrupta de cadenas `DiaHorasSemanales`.
  - *Evidencia:* `FormSesionesAcademicas.cs:2226`, `AgMaterializadorSesiones.cs:135`, `DiasHorasSemanaFormatter.cs:35`.
- **Acción Aplicada:** Actualizado `RebuildSelectedDays` con `double.TryParse(..., NumberStyles.Float, CultureInfo.InvariantCulture)` y redondeo; forzado `horas.ToString("0.##", CultureInfo.InvariantCulture)` en `AgMaterializadorSesiones` y `DiasHorasSemanaFormatter`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/FormSesionesAcademicas.cs`, `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgMaterializadorSesiones.cs`, `ConversorPDF.XLL/Academic/Application/Services/DiasHorasSemanaFormatter.cs`.

#### 3. Fase 996 — `DiasHorasSemanaFormatter` emite abreviaturas; contrato canónico usa nombres completos
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* `DiasHorasSemanaFormatter` devolvía `"LUN"`, `"MAR"`, `"MIE"`, mientras que los consumidores canónicos (`DayHoursDropdown._diasSemana`, `DayPatternParser`, `RebuildSelectedDays`) esperan nombres completos (`"LUNES"`, `"MARTES"`, ...), fallando la resolución de días.
  - *Causa Raíz:* Discrepancia entre abreviaturas del formateador y el estándar canónico del dominio.
  - *Impacto Confirmado:* Controles vacíos y excepciones de validación en UI al procesar cadenas formateadas.
  - *Evidencia:* `DiasHorasSemanaFormatter.cs:48-75`.
- **Acción Aplicada:** Delegada la resolución de nombres de día en `DiaSemanaNormalizador.ObtenerNombreDia(...)`, garantizando nombres canónicos completos sin tildes.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/DiasHorasSemanaFormatter.cs`.

#### 4. Fase 997 — Separador `" / "` en `DesdeSesiones` rompe el round-trip
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* `DesdeSesiones` unía las claves con `" / "`, impidiendo la deserialización con `raw.Split(',')` en `RebuildSelectedDays` y `DayHoursDropdown`.
  - *Causa Raíz:* Separador de presentación utilizado en lugar del separador estándar del contrato.
  - *Impacto Confirmado:* Ruptura del round-trip de serialización/deserialización.
  - *Evidencia:* `DiasHorasSemanaFormatter.cs:42`.
- **Acción Aplicada:** Reemplazado el delimitador por `", "` y anotado el recordatorio para su posterior integración en exportación y grilla.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/DiasHorasSemanaFormatter.cs`.

#### 5. Fase 998 — IDs fantasma tras rollback de F992
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Tras el rollback por sesión sin fecha válida en `GuardarStagingAsync`, las entidades `AgVinculacion` mantenían en memoria los IDs generados antes de la reversión de la transacción.
  - *Causa Raíz:* Mutación de estado en memoria no sincronizada con la reversión de base de datos.
  - *Impacto Confirmado:* Objetos en memoria con IDs huérfanos/fantasma tras fallos controlados.
  - *Evidencia:* `AgScheduleRepository.cs:708-720`.
- **Acción Aplicada:** Revertidos los IDs en memoria de las vinculaciones a sus claves locales previas antes de invocar `trans.Rollback()`.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgScheduleRepository.cs`.

---

### Verificación de Principios (Ronda 193)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F994 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `FormSesionesAcademicas.cs:2226-2240` (prevención de sobrescritura AM/PM) |
| F995 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `FormSesionesAcademicas.cs:2226`, `AgMaterializadorSesiones.cs:135`, `DiasHorasSemanaFormatter.cs:35` (`InvariantCulture`) |
| F996 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `DiasHorasSemanaFormatter.cs:46-56` (nombres completos vía `DiaSemanaNormalizador`) |
| F997 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `DiasHorasSemanaFormatter.cs:42` (delimitador canónico `", "`) |
| F998 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgScheduleRepository.cs:710-716` (restauración de IDs en memoria) |

---

## Ronda 194: F999-F1001 — `hh` mutila horarios PM y la corrupción ya está persistida

### Evaluación Crítica de las Fases

#### 1. Fase 999 — `NormalizarHorario` formatea con `hh` (12 h) → horarios PM reescritos a AM
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La normalización de franjas utilizaba `hh` (formato 12 horas) sin AM/PM en lugar de `HH` (24 horas). Esto truncaba franjas PM (ej. 13:00 a 01:00, 18:00 a 06:00). Al reabrir la vinculación en el desplegable, los horarios no coincidían con el catálogo de franjas PM, cayendo a `DBNull.Value` y borrándose al guardar.
  - *Causa Raíz:* Especificador de formato incorrecto (`hh` en lugar de `HH`) en `HorarioMilitarizador.cs` y `FranjaHoraria.cs`.
  - *Impacto Confirmado:* Corrupción de horarios vespertinos y nocturnos, provocando la pérdida del horario al editar la vinculación.
  - *Evidencia:* `HorarioMilitarizador.cs:22-23`, `FranjaHoraria.cs:34-35`.
- **Acción Aplicada:** Sustituido `hh` por `HH` en `HorarioMilitarizador.NormalizarHorario` y `FranjaHoraria.ToString()`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/HorarioMilitarizador.cs`, `ConversorPDF.XLL/Academic/Domain/Models/FranjaHoraria.cs`.

#### 2. Fase 1000 — Datos ya normalizados por V63 quedaron corruptos y la migración no se re-ejecuta
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La migración V63 persistió horarios afectados por `hh` en `sesiones` y `ag_sesiones`. Dado que 01:00 puede ser un valor válido o la corrupción de 13:00, no es posible revertirlo de manera automática sin riesgo de corrupción adicional.
  - *Causa Raíz:* Migración previa ejecutada sobre formateador con defecto de formato de 12 horas.
  - *Impacto Confirmado:* Registros persistidos con horarios AM aparentes en bases existentes.
  - *Evidencia:* `MigrationV63_UnificarHorariosMilitar.cs:36, 69`.
- **Acción Aplicada:** Implementada `MigrationV71_AuditarHorariosMeridiano` como migración no destructiva de auditoría que contabiliza filas afectadas y emite advertencia diagnóstica en `LoggerHelper.Warning`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV71_AuditarHorariosMeridiano.cs`, `ConversorPDF.XLL/Core/Helpers/SqliteHelper.cs`, `ConversorPDF.XLL/ConversorPDF.XLL.csproj`.

#### 3. Fase 1001 — El desplegable descarta el día sin emitir traza
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Ante discrepancia entre el horario asignado y los elementos del ComboBox en `DayHoursDropdown`, la celda tomaba `DBNull.Value` sin emitir traza diagnóstica, perdiendo el día de forma silenciosa al guardar.
  - *Causa Raíz:* Fallback silencioso sin instrumentación de logging en `DayHoursDropdown.cs`.
  - *Impacto Confirmado:* Descarte desapercibido de franjas horarias válidas.
  - *Evidencia:* `DayHoursDropdown.cs:182-188`.
- **Acción Aplicada:** Agregada advertencia en `LoggerHelper.Warning` reportando el día, el horario no encontrado y las opciones existentes antes de asignar `DBNull.Value`.
- **Archivo Afectado:** `ConversorPDF.XLL/UI/DayHoursDropdown.cs`.

---

### Verificación de Principios (Ronda 194)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F999 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `HorarioMilitarizador.cs:22-23`, `FranjaHoraria.cs:34-35` (`HH\\:mm`) |
| F1000 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `MigrationV71_AuditarHorariosMeridiano.cs:1-48`, `SqliteHelper.cs:372` (auditoría V71) |
| F1001 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `DayHoursDropdown.cs:182-188` (`LoggerHelper.Warning` ante descarte) |

---

## Ronda 195: F1002-F1006 — fugas de subunidad, colapso de carriles y auditoría imprecisa

### Evaluación Crítica de las Fases

#### 1. Fase 1002 — `GetBloquesHorariosAsync` ignora el parámetro `subunidad`
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La consulta SQL en `GetBloquesHorariosAsync` no filtraba por `subunidad`, permitiendo la fuga de bloques horarios entre diferentes subunidades académicas.
  - *Causa Raíz:* Parámetro recibido en la signatura del método pero omitido en el `WHERE` del comando SQL.
  - *Impacto Confirmado:* Contaminación de franjas y asignaciones horarias ajenas a la subunidad procesada.
  - *Evidencia:* `AgScheduleRepository.cs:225-231`.
- **Acción Aplicada:** Incorporada cláusula `@Subunidad = '' OR COALESCE(subunidad, '') = @Subunidad` en el `WHERE` con paso de parámetro Dapper.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgScheduleRepository.cs`.

#### 2. Fase 1003 — Carriles concurrentes colapsados por `GroupBy(...).Select(g => g.First())`
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La deduplicación por cohorte en `AgFabricaIndividuos` fusionaba aulas de la misma cohorte, descartando carriles concurrentes legítimos (ej. división `[M,X]` y `[J,V]`).
  - *Causa Raíz:* Agrupación por clave de cohorte en vez de identidad única de aula (`a.Id`).
  - *Impacto Confirmado:* Omisión de planificación del segundo carril de días en la generación de horarios.
  - *Evidencia:* `AgFabricaIndividuos.cs:88-90`.
- **Acción Aplicada:** Modificada la agrupación a `aulasValidas.GroupBy(a => a.Id).Select(g => g.First())`.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`.

#### 3. Fase 1004 — Comparación exacta `DuracionHoras == 4.0` sobre columna REAL
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Comparación de punto flotante `b.DuracionHoras == 4.0` susceptible a fallar ante pequeñas desviaciones (`3.9999999`), saltando a bloques de duración incorrecta.
  - *Causa Raíz:* Comparación estricta de igualdad sobre tipos de punto flotante provenientes de SQLite `REAL`.
  - *Impacto Confirmado:* Desviación indeseada de franjas horarias a fallbacks de 8 horas o bloques genéricos.
  - *Evidencia:* `AgFabricaIndividuos.cs:289, 360`.
- **Acción Aplicada:** Reemplazada la igualdad estricta por `Math.Abs(b.DuracionHoras - 4.0) < 0.001` en ambas ubicaciones.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`.

#### 4. Fase 1005 — Auditoría V71 con criterio demasiado amplio (falsos positivos)
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* `LIKE '% A %'` contaba todos los registros canónicos formateados, produciendo cifras falsamente infladas que oscurecían los verdaderos registros corruptos.
  - *Causa Raíz:* Condición de búsqueda genérica por delimitador en lugar de comprobación contra el catálogo de bloques.
  - *Impacto Confirmado:* Falsos positivos en logs diagnósticos de la migración V71.
  - *Evidencia:* `MigrationV71_AuditarHorariosMeridiano.cs:30-35`.
- **Acción Aplicada:** Refinado el SQL diagnóstico para contar solo filas cuyo `horariosesion` no exista en `ag_bloque_horario`.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV71_AuditarHorariosMeridiano.cs`.

#### 5. Fase 1006 — Semilla V62 introduce bloques de domingo y variantes 8 h duplicadas
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* V62 sembraba bloques para domingo (`dia = 7`) y registraba dos plantillas de 8 horas con rango idéntico `14:00-22:00` en `TARDE` y `NOCHE`.
  - *Causa Raíz:* Rango de iteración semanal inclusivo de domingos y plantillas con franja solapada.
  - *Impacto Confirmado:* Bloques de domingo innecesarios y resolución no determinista entre variantes idénticas de 8 horas.
  - *Evidencia:* `MigrationV62_AddAgAulaVentanasYBloquesSemilla.cs:44-45, 61`.
- **Acción Aplicada:** Limitado el bucle a `dia <= 6` y eliminada la variante duplicada de 8 horas en `NOCHE`.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV62_AddAgAulaVentanasYBloquesSemilla.cs`.

---

### Verificación de Principios (Ronda 195)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1002 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgScheduleRepository.cs:225-233` (filtro `@Subunidad`) |
| F1003 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFabricaIndividuos.cs:88-91` (`GroupBy(a => a.Id)`) |
| F1004 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFabricaIndividuos.cs:292, 364` (`Math.Abs(DuracionHoras - 4.0) < 0.001`) |
| F1005 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `MigrationV71_AuditarHorariosMeridiano.cs:29-39` (auditoría contra catálogo) |
| F1006 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `MigrationV62_AddAgAulaVentanasYBloquesSemilla.cs:40-45, 60` (sin domingo ni duplicado 8h) |

---

## Ronda 196: F1002-F1006 verificadas; filtro de subunidad rompe bloques globales y carriles siguen truncados (Fases 1007 a 1009)

### Evaluación Crítica de las Fases

#### 1. Fase 1007 — F1002 excluye los bloques globales (`subunidad = ''`) cuando hay subunidad activa
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Al activarse una subunidad específica en el formulario generador, `GetBloquesHorariosAsync` devolvía 0 bloques debido a que las semillas persisten con `subunidad = ''`. Al quedar `bloquesValidos` vacío, los genes se generaban sin franja horaria provocando vacantes generalizadas.
  - *Causa Raíz:* Condición estricta de igualdad en F1002 que no contemplaba el ámbito global como fallback natural.
  - *Impacto Confirmado:* Regresión en tiempo de ejecución que imposibilitaba generar horarios en cualquier subunidad distinta de vacía.
  - *Evidencia:* `AgScheduleRepository.cs:226-233`.
- **Acción Aplicada:** Ampliado el predicado SQL con `COALESCE(subunidad, '') IN ('', @Subunidad)` y ordenamiento prioritario por subunidad específica.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgScheduleRepository.cs`.

#### 2. Fase 1008 — `DiasKDe` trunca toda ventana a 2 días → un solo carril
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* En ventanas de 4 días (ej. `MAR-VIE` = `2,3,4,5`), `Take(2)` siempre seleccionaba `[2,3]`, dejando el carril `[4,5]` inalcanzable. Todas las materias de la cohorte caían en el mismo par de días colapsando la concurrencia.
  - *Causa Raíz:* Truncamiento estático sin noción de carril ni rotación en la ventana de días.
  - *Impacto Confirmado:* Imposibilidad de planificar el segundo carril semanal y solapamientos forzados de grupo.
  - *Evidencia:* `AgPlanificadorPatronSemanal.cs:42-49`, `AgFabricaIndividuos.cs:230`.
- **Acción Aplicada:** Añadido `int carril = 0` a `DiasKDe` con partición `Skip(idx * 2).Take(2)` y avance secuencial de `idxAsig` en `AgFabricaIndividuos`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorPatronSemanal.cs`, `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`.

#### 3. Fase 1009 — Semilla de domingo ya persistida y variantes 8 h duplicadas
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Bases existentes que corrieron V62 retienen bloques residuales de domingo (`dia_semana = 7`) y registros solapados de 8h `NOCHE` (`14:00-22:00`), persistiendo la posibilidad de programar domingos.
  - *Causa Raíz:* Migración V62 ya asentada en bases existentes sin script de saneamiento posterior.
  - *Impacto Confirmado:* Selección de bloques de domingo en bases existentes ante ventanas que incluyan el día 7.
  - *Evidencia:* `ag_bloque_horario` en SQLite.
- **Acción Aplicada:** Creada `MigrationV72_LimpiarBloquesDomingoYDuplicados` (versión 72) para depurar bloques globales de domingo y 8h `NOCHE`, registrándola en `SqliteHelper.cs` y `.csproj`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV72_LimpiarBloquesDomingoYDuplicados.cs`, `ConversorPDF.XLL/Core/Helpers/SqliteHelper.cs`, `ConversorPDF.XLL/ConversorPDF.XLL.csproj`.

---

### Verificación de Principios (Ronda 196)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1007 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgScheduleRepository.cs:225-231` (soporte `IN ('', @Subunidad)`) |
| F1008 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorPatronSemanal.cs:27, 46-50`, `AgFabricaIndividuos.cs:112, 230, 480` |
| F1009 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `MigrationV72_LimpiarBloquesDomingoYDuplicados.cs:1-42`, `SqliteHelper.cs:373` |

---

## Ronda 197: F1010-F1012 — `aa.DiasPatron` esquiva carriles y ventanas impares pierden día (Fases 1010 a 1012)

### Evaluación Crítica de las Fases

#### 1. Fase 1010 — `DiasKDe(aa, aula, …)` no intersecta `aa.DiasPatron` con la ventana del aula
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Al recibir `aa.DiasPatron` derivado de horas (ej. 64 h = `4,5,6`), `DiasKDe` lo devolvía crudo sin contrastar contra la ventana del aula. Un aula con ventana `MAR-VIE` (`2,3,4,5`) recibía el sábado (`6`), violando la ventana del aula y eludiendo la alternancia de carriles.
  - *Causa Raíz:* Omisión de la intersección con `aula.DiasVentana` en la rama de `AgAulaAsignatura`.
  - *Impacto Confirmado:* Sesiones asignadas en días no autorizados por la ventana del aula y fallo en la activación de carriles.
  - *Evidencia:* `AgPlanificadorPatronSemanal.cs:30-41`.
- **Acción Aplicada:** Intersectados los días de `aa.DiasPatron` con `AgDiasVentana.Dias(aula.DiasVentana)` antes de retornarlos.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorPatronSemanal.cs`.

#### 2. Fase 1011 — Ventanas impares pierden el último día (`carriles` truncado)
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La división entera `diasVentana.Count / 2` truncaba ventanas de longitud impar (`JUE-SAB` = 3 días → 1 carril), perdiendo el último día (sábado) y repitiendo el mismo carril `[4,5]`.
  - *Causa Raíz:* División entera sin redondeo hacia arriba.
  - *Impacto Confirmado:* Descarte ciego de días válidos de clase en ventanas de 3 o 5 días.
  - *Evidencia:* `AgPlanificadorPatronSemanal.cs:49-52`.
- **Acción Aplicada:** Reemplazado por `(int)Math.Ceiling(diasVentana.Count / (double)porCarril)`, habilitando el carril residual de 1 día.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorPatronSemanal.cs`.

#### 3. Fase 1012 — Validador de capacidad calcula con carril 0 para todos los carriles
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* En pre-vuelo, la estimación de capacidad se realizaba invocando `DiasKDe` sin especificar el carril, asumiendo siempre el carril 0 (de 2 días), lo que sobreestimaba la oferta en carriles residuales de 1 día.
  - *Causa Raíz:* Falta de explicitud del carril evaluado en el pre-vuelo de `AgValidadorVentana`.
  - *Impacto Confirmado:* Potencial falso negativo de capacidad cuando una materia se programa en el carril corto.
  - *Evidencia:* `AgValidadorVentana.cs:124, 161`.
- **Acción Aplicada:** Explicitado el parámetro de carril en `AgValidadorVentana.cs` para el cálculo de capacidad y validación D14.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorVentana.cs`.

---

### Verificación de Principios (Ronda 197)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1010 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorPatronSemanal.cs:39-46` (intersección con ventana) |
| F1011 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorPatronSemanal.cs:55-58` (`Math.Ceiling`) |
| F1012 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgValidadorVentana.cs:124-127, 163-165` (carril explícito) |

---

## Ronda 198: F1010-F1012 verificadas; `dias_patron` sembrado anula el reparto de carriles (Fases 1013 a 1014)

### Evaluación Crítica de las Fases

#### 1. Fase 1013 — `dias_patron` sembrado desde horas anula F1008 y excede 2 días
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* El sembrado escribía `dias_patron` con 3 a 5 días (`"2,3,4,5,6"` o `"4,5,6"`) basándose en horas. Al no quedar vacío `aa.DiasPatron`, `DiasKDe` retornaba esos días crudos e ignoraba `carril`, forzando a todas las materias a compartir los mismos días y dividiendo las horas en cargas sub-bloque de 1 a 2 horas contra bloques de 4 h u 8 h.
  - *Causa Raíz:* Asignación de patrón fijo derivado de la carga horaria en lugar de rotar de a 2 días según los carriles de la ventana del aula.
  - *Impacto Confirmado:* F1008 y la partición de carriles quedaban anulados en tiempo de ejecución, colapsando todas las asignaturas al mismo carril.
  - *Evidencia:* `AgFormGeneradorHorarios.cs:3982-3986`.
- **Acción Aplicada:** Incorporado contador `idxAulaAsig` por aula y asignado `diasPatron` de a 2 días rotado por asignatura (`idxAulaAsig % carriles`) extrayendo de `AgDiasVentana.Dias(item.DiasVentana)`.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`.

#### 2. Fase 1014 — F1012 introdujo variable inerte; capacidad sigue calculada con carril 0
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* `carrilGen = 0` era constante e inerte; la capacidad se evaluaba siempre con el carril 0 de 2 días. En ventanas impares donde el carril complementario tiene 1 día (ej. `[6]`), el pre-vuelo sobreestimaba la capacidad y no advertía al operador sobre el supuesto asumido.
  - *Causa Raíz:* Asignación constante nominal sin advertencia diagnóstica en `CapacidadDe` y validación D14.
  - *Impacto Confirmado:* Código muerto y falta de trazabilidad en validaciones pre-vuelo ante aulas multicarril.
  - *Evidencia:* `AgValidadorVentana.cs:124-127, 164`.
- **Acción Aplicada:** Reemplazado el código inerte calculando la cantidad de carriles desde la ventana del aula y emitiendo `LoggerHelper.WarningLimitado` cuando `carriles > 1` tanto en `CapacidadDe` como en la comprobación de `configs`.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorVentana.cs`.

---

### Verificación de Principios (Ronda 198)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1013 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.cs:3968-3999` (`diasPatron` de 2 días por carril) |
| F1014 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgValidadorVentana.cs:120-136, 160-172` (`WarningLimitado` si `carriles > 1`) |

---

## Ronda 199: F1013-F1014 correctas; datos legacy anulan el reparto de carriles (Fases 1015 a 1016)

### Evaluación Crítica de las Fases

#### 1. Fase 1015 — `dias_patron` legacy (3-5 días) persiste y anula F1013
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Registros creados en `ag_aula_asignatura` con lógica preexistente contenían patrones de 3 a 5 días (`"2,3,4,5,6"`, `"4,5,6"`). `INSERT OR IGNORE` no los actualizaba y `DiasKDe` retornaba por la rama de patrón fijo ignorando el carril, forzando un único carril y fragmentando la carga en 1-2 horas por día.
  - *Causa Raíz:* Datos persistidos legacy no migrados tras la implementación de F1013.
  - *Impacto Confirmado:* La partición de carriles quedaba inoperante sobre bases de datos preexistentes.
  - *Evidencia:* `ag_aula_asignatura.dias_patron`, `AgFormGeneradorHorarios.cs:3985-3988`.
- **Acción Aplicada:** Creada y registrada `MigrationV73_LimpiarDiasPatronMultidia` (versión 73) reseteando `dias_patron = ''` donde contenga más de 2 días, registrada en `SqliteHelper.cs` y `ConversorPDF.XLL.csproj`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV73_LimpiarDiasPatronMultidia.cs`, `ConversorPDF.XLL/Core/Helpers/SqliteHelper.cs`, `ConversorPDF.XLL/ConversorPDF.XLL.csproj`.

#### 2. Fase 1016 — Carril asignado sin `ORDER BY` → reparto no reproducible
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* `sqlAsigs` seleccionaba filas sin `ORDER BY`, ocasionando que la rotación `idxAulaAsig % carriles` dependiera del orden de retorno no determinista de SQLite.
  - *Causa Raíz:* Consulta de asignaturas por aula sin cláusula de ordenamiento determinista.
  - *Impacto Confirmado:* Asignación de carril oscilante entre corridas, comprometiendo la auditabilidad del smoke y la reproducibilidad.
  - *Evidencia:* `AgFormGeneradorHorarios.cs:3939-3947`.
- **Acción Aplicada:** Incorporada cláusula `ORDER BY a.id;` a `sqlAsigs`.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`.

---

### Verificación de Principios (Ronda 199)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1015 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `MigrationV73_LimpiarDiasPatronMultidia.cs:1-40`, `SqliteHelper.cs:373`, `ConversorPDF.XLL.csproj:597` |
| F1016 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.cs:3939-3947` (`ORDER BY a.id;`) |

---

## Ronda 200: F1015-F1016 verificadas; ventana ∩ patrón vacío devuelve días fuera de ventana (Fases 1017 a 1018)

### Evaluación Crítica de las Fases

#### 1. Fase 1017 — Carril residual de 1 día reduce capacidad frente a la calculada
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* En ventanas de tamaño impar (ej. `JUE-SAB` = `[4,5,6]`), el carril residual 1 tomaba únicamente 1 día (`[6]`). Dado que la capacidad pre-vuelo se calcula contra el carril 0 (de 2 días), las materias rotadas al carril residual quedaban con déficit de horas reales, degradando el fitness sin bloquear.
  - *Causa Raíz:* Truncamiento estático con `Take(porCarril)` tras `Skip(idx * porCarril)` en vez de ajustar el inicio para abarcar siempre 2 días.
  - *Impacto Confirmado:* Desequilibrio severo de capacidad entre materias de la misma aula según el carril asignado.
  - *Evidencia:* `AgPlanificadorPatronSemanal.cs:55-59`, `AgFormGeneradorHorarios.cs:3994-3996`.
- **Acción Aplicada:** Ajustado el offset con `Math.Min(idx * porCarril, Math.Max(0, count - porCarril))` en `AgPlanificadorPatronSemanal.cs` y en el sembrado de `AgFormGeneradorHorarios.cs`, asegurando que `JUE-SAB` asigne `[4,5]` y `[5,6]`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorPatronSemanal.cs`, `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`.

#### 2. Fase 1018 — Patrón fuera de ventana se devuelve sin filtrar
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Cuando `aa.DiasPatron` no tenía intersección con la ventana del aula (ej. patrón previo `[2,3]` y ventana cambiada a `JUE-SAB`), `dentro.Count == 0` retornaba la lista original sin filtrar, programando sesiones en días no autorizados.
  - *Causa Raíz:* Fallback permisivo en `DiasKDe` que conservaba el patrón disyunto ante intersección vacía.
  - *Impacto Confirmado:* Fuga de sesiones fuera de la ventana declarada del aula.
  - *Evidencia:* `AgPlanificadorPatronSemanal.cs:41-47`.
- **Acción Aplicada:** Implementado fallback que emite `LoggerHelper.WarningLimitado` y resuelve el carril de 2 días sobre `ventana` ante patrones disyuntos.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorPatronSemanal.cs`.

---

### Verificación de Principios (Ronda 200)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1017 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorPatronSemanal.cs:55-61`, `AgFormGeneradorHorarios.cs:3995-4001` (carril residual 2 días) |
| F1018 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorPatronSemanal.cs:38-51` (`WarningLimitado` y carril de ventana) |

---

## Ronda 201: F1017-F1018 verificadas; solape de carriles rompe R4 y overload textual sin paridad (Fases 1019 a 1020)

### Evaluación Crítica de las Fases

#### 1. Fase 1019 — Overload textual sin paridad F1018 → días fuera de ventana en cruce
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La sobrecarga textual `DiasKDe(string, string, AgAula)` no contenía la salvaguarda de F1018; si `dentro.Count == 0`, retornaba los días pedidos disyuntos de la ventana. En el operador de cruce (`AgOperadorCruce.cs:166`) esto transmitía días ajenos a la ventana del aula destino en el descendiente. Asimismo, la herencia de `aaSrc` en cruce no suministraba el aula destino.
  - *Causa Raíz:* Asimetría entre los overloads textual y tipado de `DiasKDe`.
  - *Impacto Confirmado:* Fuga de días fuera de ventana durante la operación de cruce genético.
  - *Evidencia:* `AgPlanificadorPatronSemanal.cs:96-107`, `AgOperadorCruce.cs:168`.
- **Acción Aplicada:** Incorporado parámetro opcional `int carril = 0`, aviso `LoggerHelper.WarningLimitado` delegando en `DiasKDe((AgAulaAsignatura)null, aula, carril)` ante patrones fuera de ventana, y pasado `AulaSrcDst(dst.IdAula)` en `AgOperadorCruce.cs`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorPatronSemanal.cs`, `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgOperadorCruce.cs`.

#### 2. Fase 1020 — Carriles solapados violan R4 (misma aula/grupo, mismo día)
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* El intento en F1017 de forzar 2 días en el carril residual mediante clamp solapado (`[4,5]` y `[5,6]`) causaba que dos asignaturas de la cohorte compartieran el día 5, incurriendo en infracción dura `R4_DosAsignaturasMismoDia` y solapamiento de aula R2.
  - *Causa Raíz:* Solapamiento artificial de días entre carriles en ventanas de longitud impar.
  - *Impacto Confirmado:* Penalizaciones masivas y estancamiento evolutivo en aulas con ventanas impares.
  - *Evidencia:* `AgPlanificadorPatronSemanal.cs:66-71`, `AgValidadorReglas.cs:286-300`.
- **Acción Aplicada:** Revertido el solape a carriles disjuntos usando división entera (`diasVentana.Count / porCarril`), emitiendo advertencia `LoggerHelper.WarningLimitado` cuando el día residual queda fuera de rotación, y espejado en el sembrado de `AgFormGeneradorHorarios.cs`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorPatronSemanal.cs`, `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`.

---

### Verificación de Principios (Ronda 201)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1019 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorPatronSemanal.cs:83-108`, `AgOperadorCruce.cs:168` |
| F1020 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorPatronSemanal.cs:62-75`, `AgFormGeneradorHorarios.cs:3994-4001` (carriles disjuntos) |

---

## Ronda 202: F1019-F1020 verificadas; tolerancia a ISO numérico en ventana y diferimiento de patrón sin ventana (Fases 1021 a 1022)

### Evaluación Crítica de las Fases

#### 1. Fase 1021 — `AgDiasVentana.Dias` sin tolerancia a ISO numérico
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* `AgDiasVentana.AIso` rechazaba cadenas que contenían directamente listas de números ISO (ej. `"1,2,3,4"` o `"2,3,4,5"`) porque solo buscaba en el diccionario por claves de nombres (`"LUN-JUE"`). `Dias` devolvía lista vacía, causando que `DiasKDe` recurriera al valor por defecto `{2,3}` y `AgValidadorVentana` computara `carriles = 1`.
  - *Causa Raíz:* Falta de parser numérico de respaldo en `AgDiasVentana.AIso`.
  - *Impacto Confirmado:* Desconexión de la ventana real del aula en bases que almacenan números ISO directos.
  - *Evidencia:* `AgDiasVentana.cs:22-38`.
- **Acción Aplicada:** Agregado parseo secundario para números ISO [1..7] separados por coma o punto y coma, devolviendo el formato canónico ordenado.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgDiasVentana.cs`.

#### 2. Fase 1022 — Sembrado con ventana vacía fuerza `4,5`/`2,3`
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Cuando un aula no poseía `DiasVentana` especificado, el sembrado asignaba patrones estáticos `"4,5"` o `"2,3"` en `ag_aula_asignatura`. Si el aula operaba en una jornada distinta o no hábil para esos días, el patrón introducía rigidez artificial.
  - *Causa Raíz:* Forzado de patrón estático predeterminado en ausencia de ventana declarada.
  - *Impacto Confirmado:* Patrones fuera de disponibilidad sembrados rígidamente en la base de datos.
  - *Evidencia:* `AgFormGeneradorHorarios.cs:4002-4004`.
- **Acción Aplicada:** Establecido `diasPatron = string.Empty` cuando no hay ventana declarada y registrada traza `AulaAsignaturaSinVentana` para diferir la resolución al motor.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`.

---

### Verificación de Principios (Ronda 202)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1021 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgDiasVentana.cs:22-34` (tolerancia ISO numérico) |
| F1022 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.cs:4001-4006` (`diasPatron = string.Empty;`) |

---

## Ronda 203: Colisión de nombre Excel 'De' (Fases 1033 a 1038)

### Evaluación Crítica de las Fases

#### 1. Fase 1033 — Eliminar el nombre Excel 'De' en AgSemanaIso y AgClaveCohorte
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Excel-DNA 1.8.0 registra por defecto todos los métodos estáticos públicos de clases públicas cuyo tipo de retorno y tipos de parámetros pertenecen al conjunto soportado, derivando el nombre Excel directamente del nombre del método sin prefijo de clase. `AgSemanaIso.De(DateTime)` y `AgClaveCohorte.De(...)` colisionaban registrando ambas la función `=DE(...)`.
  - *Causa Raíz:* Registro automático por reflexión en Excel-DNA 1.8.0 que procesa todas las clases públicas y exporta métodos compatibles usando únicamente el nombre del método como identificador Excel.
  - *Impacto Confirmado:* Mensaje de error en el log de Excel-DNA: `Repeated function name: 'De' - previous registration will be overwritten.` donde la última función sobrescribe a la primera.
  - *Evidencia:* `AgSemanaIso.cs:30`, `AgClaveCohorte.cs:11-36`, `AgFabricaIndividuos.cs:84`.
- **Acción Aplicada:** Eliminado el método alias `De(DateTime)` en `AgSemanaIso.cs` (manteniendo `DiaIso(DateTime)` como API canónica). Renombradas las sobrecargas en `AgClaveCohorte.cs` a `DesdeCohorte`, `DesdeAula` y `DesdeGen`. Actualizada la llamada local en `AgFabricaIndividuos.cs:84` a `AgClaveCohorte.DesdeAula(a)`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgSemanaIso.cs`, `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgClaveCohorte.cs`, `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`.

#### 2. Fase 1034 — Test guardián de registro de métodos Excel-DNA
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Falta de test de regresión automatizado que replique las reglas de introspección de Excel-DNA 1.8.0 sobre `ConversorPDF.XLL.dll` para prevenir reintroducción de funciones duplicadas o con el nombre reservado `De`.
  - *Causa Raíz:* Ausencia de prueba unitaria de arquitectura para el descubrimiento de UDFs en el ensamble XLL.
  - *Impacto Confirmado:* Riesgo de nuevas colisiones no detectadas hasta el despliegue y apertura en Excel.
  - *Evidencia:* Nuevo archivo `ConversorPDF.Tests/Unit/ExcelDnaRegistrationGuardTests.cs`.
- **Acción Aplicada:** Creado `ExcelDnaRegistrationGuardTests.cs` con xUnit y FluentAssertions, validando que ningún método público estático registrable comparta nombre con otro (`StringComparer.OrdinalIgnoreCase`) y que ningún método registrable se llame `De`.
- **Archivo Afectado:** `ConversorPDF.Tests/Unit/ExcelDnaRegistrationGuardTests.cs`.

#### 3. Fase 1036 — `DesdeCohorte` queda expuesto como UDF inválido (6 argumentos requeridos)
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* `DesdeCohorte(int, int, int, string, string aulaRef = null)` se exponía como UDF en Excel exigiendo todos los parámetros de manera obligatoria. Excel-DNA 1.8.0 no evalúa `IsOptional` ni valores por defecto, por lo que las llamadas documentadas con menos argumentos fallaban en Excel.
  - *Causa Raíz:* Ausencia de soporte de parámetros opcionales en el registro automático de Excel-DNA 1.8.0 y visibilidad `public` en un helper de dominio.
  - *Impacto Confirmado:* Registro innecesario de un helper de dominio como UDF defectuosa en hojas de cálculo.
  - *Evidencia:* `AgClaveCohorte.cs:13`.
- **Acción Aplicada:** Cambiada visibilidad a `internal static string DesdeCohorte(...)`.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgClaveCohorte.cs`.

#### 4. Fase 1037 — El test guardián no replica las reglas de Excel-DNA 1.8.0
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* El test guardián excluía accesores de propiedades `get_`/`set_` estáticos públicos mediante `IsSpecialName` y omitía filtros de arrays y genéricos abiertos de Excel-DNA 1.8.0, evaluando 68 métodos en lugar de los 75 reales registrados.
  - *Causa Raíz:* Regla ad-hoc en el test que no replicaba fielmente `AssemblyLoader.GetExcelMethods` y `IsMethodSupported`.
  - *Impacto Confirmado:* Falsos negativos potenciales y subconteo en la inspección de métodos UDF.
  - *Evidencia:* `ExcelDnaRegistrationGuardTests.cs:44-110`.
- **Acción Aplicada:** Reemplazados los filtros en `EsMetodoRegistrableExcelDna180` y `ObtenerMetodosRegistrables` para reproducir exactamente la lógica de Excel-DNA 1.8.0.
- **Archivo Afectado:** `ConversorPDF.Tests/Unit/ExcelDnaRegistrationGuardTests.cs`.

#### 5. Fase 1038 — Trazabilidad y datos incorrectos en documentación
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Ausencia de la Ronda 203 en `implementation_plan.md` y `walkthrough.md` de la raíz del repositorio; cita incorrecta de excepción en lugar del log real de sobrescritura de Excel-DNA; y omisión de advertencia de error `#NAME?` para fórmulas que usaban `=De(...)`.
  - *Causa Raíz:* Desincronización documental entre carpetas y asunciones teóricas no validadas con la fuente de Excel-DNA.
  - *Impacto Confirmado:* Inconsistencias en la trazabilidad del proyecto y en la documentación técnica para usuarios y desarrolladores.
  - *Evidencia:* `implementation_plan.md`, `walkthrough.md`, `docs/articles/changes/2026-10-22-colision-nombre-funcion-de.md`.
- **Acción Aplicada:** Actualizados los artefactos de la raíz del workspace y corregido el artículo de cambios reflejando el comportamiento exacto de Excel-DNA 1.8.0 y advertencias de compatibilidad.
- **Archivos Afectados:** `implementation_plan.md`, `walkthrough.md`, `docs/articles/changes/2026-10-22-colision-nombre-funcion-de.md`.

---

### Verificación de Principios (Ronda 203)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1033 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | AgSemanaIso.cs:27-29, AgClaveCohorte.cs:11-37, AgFabricaIndividuos.cs:84-85 |
| F1034 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | ExcelDnaRegistrationGuardTests.cs:1-135 |
| F1036 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | AgClaveCohorte.cs:10-15 (`internal static string DesdeCohorte`) |
| F1037 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | ExcelDnaRegistrationGuardTests.cs:44-105 (reglas 1.8.0 exactas) |
| F1038 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | implementation_plan.md, walkthrough.md, 2026-10-22-colision-nombre-funcion-de.md |

---

## Ronda 204: Corrección de Solapamientos Espaciales y Extensión de `treeContextoAcademico` (Fases 1039 a 1044)

### Evaluación Crítica de las Fases

#### 1. Fase 1039 — Splitter Horizontal y Reorganización de `tabSimulacion`
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La pestaña `tabSimulacion` acumulaba paneles apilados con `Dock=Bottom` fijos (82 + 58 + 118 = 258 px), comprimiendo el contenedor `grpSeleccion` a 357 px y confinando el árbol `treeContextoAcademico` a solo 138 px de altura útil con ~105 px de fondo gris inútil sin posibilidad de ajuste por el usuario.
  - *Causa Raíz:* Jerarquía rígida de acoplamiento (`Dock=Bottom`) sobre el contenedor de pestaña sin divisor dinámico.
  - *Impacto Confirmado:* Área de navegación académica asfixiada en resoluciones estándar (1120x720 / 1366x768).
  - *Evidencia:* `AgFormGeneradorHorarios.Designer.cs:149-178`, `AgFormGeneradorHorarios.cs:105-112`.
- **Acción Aplicada:** Incorporado `splitContenedorSim` (`SplitContainer`, `Orientation=Horizontal`, `Dock=Fill`, `FixedPanel=Panel2`, `SplitterWidth=6`) como único hijo de `tabSimulacion`. Distribuidos `grpSeleccion` y `pnlConfiguracionInferior` en `Panel1`, y `pnlBotonesSim` y `txtLogAG` en `Panel2`. En `OnLoad`, se calibra dinámicamente `SplitterDistance = Math.Max(280, Height - 300)`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.Designer.cs`, `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`.

#### 2. Fase 1040 — `treeContextoAcademico` con `Dock=Fill` y Toolbar Desacoplada sin Colisiones
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* `treeContextoAcademico` usaba coordenadas y anclajes fijos perdiendo altura. En `pnlArbolToolbar`, `lblConteoCombinaciones` en X=738 con `AutoSize=true` crecía hasta ~928 px solapando por ~39 px a `btnGenerarAulas` (X=889) y tapándolo al pintarse encima por orden de z-order.
  - *Causa Raíz:* Controles posicionados por coordenadas absolutas (`Location`) con `AutoSize=true` sin contención elástica en la barra de herramientas del árbol.
  - *Impacto Confirmado:* Texto de conteo sobreescrito visualmente sobre el botón de acción principal de generación.
  - *Evidencia:* `AgFormGeneradorHorarios.Designer.cs:180-214`, `392-417`.
- **Acción Aplicada:** `treeContextoAcademico` pasa a `Dock=Fill` con `ItemHeight=20`, `BorderStyle=FixedSingle`, `HideSelection=false`, `ShowNodeToolTips=true` y `Scrollable=true`. En `pnlArbolToolbar`, se agruparon los botones en `pnlToolbarIzq` (`FlowLayoutPanel`, `Dock=Left`), `btnGenerarAulas` se fijó con `Dock=Right` (150 px), y `lblConteoCombinaciones` se configuró con `Dock=Fill`, `AutoSize=false`, `AutoEllipsis=true`, `TextAlign=MiddleLeft` y `toolTipConteo`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.Designer.cs`, `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`.

#### 3. Fase 1041 — Cabeceras Nivel 6 Desacopladas y Cálculo Proporcional Unificado
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Las cabeceras de Nivel 6 (`lblColJornada`, `lblColDias`, `lblColHorasDia`) residían en `pnlArbolToolbar` en Y=36 con coordenadas fijas (253, 478, 893), mostrándose siempre visibles aún sin ramas de modalidad desplegadas, y desfasadas respecto a los 11 offsets mágicos en `DrawNode`, `ProcesarClickDimensionNivel6` y `MouseMove`.
  - *Causa Raíz:* Falta de fuente única de verdad para el cálculo métrico y dispersión de coordenadas en el código-behind y diseñador.
  - *Impacto Confirmado:* Desalineación milimétrica entre las cabeceras y los radio-botones dibujados en los nodos hoja, con riesgo de clics erróneos en DPIs distintos de 100%.
  - *Evidencia:* `AgFormGeneradorHorarios.Designer.cs:215-278`, `AgFormGeneradorHorarios.cs:2405-2446`, `3163-3179`.
- **Acción Aplicada:** Creado `pnlCabecerasNivel6` (`Dock=Top`, alto 20, `Visible=false`) con separadores verticales de 1 px (`sepNivel6A`, `sepNivel6B`). Implementado struct `AgColumnasNivel6` y método unificado `CalcularColumnasNivel6()` con reparto proporcional 30% / 48% / 22% y anchos mínimos (62, 74, 64 px). Sincronizados `DrawNode`, `ProcesarClickDimensionNivel6`, `MouseMove` y `ActualizarCabecerasNivel6()` usando el struct, condicionando visibilidad a `HayNodoNivel6Visible()`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.Designer.cs`, `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`.

#### 4. Fase 1042 — Reorganización de Barra de Acciones y Depuración de Emojis
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* `pnlAccionesSim` mantenía 700 px fijos dejando 89 px desaprovechados y apenas 7 px entre `chkModoDebug` y los botones. Con fuentes sin glifos emoji (`🔍`, `▶`) en Segoe UI 9 o bajo DPI 125%, el ensanchamiento causaba solapes sobre la barra de progreso y el estado.
  - *Causa Raíz:* Ancho estático en contenedor de acciones y caracteres emoji no soportados nativamente en fuentes GDI+ estándar.
  - *Impacto Confirmado:* Textos cortados y solapes sobre controles de monitoreo en ejecución.
  - *Evidencia:* `AgFormGeneradorHorarios.Designer.cs:723-833`.
- **Acción Aplicada:** Reorganizado `pnlBotonesSim` con `pnlProgreso` (`Dock=Left`, 380 px, conteniendo `progressBarAG` y `lblStatusAG` con `AutoEllipsis=true`), `chkModoDebug` (`Dock=Right`, 180 px), y las acciones en `flowAcciones` (`FlowLayoutPanel`, `Dock=Right`, `FlowDirection=RightToLeft`, `AutoSize=true`). Retirados emojis de los botones (`"Diagnóstico Pre-vuelo"`, `"Optimizar Horarios"`, `"Cancelar"`).
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.Designer.cs`, `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`.

#### 5. Fase 1043 — Estandarización Tipográfica, DPI Base y Suscripciones de Redimensión
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* El formulario omitía definición de `Font` propia, heredando MS Sans Serif 8.25 mientras los controles hijos especificaban Segoe UI 9, generando distorsión en `AutoScaleDimensions = (6F, 13F)` al escalar en monitores con DPI 125%/150%. Además, contraer nodos o redimensionar el splitter dejaba cabeceras huérfanas sin actualizar.
  - *Causa Raíz:* Inconsistencia métrica entre la fuente base del formulario y las fuentes de sus controles, junto a omisión de eventos de colapso y cambio de tamaño.
  - *Impacto Confirmado:* Cabeceras visibles permanentemente tras contraer el árbol y desalineación en pantallas de alta densidad.
  - *Evidencia:* `AgFormGeneradorHorarios.Designer.cs:1151-1156`, `AgFormGeneradorHorarios.cs:242-245`, `3187-3198`.
- **Acción Aplicada:** Asignado en el formulario `Font = new Font("Segoe UI", 9F)`, `AutoScaleDimensions = (7F, 15F)`, `ClientSize = (1120, 720)` y `MinimumSize = (1120, 700)`. Suscritos `treeContextoAcademico.AfterCollapse`, `treeContextoAcademico.ClientSizeChanged` y `splitContenedorSim.SplitterMoved` a `ActualizarCabecerasNivel6()`. Enriquecido `ObtenerXBaseNivel6` con salvaguarda `Math.Max(..., ClientSize.Width * 0.09)`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.Designer.cs`, `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`.

#### 6. Fase 1044 — Sincronización Documental y Trazabilidad en Base de Conocimiento
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Necesidad de reflejar las mejoras de arquitectura espacial de UI en la documentación oficial del motor genético (§8.2) y sincronizar la trazabilidad de la Ronda 204 en los registros de cambios del repositorio.
  - *Causa Raíz:* Desincronización entre la implementación de interfaz y los manuales técnicos.
  - *Impacto Confirmado:* Falta de registro oficial de las decisiones de diseño espacial y métricas para mantenimiento futuro.
  - *Evidencia:* `docs/articles/changes/2026-ag-horarios-academicos.md:155-160`, `implementation_plan.md:6564-6615`.
- **Acción Aplicada:** Actualizada la subsección §8.2 de `docs/articles/changes/2026-ag-horarios-academicos.md` documentando el modelo de docking responsivo, el SplitContainer y la distribución paramétrica de Nivel 6. Consolidadas las Fases 1039 a 1044 en `implementation_plan.md` y `walkthrough.md`.
- **Archivos Afectados:** `docs/articles/changes/2026-ag-horarios-academicos.md`, `implementation_plan.md`, `walkthrough.md`.

---

### Verificación de Principios (Ronda 204)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1039 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.Designer.cs:149-178`, `AgFormGeneradorHorarios.cs:105-112` |
| F1040 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.Designer.cs:180-214, 392-417`, `AgFormGeneradorHorarios.cs:3410-3439` |
| F1041 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.Designer.cs:215-278`, `AgFormGeneradorHorarios.cs:2405-2446, 3163-3179` |
| F1042 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.Designer.cs:723-833` |
| F1043 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.Designer.cs:1151-1156`, `AgFormGeneradorHorarios.cs:242-245, 3187-3198` |
| F1044 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `docs/articles/changes/2026-ag-horarios-academicos.md:155-160`, `implementation_plan.md`, `walkthrough.md` |

---

## Ronda 205: Corrección de Contenedor de Componentes Nulo, Calibración de Splitter, Anti-Reentrada y Separadores (Fases 1045 a 1051)

### Evaluación Crítica de las Fases

#### 1. Fase 1045 — Instanciación de `components` en `InitializeComponent`
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* `new System.Windows.Forms.ToolTip(this.components)` con `components == null` lanzaba `ArgumentNullException` en tiempo de ejecución, provocando que el formulario no pudiera instanciarse y bloqueando el Módulo S.
  - *Causa Raíz:* El campo `components` se declaraba `= null` y ningún punto del parcial de diseño lo inicializaba antes del constructor del ToolTip.
  - *Impacto Confirmado:* Error fatal de inicialización que impedía abrir `AgFormGeneradorHorarios`.
  - *Evidencia:* `AgFormGeneradorHorarios.Designer.cs:17`.
- **Acción Aplicada:** Agregada la inicialización `this.components = new System.ComponentModel.Container();` en la primera línea de `InitializeComponent()`.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.Designer.cs`.

#### 2. Fase 1046 — Calibración de Splitter y Corrección de `grpSeleccion.MinimumSize`
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La fórmula `Height - 300` asignaba un espacio excesivo a `Panel2` (294 px) limitando el árbol a 174 px útiles, mientras que `grpSeleccion.MinimumSize = (0, 240)` provocaba desborde de 7 px sobre `pnlConfiguracionInferior` en resolución mínima (1120x700).
  - *Causa Raíz:* Constante estática de splitter no ajustada a la altura real de `Panel2` (239 px) y tamaño mínimo de GroupBox incompatible con el panel contenedor.
  - *Impacto Confirmado:* Reintroducción del solape entre el GroupBox y los parámetros de configuración al redimensionar a tamaño mínimo.
  - *Evidencia:* `AgFormGeneradorHorarios.cs:105-121`, `AgFormGeneradorHorarios.Designer.cs:192`.
- **Acción Aplicada:** Establecido `splitContenedorSim.Panel1MinSize = 300`, `splitContenedorSim.Panel2MinSize = 120`, y calibrado `SplitterDistance = alto - altoPanel2 - anchoSeparador` (objetivo: 229 px para el árbol). En Designer, ajustado `grpSeleccion.MinimumSize = new Size(0, 190)`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`, `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.Designer.cs`.

#### 3. Fase 1047 — Protección Anti-Reentrada en `ActualizarCabecerasNivel6`
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La mutación de `pnlCabecerasNivel6.Visible` alteraba el alto del árbol disparando `treeContextoAcademico.ClientSizeChanged`, lo que volvía a invocar `ActualizarCabecerasNivel6` durante el layout provocando parpadeo y ciclos de relayout redundantes.
  - *Causa Raíz:* Ausencia de bandera de reentrada al mutar la visibilidad de un control acoplado hermano.
  - *Impacto Confirmado:* Múltiples relayouts y parpadeo visible al expandir/contraer ramas de nivel 6 y arrastrar el splitter.
  - *Evidencia:* `AgFormGeneradorHorarios.cs:2416, 3182-3204`.
- **Acción Aplicada:** Incorporada bandera de bloqueo `private bool _alineandoCabeceras;` con bloque `try/finally` y asignación condicional `if (pnlCabecerasNivel6.Visible != hayVisible)`.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`.

#### 4. Fase 1048 — Unificación a Recorrido Único de Búsqueda de Nivel 6
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* `ActualizarCabecerasNivel6` ejecutaba dos recorridos recursivos O(n) por cada evento de cambio de tamaño (`HayNodoNivel6Visible` y `ObtenerXBaseNivel6`).
  - *Causa Raíz:* Falta de propagación del nodo `TreeNode n6` ya localizado hacia el cálculo de columnas.
  - *Impacto Confirmado:* Sobrecarga de cómputo en la interfaz durante redimensionamiento y expansión.
  - *Evidencia:* `AgFormGeneradorHorarios.cs:2418, 3185-3211`.
- **Acción Aplicada:** Refactorizado `BuscarPrimerNodoNivel6` para ejecutarse una sola vez por actualización, pasando `n6` directamente a `CalcularColumnasNivel6(n6)` y `ObtenerXBaseNivel6(n6)`.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`.

#### 5. Fase 1049 — Compensación de Borde de 1 px para `BorderStyle.FixedSingle`
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* `DrawNode` dibuja en coordenadas cliente internas del árbol mientras `lblCol*` y `sepNivel6*` se posicionan en el panel exterior; el borde de 1 px producía un desfase horizontal perceptible entre las cabeceras y los radio-botones.
  - *Causa Raíz:* Inconsistencia entre el origen cliente y el origen del panel contenedor sin compensación métrica.
  - *Impacto Confirmado:* Desalineación visual de 1 px de las etiquetas respecto a sus opciones interactivas.
  - *Evidencia:* `AgFormGeneradorHorarios.cs:3196-3201`.
- **Acción Aplicada:** Introducida constante `const int BordeArbol = 1;` restando 1 px en el cálculo de `Left` para todas las etiquetas y separadores de Nivel 6.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`.

#### 6. Fase 1050 — Restauración de Separador de Toolbar en `pnlToolbarIzq`
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* El retiro de los paneles de separación runtime en F1040 dejó los 8 botones de la toolbar izquierda agrupados sin delimitación visual entre las acciones de persistencia y las de navegación del árbol.
  - *Causa Raíz:* Omisión de un separador nativo al migrar controles a `FlowLayoutPanel`.
  - *Impacto Confirmado:* Degradación de la jerarquía visual y accesibilidad en la barra de herramientas del árbol.
  - *Evidencia:* `AgFormGeneradorHorarios.Designer.cs:51, 303, 311-318, 1233`.
- **Acción Aplicada:** Declarado e instanciado `sepToolbar1` (`Panel`, 1x20 px, margen `(3, 5, 3, 0)`) insertado entre `btnEliminarArbol` y `btnExpandirArbol`.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.Designer.cs`.

#### 7. Fase 1051 — Corrección de Citas de Evidencia y Sincronización Documental
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La tabla de verificación de F1040 citaba `3326-3350` en lugar de la ubicación real `3410-3439` de `ActualizarConteoCombinaciones`, y la estimación de extensión requería precisión acorde a la calibración de F1046.
  - *Causa Raíz:* Desfase de numeración de líneas por adición de métodos y estimación estática previa.
  - *Impacto Confirmado:* Desincronía en las evidencias de auditoría del walkthrough.
  - *Evidencia:* `walkthrough.md:5604`.
- **Acción Aplicada:** Sincronizada la cita a `AgFormGeneradorHorarios.cs:3410-3439`, actualizada la métrica de extensión a 138 px → 229 px (objetivo F1046) y consolidados los registros en `walkthrough.md` e `implementation_plan.md`.
- **Archivos Afectados:** `walkthrough.md`, `implementation_plan.md`.

---

### Verificación de Principios (Ronda 205)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1045 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.Designer.cs:17` (`components = new Container()`) |
| F1046 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.cs:105-121`, `AgFormGeneradorHorarios.Designer.cs:192` (`MinimumSize=(0,190)`) |
| F1047 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.cs:2416, 3182-3204` (`_alineandoCabeceras`) |
| F1048 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.cs:2418, 3185-3211` (`n6` propagado) |
| F1049 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.cs:3196-3201` (`- BordeArbol`) |
| F1050 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.Designer.cs:51, 303, 311-318, 1233` (`sepToolbar1`) |
| F1051 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `walkthrough.md:5604`, `implementation_plan.md` |

---

## 📦 Entrega de Release v1.6.3.64 (Instalador Oficial)

- **Versión Oficial:** `1.6.3.64`
- **Archivos de Versión Sincronizados:**
  - `ConversorPDF.XLL\Core\Helpers\App.cs` (`APP_VERSION = "1.6.3.64"`)
  - `ConversorPDF.XLL\ConversorPDF.XLL.csproj` (`<Version>1.6.3.64</Version>`)
  - `installer\InstaladorConversorPDF.wxs` (`Version="1.6.3.64"`, `DefaultVersion="!(bind.FileVersion.filXllDll)"`)
- **Verificación Integral de Pruebas Unitarias:**
  - Suite completa: **717 / 717 pruebas aprobadas (0 errores, 57 s)** (`dotnet test ConversorPDF.Tests\ConversorPDF.Tests.csproj -c Release`).
- **Compilación de Add-In:**
  - `ConversorPDF.XLL` (Release|x64): **0 Errores** (`ConversorPDF.XLL-AddIn64.xll`).
- **Paquete Instalador Generado (`/build-release 1.6.3.64`):**
  - Archivo: `installer\InstaladorConversorPDFv1.6.3.64.msi`
  - Tamaño: **5,603,328 bytes** (5.60 MB)
  - Hash SHA-256: `1BAF0C1ED92490354098363FC722116D4A6BCB98E0DC7D0888D48D66F4DCA003`
  - Estado: Listo para distribución con auto-activación en Excel, calibración elástica de UI y blindaje completo de componentes.

---

## 🔬 Ronda 206 — Sistema de Carriles Disjuntos y Oleadas Secuenciales (Fases 1052 a 1065)

### Resumen de Cambios y Fases Implementadas

#### 1. Fase 1052 — Persistencia de Decisiones de Selección en `ag_aula` y `ag_aula_asignatura` (I1)
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** Al reconfigurar el árbol y volver a presionar "Generar aulas", las aulas existentes se omitían (`omitidas++`), dejando valores obsoletos de `jornada`, `horas_dia` y `dias_ventana`. Además, `INSERT OR IGNORE` no actualizaba `dias_patron` ni `horas_objetivo`.
- **Acción Aplicada:** En `AgFormGeneradorHorarios.cs`, se añadió `UPDATE ag_aula SET jornada = @Jornada, horas_dia = @HorasDia, dias_ventana = @DiasVentana WHERE aula_ref = @AulaRef;` y se reemplazó `INSERT OR IGNORE` por `ON CONFLICT(id_aula, id_asignatura) DO UPDATE SET dias_patron = excluded.dias_patron, horas_objetivo = excluded.horas_objetivo`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`.

#### 2. Fase 1053 — Mapeador y Validador de Selección `AgSeleccionMapper`
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** La lectura de selecciones del árbol no validaba exhaustivamente las 3 dimensiones de decisión (`Jornada`, `HorasDia`, `DiasVentana`), permitiendo combinaciones con horas inválidas o ventanas no reconocidas.
- **Acción Aplicada:** Creado `AgSeleccionMapper.cs` validando `Jornada` (no vacía), `HorasDia` (estrictamente 4 u 8) y `DiasVentana` (al menos un día ISO válido), conservando la etiqueta canónica.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Mappers/AgSeleccionMapper.cs`, `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`.

#### 3. Fase 1054 — Modelo Matemático y de Dominio `AgCarriles`
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** No existía una abstracción desacoplada para la partición disjunta de la ventana semanal en pares de 2 días ($\max(1, |\beta|/2)$), rotación modular de asignaturas y aislamiento de días residuales.
- **Acción Aplicada:** Creado `AgCarriles.cs` con métodos `Ventana`, `Cantidad`, `Indice`, `DelCarril` y `Residual`. Clase interna para prevenir exportación no deseada de UDFs en Excel-DNA.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgCarriles.cs`.

#### 4. Fase 1055 — Modelo Inmutable de Asignación `AgPlanOleada`
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** Se carecía de un modelo que encapsulara el carril, oleada, sub-ventana `[FechaInicio, FechaFin]` y capacidad de ranuras por asignatura dentro de una cohorte.
- **Acción Aplicada:** Creado `AgPlanOleada.cs` como POCO sellado e inmutable con constructor explícito.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgPlanOleada.cs`.

#### 5. Fase 1056 — Blindaje de Sesiones a 4.0h en Pregrado y Regla R11
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** En pregrado aparecían sesiones de 8h por degradación de fallbacks cuando no se encontraba un bloque de 4h para ese día.
- **Acción Aplicada:** En `AgExpansorSesiones.cs` se agregó defensa en profundidad comparando el bloque con `AgUnidadesSemanales.HorasSesionAula`. En `AgValidadorReglas.cs` se implementó la regla dura `R11_DuracionSesionInvalida`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgExpansorSesiones.cs`, `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorReglas.cs`.

#### 6. Fase 1057 — Planificador Determinista de Oleadas `AgPlanificadorOleadas`
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** Las asignaturas de una cohorte se programaban con fechas de ventana completa simultáneamente, produciendo colisiones masivas entre materias de la misma cohorte.
- **Acción Aplicada:** Creado `AgPlanificadorOleadas.cs` que segmenta las semanas útiles en oleadas secuenciales por carril y valida capacidad de ranuras (Invariante I5).
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorOleadas.cs`.

#### 7. Fase 1058 — Integración Plan-First en Fábrica y Erradicación de Degradación a 1 Día
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** Si `AgPlanificadorPatronSemanal` fallaba, reducía silenciosamente la materia a 1 día (`diasK[0]`), multiplicando sesiones en ese día.
- **Acción Aplicada:** En `AgFabricaIndividuos.cs`, ordenamiento canónico `(OrdenCurricular, Id)`, pre-planificación de oleadas y sub-ventanas, supresión del fallback a 1 día (marcado vacante `SIN_DISPONIBILIDAD`), y filtrado estricto de bloques por duración exacta.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`.

#### 8. Fase 1059 — Adaptación de Reglas R5 y R6 para Soporte de Carriles Concurrentes
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** R5 bloqueaba más de 2 materias distintas en el semestre por cohorte matando mallas de 4 asignaturas. R6 ordenaba todas las materias de la cohorte sin distinguir carril, penalizando falsamente materias paralelas en carriles disjuntos.
- **Acción Aplicada:** R5 ahora valida concurrencia por semana ISO acotada a la cantidad de carriles concurrentes del aula. R6 agrupa por carril (`DiasPatron`), permitiendo materias concurrentes en distintos carriles.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorReglas.cs`.

#### 9. Fase 1060 — Compuertas de Seguridad en Exportación y Promoción (I8)
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** La exportación a Excel / CSV permitía descargar mallas con restricciones duras (`HardViolations > 0`) sin advertencia previa.
- **Acción Aplicada:** Añadida compuerta de seguridad interactiva con confirmación en `ExportarResultadosExcel` y `ExportarProyeccionEspejo` en `AgFormGeneradorHorarios.cs`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`.

#### 10. Fase 1061 — Migración V74 Sembrando Bloques de 4.0h Faltantes
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** Ciertos días de la semana carecían de bloques de 4.0h en la tabla `ag_bloque_horario`, provocando que la fábrica o la mutación buscaran cualquier bloque del día (incluso de 8h).
- **Acción Aplicada:** Creada `MigrationV74_SembrarBloques4hFaltantes.cs` asegurando bloques activos de 4.0h para días 1 a 6 e insertada en `SqliteHelper.cs`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV74_SembrarBloques4hFaltantes.cs`, `ConversorPDF.XLL/Core/Helpers/SqliteHelper.cs`, `ConversorPDF.XLL/ConversorPDF.XLL.csproj`.

#### 11. Fase 1062 — Mutación Constructiva Preservando Nivel y Carril
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** `AgOperadorMutacion` seleccionaba bloques aleatorios sin verificar la duración del nivel o si el nuevo día pertenecía al carril de la asignatura.
- **Acción Aplicada:** Restringida la selección de bloques en `AgOperadorMutacion.cs` para coincidir estrictamente con la duración de sesión requerida y mantenerse dentro de los días del carril.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgOperadorMutacion.cs`.

#### 12. Fase 1063 — Suite Exhaustiva de Pruebas Unitarias (T1 a T14)
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** Se requería cobertura completa para los nuevos modelos, mappers, planificadores y migraciones.
- **Acción Aplicada:** Creados `AgCarrilesTests.cs`, `AgPlanificadorOleadasTests.cs`, `AgSeleccionMapperTests.cs` y `MigrationV74Tests.cs`.
- **Archivos Afectados:** `ConversorPDF.Tests/Academic/Domain/AgCarrilesTests.cs`, `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgPlanificadorOleadasTests.cs`, `ConversorPDF.Tests/Academic/Application/Mappers/AgSeleccionMapperTests.cs`, `ConversorPDF.Tests/Academic/Infrastructure/Persistence/MigrationV74Tests.cs`.

#### 13. Fase 1064 — Artículo Técnico de Cambios
- **Estado:** ACEPTADA
- **Acción Aplicada:** Generado `docs/articles/changes/2026-10-22-ag-carriles-oleadas.md`.
- **Archivos Afectados:** `docs/articles/changes/2026-10-22-ag-carriles-oleadas.md`.

#### 14. Fase 1065 — Sincronización de Artefactos de Auditoría y Verificación de Principios
- **Estado:** ACEPTADA
- **Acción Aplicada:** Actualizados `implementation_plan.md` y `walkthrough.md`.
- **Archivos Afectados:** `walkthrough.md`, `implementation_plan.md`.

---

### Verificación de Principios (Ronda 206)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1052 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.cs:4131-4140, 4173-4185` (`UPDATE ag_aula`, `ON CONFLICT DO UPDATE`) |
| F1053 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgSeleccionMapper.cs:13-58`, `AgFormGeneradorHorarios.cs:3383` (`DesdeNodo`) |
| F1054 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgCarriles.cs:12-65` (`DelCarril`, `Cantidad`, `Residual`) |
| F1055 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanOleada.cs:10-46` (POCO inmutable) |
| F1056 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgExpansorSesiones.cs:94-106`, `AgValidadorReglas.cs:420-435` (`R11`) |
| F1057 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorOleadas.cs:14-142` (`Construir`) |
| F1058 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFabricaIndividuos.cs:110-380` (Sin degradación silenciosa, filtro 4h) |
| F1059 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgValidadorReglas.cs:220-330` (R5 por semana ISO, R6 por carril) |
| F1060 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.cs:1983-1996, 2035-2048` (Compuerta I8) |
| F1061 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `MigrationV74_SembrarBloques4hFaltantes.cs:14-65`, `SqliteHelper.cs:375` |
| F1062 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgOperadorMutacion.cs:115-135` (Preservar duración y carril) |
| F1063 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgCarrilesTests.cs`, `AgPlanificadorOleadasTests.cs`, `AgSeleccionMapperTests.cs`, `MigrationV74Tests.cs` |
| F1064 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `docs/articles/changes/2026-10-22-ag-carriles-oleadas.md` |
| F1065 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `walkthrough.md`, `implementation_plan.md` |

---

## 🔧 Ronda 207 — Aislamiento del Log de Pruebas y Descarte del Falso Positivo V66 (Fases 1066 a 1070)

### Diagnóstico y Contexto
- **Falso Positivo Descartado:** El error `[V66] Bloque vinculaciones falló.` con inner exception `[TEST-V66] fallo intencional` en `log_20261022.txt` fue rastreado a la ejecución del test unitario `MigrationV66Tests.Up_WhenBlockFails_ShouldRollbackAndThrow()`. En BD real la tabla `vinculaciones` no tiene restricción UNIQUE/CHECK sobre `modalidadVinculacion` y V66 corrió exitosamente en producción 4 veces ese mismo día.
- **Causa Raíz:** `LoggerHelper._logDirectory` era `readonly` e invariable, fijado en AppData de producción. La suite de pruebas compartía la misma instancia estática, contaminando el archivo de auditoría productiva.
- **Solución Arquitectural:** Redirección dinámica de `LoggerHelper.DirectorioLog` y bootstrap con `[ModuleInitializer]` en el ensamblado de pruebas hacia `%TEMP%\ConversorPDF.Tests\log`.

### Fases Ejecutadas

#### 1. Fase 1066 — Redirección Configurable de la Salida de Log
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** `_logDirectory` era `private static readonly string`, imposibilitando desviar la ruta de registro sin tocar el constructor estático.
- **Acción Aplicada:** Modificado `_logDirectory` a mutable y expuesta la propiedad `internal static string DirectorioLog { get; set; }` con creación automática de directorio.
- **Archivos Afectados:** `ConversorPDF.XLL/Core/Helpers/LoggerHelper.cs`.

#### 2. Fase 1067 — Shim de `ModuleInitializerAttribute` para net48
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** El proyecto de pruebas compila con `net48`, donde `ModuleInitializerAttribute` no existe en BCL, aunque C# 9+ lo soporta si está declarado en el namespace `System.Runtime.CompilerServices`.
- **Acción Aplicada:** Creado `ModuleInitializerAttribute.cs` protegido por `#if !NET5_0_OR_GREATER`.
- **Archivos Afectados:** `ConversorPDF.Tests/ModuleInitializerAttribute.cs`.

#### 3. Fase 1068 — Bootstrap de Pruebas con Redirección a `%TEMP%`
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** Las pruebas unitarias escribían por defecto en `%LOCALAPPDATA%\ConversorPDF\log\` al carecer de inicializador global de ensamblado.
- **Acción Aplicada:** Creado `TestLogBootstrap.cs` con método `[ModuleInitializer]` que asigna `LoggerHelper.DirectorioLog = Path.Combine(Path.GetTempPath(), "ConversorPDF.Tests", "log")`.
- **Archivos Afectados:** `ConversorPDF.Tests/TestLogBootstrap.cs`.

#### 4. Fase 1069 — Prueba de No Regresión de la Redirección
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** Se requería verificar que `LoggerHelper` efectivamente redirige la salida al directorio configurado y restaura su estado limpiamente.
- **Acción Aplicada:** Creado `LoggerHelperRedirectionTests.cs` en `ConversorPDF.Tests/Unit/Core` validando emisión en carpeta temporal y restauración.
- **Archivos Afectados:** `ConversorPDF.Tests/Unit/Core/LoggerHelperRedirectionTests.cs`.

#### 5. Fase 1070 — Trazabilidad Documental y Clarificación en Pruebas
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** Se debía registrar que el error V66 es intencional y formalizar el descarte del falso positivo en la documentación de cambios.
- **Acción Aplicada:** Comentario explicativo en `MigrationV66Tests.cs`, artículo de cambios `docs/articles/changes/2026-10-22-aislamiento-log-pruebas.md` y actualización de `implementation_plan.md` y `walkthrough.md`.
- **Archivos Afectados:** `ConversorPDF.Tests/Academic/Infrastructure/Persistence/MigrationV66Tests.cs`, `docs/articles/changes/2026-10-22-aislamiento-log-pruebas.md`, `implementation_plan.md`, `walkthrough.md`.

---

### Verificación de Principios (Ronda 207)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1066 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `LoggerHelper.cs:24, 34-53` (`DirectorioLog`, `_logDirectory`) |
| F1067 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `ModuleInitializerAttribute.cs:1-17` (Shim net48) |
| F1068 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `TestLogBootstrap.cs:1-25` (`[ModuleInitializer]`) |
| F1069 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `LoggerHelperRedirectionTests.cs:1-74` (xUnit + FluentAssertions) |
| F1070 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `MigrationV66Tests.cs:366-369`, `2026-10-22-aislamiento-log-pruebas.md` |

---

## 🔧 Ronda 208 — Carriles y Oleadas: Fugas de 8 h, Oleada Sobrescrita por Cruce y Carril Obsoleto (Fases 1071 a 1082)

### Diagnóstico y Contexto
- **Fugas de 8 h en Pregrado:** Ocurrían cuando un día carecía de bloque de 4 h y la fábrica recurría a `fallbackBloqueId` (un bloque activo cualquiera, frecuentemente de 8 h o de otro día). Además, V74 sembraba franja TARDE de 4 h de 14:00 a 18:00 mientras V62 y `AgFranjaJornada` definían 13:00 a 17:00, generando fallos de resolución exacta.
- **Oleada Sobrescrita en Cruce:** El descendiente mantenía su carril pero heredaba incondicionalmente las fechas del progenitor (`FechaInicioEstimada`, `FechaFinEstimada`), destruyendo su plan de oleadas y generando falsos solapes R13.
- **Carril Desalineado ante Cambio de Ventana:** `AgFormGeneradorHorarios` reusaba `dias_patron` preexistente; al cambiar la ventana (p. ej. LUN-JUE a MAR-VIE) el patrón no rotaba.
- **Falso Positivo Inter-Aula en Cohortes:** R6 y R13 agrupaban únicamente por `DiasPatron` dentro de la cohorte, provocando colisiones entre aulas distintas (p. ej. Presencial y Plataforma) que compartían el mismo carril.
- **Plan Infactible Degradado Silenciosamente:** Cuando una asignatura no cabía en su oleada (I5), toda la cohorte se marcaba vacante y el AG la aceptaba con penalización blanda en vez de rechazarla como inviable.
- **Ranuras Sobreestimadas:** `AgPlanificadorOleadas` no descontaba días festivos en el cálculo preliminar de capacidad.

### Fases Ejecutadas

#### 1. Fase 1071 — Cuarentena Estricta en Fábrica y Erradicación de `fallbackBloqueId`
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** Al no encontrar bloque de la duración esperada para el nivel en el día asignado, se asignaba `fallbackBloqueId` (bloque arbitrario activo), introduciendo bloques de 8 h o días incorrectos.
- **Acción Aplicada:** Reemplazado fallback por cuarentena estricta: `IdBloqueHorario = 0`, `DiaSemana = diaIso`, `HorasSemanales = horasEfectivas`, `EsVacante = 1`, `Estado = AgEstadoGen.Incompleto` y advertencia limitada. *(Nota de trazabilidad Ronda 209: F1071 cubrió la rama multi-día L525-537; la rama de día único L414-427 y la eliminación de la variable de L65 se completaron en F1083).*
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs:522-536`.

#### 2. Fase 1072 — Protección de Sub-Ventana de Oleada en Operador de Cruce
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** `dst.FechaInicioEstimada = src.FechaInicioEstimada` sobreescribía la sub-ventana de oleada calculada para el descendiente, mezclando oleadas y causando solapes en el carril.
- **Acción Aplicada:** Condicionado a herencia sólo si el descendiente carece de fechas: `if (!dst.FechaInicioEstimada.HasValue) dst.FechaInicioEstimada = src.FechaInicioEstimada`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgOperadorCruce.cs:143-146`.

#### 3. Fase 1073 — Derivación Modular y Resiembra del Carril sobre Ventana Vigente
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** `dias_patron` se leía de la BD sin recalcularlo al cambiar la ventana en la selección del árbol; el UPSERT nunca actualizaba si el valor anterior era no vacío.
- **Acción Aplicada:** Derivación forzosa siempre de `diasPatron` con `string.Join(",", AgCarriles.DelCarril(ventanaIso, idxAulaAsig))` y registro de traza si difiere del valor previo.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs:4195-4220`.

#### 4. Fase 1074 — Validación de Carril Disjunto Consecutivo en `DiasKDe`
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** Si el patrón guardado contenía días dentro de la ventana pero sin formar un carril canónico (p. ej. `1,3`), se retornaba directamente rompiendo el contrato I2.
- **Acción Aplicada:** Validado `esCarril` (2 días, gap 1 entre ellos); en caso de discrepancia, se normaliza al carril modular `AgCarriles.DelCarril(ventana, carril)`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorPatronSemanal.cs:52-64`.

#### 5. Fase 1075 — Discriminación por `IdAula` en R6 y R13
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** Agrupar sólo por `DiasPatron` dentro de la cohorte trataba aulas distintas (Presencial y Plataforma) como un único flujo, disparando violaciones duras de solape de oleada.
- **Acción Aplicada:** Agrupación corregida a `GroupBy(x => $"{x.IdAula}|{(string.IsNullOrWhiteSpace(x.DiasPatron) ? "DEF" : x.DiasPatron)}")` en R6 y R13.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorReglas.cs:374, 442`.

#### 6. Fase 1076 — Validación de Duración Real del Bloque en Regla R11
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** R11 validaba únicamente `gen.HorasSemanales`; si un bloque de 8 h se asignaba con `HorasSemanales = 4`, la regla no detectaba la discrepancia.
- **Acción Aplicada:** Resolución de `duracionBloque` consultando el bloque real en `contexto.BloquesPorId[gen.IdBloqueHorario].DuracionHoras`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorReglas.cs:425-433`.

#### 7. Fase 1077 — Aislamiento de Asignaturas Infactibles y Sanción Dura R7
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** La infactibilidad de oleada degradaba a toda el aula a vacante con penalización suave, permitiendo que el AG convergiera sobre individuos sin plan.
- **Acción Aplicada:** Aislamiento de la asignatura infactible mediante `asigsInfactibles` en fábrica y emisión de violación dura `AgReglaViolacion.R7_CoberturaHorariaIncompleta` por gen sin plan factible.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs:230-244, 271-280`, `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorReglas.cs:471-480`.

#### 8. Fase 1078 — Descuento de Festivos en Cálculo de Capacidad de Oleadas
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** `semOla * diasPatron.Count` sobreestimaba ranuras al ignorar días festivos, difiriendo el fallo I5 hasta fases tardías.
- **Acción Aplicada:** Conteo de `diasEfectivos` iterando de `iniOla` a `finOla` descontando `esFestivo(d)`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorOleadas.cs:120-130`.

#### 9. Fase 1079 — Franja TARDE 4h Canónica 13:00-17:00 en V74
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** V74 sembraba TARDE 4 h con 14:00-18:00, generando inconsistencia con V62 y `AgFranjaJornada.cs:37` (13:00-17:00).
- **Acción Aplicada:** Alineado horario a 13:00-17:00 para días 1 a 5 y día 6.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV74_SembrarBloques4hFaltantes.cs:33-37`.

#### 10. Fase 1080 — Eliminación de `<InternalsVisibleTo Include>` Inerte en csproj Legacy
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** Elemento inerte en csproj no-SDK que arriesgaba conflicto CS0579 al existir ya en `AssemblyInfo.cs:17`.
- **Acción Aplicada:** Eliminado del archivo `.csproj`.
- **Archivos Afectados:** `ConversorPDF.XLL/ConversorPDF.XLL.csproj:685-687`.

#### 11. Fase 1081 — Detección Bloqueante de Combinaciones Inválidas/Incompletas en UI
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** Combinaciones sin jornada, ventana o duración eran descartadas silenciosamente en log, reduciendo aulas generadas sin feedback al operador.
- **Acción Aplicada:** Acumulación en parámetro `descartes` en `ObtenerCombinacionesSeleccionadas` y advertencia bloqueante con `MessageBox` en `GenerarEncabezadosAulasAsync`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs:3390, 4022`.

#### 12. Fase 1082 — Trazabilidad Documental y Batería de Pruebas de No Regresión
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** Sincronización de trazabilidad documental para Rondas 206, 207 y 208, y cobertura de pruebas unitarias.
- **Acción Aplicada:** Creación de pruebas unitarias de F1071 a F1079 en `AgFabricaIndividuosTests.cs`, `AgOperadorCruceTests.cs`, `AgPlanificadorPatronSemanalTests.cs`, `AgValidadorReglasTests.cs`, `AgPlanificadorOleadasTests.cs`, `MigrationV74Tests.cs` y actualización de artefactos.
- **Archivos Afectados:** Archivos de prueba en `ConversorPDF.Tests/` e `implementation_plan.md`, `walkthrough.md`.

---

### Verificación de Principios (Ronda 208)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1071 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFabricaIndividuos.cs:522-536` (`IdBloqueHorario = 0`, `Incompleto`, `I6`) |
| F1072 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgOperadorCruce.cs:143-146` (`if (!dst.FechaInicioEstimada.HasValue)`) |
| F1073 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.cs:4195-4220` (`AgCarriles.DelCarril`, `diasPatron`) |
| F1074 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorPatronSemanal.cs:52-64` (`esCarril`, 2 días, gap 1) |
| F1075 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgValidadorReglas.cs:374, 442` (`${x.IdAula}\|${DiasPatron}`) |
| F1076 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgValidadorReglas.cs:425-433` (`duracionBloque` de `BloquesPorId`) |
| F1077 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFabricaIndividuos.cs:230-244`, `AgValidadorReglas.cs:471-480` (`R7_CoberturaHorariaIncompleta`) |
| F1078 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorOleadas.cs:120-130` (`diasEfectivos` sin festivos) |
| F1079 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `MigrationV74_SembrarBloques4hFaltantes.cs:33-37` (`13:00`, `17:00`) |
| F1080 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `ConversorPDF.XLL.csproj:685` (eliminado `InternalsVisibleTo`) |
| F1081 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.cs:3390, 4022` (`descartes`, `MessageBox`) |
| F1082 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `implementation_plan.md`, `walkthrough.md`, 6 suites de pruebas actualizadas |

---

## 🔧 Ronda 209 — F1071 y F1074 Incompletas: 8 h y Carril Inválido Persisten por Rutas Alternas (Fases 1083 a 1089)

### Diagnóstico y Contexto
- **F1071 Incompleta (8 h por Día Único):** La erradicación de `fallbackBloqueId` en Ronda 208 sólo se aplicó a la rama multi-día (`bloqueDia == null`). La rama de día único (`horasEfectivas == asig.Horas`) aún conservaba `gen.IdBloqueHorario = fallbackBloqueId`, asignando bloques de 8 h o de otros días cuando no existía bloque de 4 h en el catálogo.
- **F1074 Incompleta (Carril Inválido en Cruce):** La validación de carril consecutivo disjunto en `AgPlanificadorPatronSemanal.DiasKDe` sólo se implementó en la sobrecarga con `AgAulaAsignatura`. La sobrecarga textual `DiasKDe(string diasPatron, ...)` no validaba carril y devolvía patrones arbitrarios (`1,3`, `2,4`, etc.). Como `AgOperadorCruce.cs:166` invoca dicha sobrecarga textual, los descendientes heredaban patrones que quebrantaban los invariantes I2/I3.
- **Ambigüedad de Motivo en R7:** El código `SIN_DISPONIBILIDAD` se usaba tanto para infactibilidad de oleadas (I5) como para la falta de bloque de 4 h en BD. La regla dura `R7_CoberturaHorariaIncompleta` sancionaba ambas indistintamente, bloqueando corridas con mensajes erróneos sobre planes infactibles.
- **Fallback de Cruce sin Filtro de Duración:** En `AgOperadorCruce.cs:194-205`, el fallback final seleccionaba el primer bloque activo del día sin filtrar por `unidadCruce` (4 h), reintroduciendo bloques de 8 h en pregrado.
- **Concurrencia R5 Mal Calibrada en Multi-Aula:** La regla R5 determinaba `maxCarriles` examinando únicamente la primera aula del grupo; en cohortes con aulas presenciales (2 carriles) y virtuales (1 carril), se generaban falsos positivos de sobre-concurrencia.
- **Duplicación de Lógica en Oleadas:** `AgPlanificadorOleadas.cs` realizaba un cálculo propio de ranuras y festivos en lugar de utilizar `AgPlanificadorPatronSemanal.Construir`, abriendo brechas de consistencia funcional.

### Fases Ejecutadas

#### 1. Fase 1083 — Cuarentena Estricta en Rama de Día Único y Erradicación Total de `fallbackBloqueId`
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** `fallbackBloqueId` asignaba bloques ajenos en la rama de día único (`horasEfectivas == asig.Horas`) al faltar bloque de la duración requerida.
- **Acción Aplicada:** Reemplazado por cuarentena estricta: `IdBloqueHorario = 0`, `DiaSemana = diaIsoSingle`, `HorasSemanales = horasEfectivas`, `EsVacante = 1`, `Estado = AgEstadoGen.Incompleto`, advertencia con `LoggerHelper.WarningLimitado` y eliminación total del campo `fallbackBloqueId`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs:65, 414-427`.

#### 2. Fase 1084 — Validación de Carril Disjunto en Sobrecarga Textual de `DiasKDe`
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** La sobrecarga textual `DiasKDe(string, string, AgAula, int)` no validaba consecutividad de 2 días (`gap == 1`), permitiendo que el operador de cruce generara patrones disonantes.
- **Acción Aplicada:** Agregada validación `esCarril = dentro.Count == 2 && dentro.Zip(dentro.Skip(1), (a, b) => b - a).All(gap => gap == 1)`; ante fallo, emite advertencia limitada y delega al carril canónico de la ventana.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorPatronSemanal.cs:111-122`.

#### 3. Fase 1085 — Desacoplamiento de Motivo Vacante `SIN_PLAN_OLEADA` y Corrección de R7
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** Confusión entre catálogo incompleto (`SIN_DISPONIBILIDAD`) e infactibilidad matemática de oleada (I5), disparando infracciones duras de R7 equívocas.
- **Acción Aplicada:** Creado `AgMotivoVacante.SIN_PLAN_OLEADA = 6`; asignado en `AgFabricaIndividuos.cs` ante fallo de plan de oleadas; regla R7 en `AgValidadorReglas.cs` evalúa exclusivamente `SIN_PLAN_OLEADA`; UI en `AgFormGeneradorHorarios.cs` configurada para ofrecer ajuste de ventana.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Domain/Enums/AgMotivoVacante.cs:14`, `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs:288-298`, `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorReglas.cs:482-490`, `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs:2285, 2327`.

#### 4. Fase 1086 — Filtro de Duración en Fallback de Bloque en Cruce
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** Último recurso de asignación de bloques en el cruce seleccionaba cualquier bloque activo del día sin filtrar por `unidadCruce`, asignando 8 h en pregrado.
- **Acción Aplicada:** Filtro condicional `Math.Abs(b.DuracionHoras - unidadCruce) < 0.001`; si no se halla bloque coincidente, `dst.IdBloqueHorario = 0`, `dst.HorasSemanales = unidadCruce`, `dst.Estado = AgEstadoGen.Incompleto`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgOperadorCruce.cs:193-211`.

#### 5. Fase 1087 — Calibración de `maxCarriles` en Regla R5 para Cohortes Multi-Aula
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** R5 tomaba el aula del primer gen para calcular la capacidad de carriles de la cohorte, provocando falsos positivos si la primera aula tenía ventana corta (1 carril) y las demás 2 carriles.
- **Acción Aplicada:** Calculado `maxCarriles` como el máximo entre todas las aulas activas de la cohorte con `.Select(a => AgCarriles.Cantidad(AgCarriles.Ventana(a.DiasVentana))).DefaultIfEmpty(2).Max()`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorReglas.cs:325-331`.

#### 6. Fase 1088 — Delegación de Capacidad de Oleada a `Construir` como Fuente Única de Verdad
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** Conteo paralelo de feriados y ranuras en `AgPlanificadorOleadas.cs` podía discrepar del algoritmo de reparto y fechas reales de `AgPlanificadorPatronSemanal.Construir`.
- **Acción Aplicada:** Delegada la validación a `new AgPlanificadorPatronSemanal().Construir(...)`; ante error retorna inmediatamente `I5_INFACTIBLE`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorOleadas.cs:118-130`.

#### 7. Fase 1089 — Trazabilidad Documental y Cobertura de Pruebas Unitarias
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** Necesidad de sincronizar artefactos de auditoría y asegurar cobertura automatizada para todas las correcciones de la ronda.
- **Acción Aplicada:** Pruebas unitarias añadidas en `AgFabricaIndividuosTests.cs`, `AgPlanificadorPatronSemanalTests.cs`, `AgValidadorReglasTests.cs`, `AgOperadorCruceTests.cs`, `AgPlanificadorOleadasTests.cs`, `AgMotivoVacanteTests.cs` y documentación sincronizada en `implementation_plan.md` y `walkthrough.md`.
- **Archivos Afectados:** Suites de prueba en `ConversorPDF.Tests/` y artefactos Markdown.

---

### Verificación de Principios (Ronda 209)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1083 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFabricaIndividuos.cs:414-427` (cuarentena día único, erradicado `fallbackBloqueId`) |
| F1084 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorPatronSemanal.cs:111-122` (`esCarril` en sobrecarga textual) |
| F1085 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgMotivoVacante.cs:14`, `AgValidadorReglas.cs:482`, `AgFabricaIndividuos.cs:292` (`SIN_PLAN_OLEADA`) |
| F1086 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgOperadorCruce.cs:193-211` (`Math.Abs(b.DuracionHoras - unidadCruce) < 0.001`) |
| F1087 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgValidadorReglas.cs:325-331` (`maxCarriles` con `.Max()`) |
| F1088 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorOleadas.cs:118-130` (delegado en `Construir`) |
| F1089 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `implementation_plan.md`, `walkthrough.md`, 6 suites de pruebas actualizadas |

---

## 🔧 Ronda 210 — Plan de Oleadas Recalculado por Individuo, Aborto Total del Plan y `SIN_PLAN_OLEADA` no Inmune a `Sellar` (Fases 1090 a 1096)

### Diagnóstico y Contexto
- **Recálculo O(individuos × asignaturas):** Con F1088, cada invocación de `CrearAleatorio` ejecutaba `AgPlanificadorOleadas.Construir` e instanciaba un `new AgPlanificadorPatronSemanal()` en cada iteración. El plan es determinista por parámetros de aula y ventana semestral, por lo que recalcularlo causaba una severa regresión de rendimiento.
- **`SIN_PLAN_OLEADA` Vulnerable a `Sellar`:** Al cruzar dos individuos o mutar, si un grupo cromosómico quedaba con algún gen con docente real y otro con `SIN_PLAN_OLEADA`, `AgCoherenciaGrupo.Sellar` sobreescribía la vacante asignando docente a todos, suprimiendo la violación dura R7 y falseando el estado en UI.
- **Aborto Total del Plan por una Asignatura:** Si una sola materia de la cohorte no cabía en su sub-ventana de oleada (I5), `AgPlanificadorOleadas` retornaba fallo total y la fábrica descartaba `mapaPlanes` para toda la cohorte, provocando que las demás materias perdieran su carril/oleada y cayeran en solapes R13 y colisiones R5.
- **Fallback Silencioso de Aula sin Ventana:** `AgValidadorVentana.ValidarFactibilidad` omitía aulas sin `dias_ventana` y `AgPlanificadorPatronSemanal.DiasKDe` devolvía `{2,3}` silenciosamente, reproduciendo defectos del pasado.
- **Cálculo de Carriles No Canónico:** La advertencia de multicarril en `AgValidadorVentana.cs` calculaba carriles con `Math.Ceiling(count / 2.0)` reportando cantidades disonantes con la regla canónica de división entera.
- **Capacidad Tautológica:** `CapacidadRanuras` en `AgPlanOleada` reportaba la cantidad de sesiones requeridas `S` en vez de la holgura real de ranuras en la oleada (`semanasOleada × diasCarril`).
- **Motivo de Vacante Sobrecargado:** Ausencia de bloque horario de 4h en el catálogo usaba `SIN_DISPONIBILIDAD`, confundiendo diagnósticos de disponibilidad docente con carencia de catálogo.

### Fases Ejecutadas

#### 1. Fase 1090 — Memoización de Planes de Oleadas e Inyección de Planificador
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** Recálculo redundante del plan de oleadas por individuo y generación de instancias de `AgPlanificadorPatronSemanal` por iteración.
- **Acción Aplicada:** Campo `_cachePlanes` en `AgFabricaIndividuos.cs` e inyección de `_planificadorPatron` en el constructor de `AgPlanificadorOleadas.cs`.
- **Archivos Afectados:** `AgFabricaIndividuos.cs:20, 39, 220`, `AgPlanificadorOleadas.cs:17-21`.

#### 2. Fase 1091 — Inmunidad de `SIN_PLAN_OLEADA` en `AgCoherenciaGrupo.Sellar`
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** Normalización de grupo eliminaba `SIN_PLAN_OLEADA` si existía algún gen con docente real.
- **Acción Aplicada:** Guarda prioritaria previa al cálculo de docentes reales que mantiene la vacante y motivo `SIN_PLAN_OLEADA` para todo el grupo.
- **Archivos Afectados:** `AgCoherenciaGrupo.cs:234-245`.

#### 3. Fase 1092 — Planificación Parcial con `AgPlanOleadasResultado`
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** Falla de una sola asignatura abortaba el plan completo del aula.
- **Acción Aplicada:** Creado modelo `AgPlanOleadasResultado.cs`, acumulación de infactibles en `AgPlanificadorOleadas.cs` y preservación de planes factibles en `AgFabricaIndividuos.cs`.
- **Archivos Afectados:** `AgPlanOleadasResultado.cs:1-20`, `AgPlanificadorOleadas.cs:27, 115-152`, `AgFabricaIndividuos.cs:235-245`.

#### 4. Fase 1093 — Detección Bloqueante de Aula sin Ventana y Traza en Fallback
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** Omisión de aulas sin ventana en validación de factibilidad y asignación permisiva de `{2,3}`.
- **Acción Aplicada:** Bloqueo `[I1_VENTANA_VACIA]` en `AgValidadorVentana.cs`, traza explícita en `AgPlanificadorPatronSemanal.cs` y marcado de cohorte en `AgFabricaIndividuos.cs`.
- **Archivos Afectados:** `AgValidadorVentana.cs:209-214`, `AgPlanificadorPatronSemanal.cs:90-94`, `AgFabricaIndividuos.cs:253-260`.

#### 5. Fase 1094 — Corrección de Cálculo de Carriles en Validador de Ventana
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** `Math.Ceiling(count / 2.0)` producía conteo incorrecto de carriles en ventanas con días impares.
- **Acción Aplicada:** Sustituido por `AgCarriles.Cantidad(ventanaIsoCfg)` canónico.
- **Archivos Afectados:** `AgValidadorVentana.cs:178`.

#### 6. Fase 1095 — Capacidad Teórica Real en Modelo de Oleada
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** `CapacidadRanuras` se igualaba a `S`, perdiendo la holgura real de ranuras de la oleada.
- **Acción Aplicada:** Calculado `capacidadTeorica = semanasPorOla[oleada] * diasPatron.Count`.
- **Archivos Afectados:** `AgPlanificadorOleadas.cs:136`.

#### 7. Fase 1096 — Desacoplamiento de Motivo Vacante `SIN_BLOQUE_NIVEL`
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** `SIN_DISPONIBILIDAD` ocultaba la ausencia de bloques horarios de 4h en catálogo.
- **Acción Aplicada:** Añadido `AgMotivoVacante.SIN_BLOQUE_NIVEL = 7`, asignado en fábrica ante falta de bloque y mapeado en UI de vacantes.
- **Archivos Afectados:** `AgMotivoVacante.cs:15`, `AgFabricaIndividuos.cs:437, 546`, `AgFormGeneradorHorarios.cs:2287, 2325`.

---

### Verificación de Principios (Ronda 210)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1090 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFabricaIndividuos.cs:20, 220`, `AgPlanificadorOleadas.cs:17-21` (`_cachePlanes`) |
| F1091 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgCoherenciaGrupo.cs:234-245` (guarda `vacantePlan`) |
| F1092 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanOleadasResultado.cs`, `AgPlanificadorOleadas.cs:126`, `AgFabricaIndividuos.cs:235` |
| F1093 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgValidadorVentana.cs:209-214`, `AgPlanificadorPatronSemanal.cs:90-94`, `AgFabricaIndividuos.cs:253` |
| F1094 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgValidadorVentana.cs:178` (`AgCarriles.Cantidad`) |
| F1095 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorOleadas.cs:136` (`capacidadTeorica`) |
| F1096 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgMotivoVacante.cs:15`, `AgFabricaIndividuos.cs:437, 546`, `AgFormGeneradorHorarios.cs:2287` |

---

## 🔧 Ronda 211 — Errores de Compilación por Tipo No Registrado y `using` Faltante; Cuarentena I1 Efectiva (Fases 1097 a 1102)

### Diagnóstico y Contexto
- **Error Bloqueante CS0246 por Archivo No Registrado en csproj:** `AgPlanOleadasResultado.cs` fue creado en disco en F1092, pero en el formato de proyecto legacy (.NET Framework 4.8 / MSBuild 2003) los archivos no se incluyen automáticamente mediante globbing. La falta de `<Compile Include>` impedía compilar `AgPlanificadorOleadas.cs` y `AgFabricaIndividuos.cs`.
- **Error Bloqueante CS0246 por Directiva using Faltante:** `AgFabricaIndividuos.cs` utilizaba `Result<AgPlanOleadasResultado>` en `_cachePlanes` sin importar `CSharpFunctionalExtensions`.
- **Cuarentena I1 No Efectiva en Materialización:** Cuando un aula carecía de `DiasVentana`, la fábrica marcaba las asignaturas con `SIN_PLAN_OLEADA`, pero los genes se creaban con bloque de 4 h válido y patrón default `{2,3}`. Al llegar a `AgMaterializadorSesiones`, este sólo filtra si `IdBloqueHorario <= 0`, por lo que materializaba sesiones espurias de MAR/MIÉ con el docente centinela.
- **Doble Sanción Dura por Falta de Bloque de Nivel:** Un gen con motivo `SIN_BLOQUE_NIVEL` e `IdBloqueHorario = 0` activaba la regla genérica `R14_BloqueHorarioNoAsignado`, emitiendo un diagnóstico redundante y confuso.
- **Falta de Verificación Real en Prueba de Memoización:** La prueba previa comparaba únicamente que dos individuos tuvieran el mismo patrón semanal, lo cual pasaba incluso si no había caché.
- **Riesgo de Reutilización de Caché entre Corridas Diferentes:** `_cachePlanes` no invalidaba su contenido al cambiar `idCorrida` ni incluía la identidad de corrida/periodo en su clave.

### Fases Ejecutadas

#### 1. Fase 1097 — Registro de `AgPlanOleadasResultado.cs` en el `.csproj`
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** Proyecto legacy no compila por falta del ítem `<Compile Include>` correspondiente al nuevo modelo `AgPlanOleadasResultado`.
- **Acción Aplicada:** Agregado `<Compile Include="Academic\Domain\Models\Genetica\AgPlanOleadasResultado.cs" />` en `ConversorPDF.XLL.csproj`.
- **Archivos Afectados:** `ConversorPDF.XLL/ConversorPDF.XLL.csproj:653`.

#### 2. Fase 1098 — Importación de Namespace `CSharpFunctionalExtensions` en Fábrica
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** Falta de `using CSharpFunctionalExtensions;` impedía resolver el tipo `Result<T>` en el campo `_cachePlanes`.
- **Acción Aplicada:** Incorporada la directiva `using CSharpFunctionalExtensions;` en los encabezados del archivo.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs:8`.

#### 3. Fase 1099 — Efectividad Operativa de Cuarentena I1 en Construcción de Genes
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** Cuarentena I1 sólo modificaba metadatos de vacante pero dejaba `IdBloqueHorario > 0` y días asignados, permitiendo que el materializador persistiera sesiones en base de datos.
- **Acción Aplicada:** Anulados `gen.IdBloqueHorario = 0`, `gen.DiasPatron = string.Empty` y forzado `gen.Estado = AgEstadoGen.Incompleto` y `gen.MotivoVacante = SIN_PLAN_OLEADA` cuando `cohorteSinPlan` es verdadero.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs:312, 455, 595`.

#### 4. Fase 1100 — Exclusión de `SIN_BLOQUE_NIVEL` en Regla R14
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** Doble penalización dura (`R14` genérica y `SIN_BLOQUE_NIVEL`) ante la ausencia de bloque horario de 4h.
- **Acción Aplicada:** Agregada guarda en `AgValidadorReglas.cs` para ignorar `R14` si el motivo de la vacante es `SIN_BLOQUE_NIVEL`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorReglas.cs:172-178`.

#### 5. Fase 1101 — Contador de Construcciones en Planificador de Oleadas y Prueba Estricta
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** Carencia de un punto de observación para validar de forma determinista que la memoización ejecute exactamente una construcción.
- **Acción Aplicada:** Agregada propiedad interna `Construcciones` con incremento en `Construir`, y actualizada la prueba unitaria para verificar `planner.Construcciones.Should().Be(1)`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorOleadas.cs:19, 34`, `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgFabricaIndividuosTests.cs:310-331`.

#### 6. Fase 1102 — Invalidación de Caché por Corrida y Clave Extendida
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** Reutilización de planes entre corridas distintas o periodos distintos ante colisión de claves.
- **Acción Aplicada:** Implementada limpieza `_cachePlanes.Clear()` si `_ultimaCorrida != idCorrida`, y enriquecida la clave con `idCorrida` y `IdPeriodo`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs:23, 53, 230`, `AgFabricaIndividuosTests.cs:380-405`.

---

### Verificación de Principios (Ronda 211)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1097 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `ConversorPDF.XLL.csproj:653` (`<Compile Include="...AgPlanOleadasResultado.cs" />`) |
| F1098 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFabricaIndividuos.cs:8` (`using CSharpFunctionalExtensions;`) |
| F1099 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFabricaIndividuos.cs:312, 455, 595` (`cohorteSinPlan`, `IdBloqueHorario = 0`, `Incompleto`) |
| F1100 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgValidadorReglas.cs:172-178` (guarda `SIN_BLOQUE_NIVEL` en R14) |
| F1101 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorOleadas.cs:19, 34`, `AgFabricaIndividuosTests.cs:330` (`Construcciones == 1`) |
| F1102 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFabricaIndividuos.cs:23, 53, 230` (`_ultimaCorrida`, clave con corrida/periodo) |

---

## 🔧 Ronda 212 — `SIN_BLOQUE_NIVEL` sin Sanción Dura, Preservación en Sellar y Corrección de Factibilidad I5 (Fases 1103 a 1107)

