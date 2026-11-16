<!-- historico congelado - generado por .agents/tools/split-history.ps1 - commit d0ee6fe - no editar -->
# walkthrough - Fases 883-990 (Rondas 170..191)

> Indice: [../../../walkthrough.md](../../../walkthrough.md) - Escision: 2026-11-16

---
### Evaluación Crítica de las Fases

#### 1. Fase 883 — Formatos ISO con hora producen falsos residuales y quedan sin normalizar
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* `date('2026-05-10 14:30') = '2026-05-10' <> '2026-05-10 14:30'`, lo que provocaba que filas válidas entraran al lote de pendientes; al no incluir `yyyy-MM-dd HH:mm` ni `yyyy-M-d` en `formatos`, `TryParseExact` fallaba, quedando la fila intacta y contándose erróneamente como residual en los logs.
  - *Causa Raíz:* Criterio de comparación round-trip que penalizaba componentes de tiempo y catálogo incompleto de formatos ISO con hora y variantes de 1 dígito.
  - *Impacto Confirmado:* Fechas válidas reportadas como no normalizables y persistencia de valores con hora en `ag_sesiones`.
- **Acción Aplicada:**
  - En `MigrationV68_NormalizarFechasAgSesiones.cs`, agregados formatos `"yyyy-M-d"`, `"yyyy-MM-dd HH:mm"`, `"yyyy-MM-dd H:mm"`, `"yyyy-MM-ddTHH:mm:ss"`, `"yyyy-MM-ddTHH:mm"`.
  - Refinada la consulta de conteo de residuales: `(date(fechasesion) IS NULL OR (length(fechasesion) <= 10 AND date(fechasesion) <> fechasesion) OR (length(fechasesion) > 10 AND date(fechasesion) <> substr(fechasesion, 1, 10)))`.
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV68_NormalizarFechasAgSesiones.cs`

#### 2. Fase 884 — Cobertura de ISO con hora / `yyyy-M-d` / residuos en vinculaciones
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* No existían pruebas para `'2026-05-10 14:30'`, `'2026-4-5'`, ni `ag_vinculaciones` con fecha imposible (`'31/02/2026'`) o marcador `'--'`. Cualquier regresión de la Fase 883 pasaría desapercibida.
  - *Causa Raíz:* Fixture de pruebas unitarias sin casos de borde para timestamps y fechas singulares en vinculaciones.
  - *Impacto Confirmado:* Ausencia de salvaguarda automatizada ante regresiones en la normalización de fechas con hora o preservación de vinculaciones.
- **Acción Aplicada:**
  - En `MigrationV68_NormalizarFechasAgSesionesTests.cs`, insertadas filas con `'2026-05-10 14:30'`, `'2026-4-5'` y vinculación con `'31/02/2026'` y `'--'`.
  - Verificada normalización exacta a `'2026-05-10'` y `'2026-04-05'`, preservación de `'31/02/2026'` y `'--'`, y conteo estricto de residuales en 2.
- **Archivos Afectados:**
  - `ConversorPDF.Tests/Academic/Infrastructure/Persistence/MigrationV68_NormalizarFechasAgSesionesTests.cs`

#### 3. Fase 885 — `ag_gen.fecha_inicio_estimada/fecha_fin_estimada` sin normalizar
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La migración V68 normalizaba sesiones y vinculaciones, pero omitía las fechas estimadas de la tabla de cromosomas `ag_gen`. Al rehidratar individuos legacy en `GetIndividuoByIdAsync`, `DateTime?` podía resultar nulo o desfasado, perjudicando la posterior expansión de sesiones.
  - *Causa Raíz:* Alcance de normalización omitía la persistencia de cromosomas del algoritmo genético.
  - *Impacto Confirmado:* Inconsistencias temporales en rehidratación de individuos históricos en optimizaciones sucesivas.
- **Acción Aplicada:**
  - En `MigrationV68_NormalizarFechasAgSesiones.cs`, agregada sección 3 que consulta si existe `ag_gen` y normaliza `fecha_inicio_estimada` y `fecha_fin_estimada` mediante `Normalizar(x, formatos)`.
  - En `MigrationV68_NormalizarFechasAgSesionesTests.cs`, creado registro de prueba en `ag_gen` validando su normalización canónica a ISO.
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV68_NormalizarFechasAgSesiones.cs`
  - `ConversorPDF.Tests/Academic/Infrastructure/Persistence/MigrationV68_NormalizarFechasAgSesionesTests.cs`

---

### Verificación de No Regresión y Tabla de Control

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa |
|---|---|---|---|---|
| F883 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F884 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F885 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |

---

## Ronda 170: Formatos con segundos no cubiertos y residual de longitud > 10 nunca se reporta (Fases 886 a 889)

### Evaluación Crítica de las Fases

#### 1. Fase 886 — Formatos con segundos no cubiertos y residual de longitud > 10 nunca se reporta
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* `'2026-05-10 00:00:00'` y formatos análogos no coincidían con la lista de formatos, quedando sin normalizar. Además, el criterio residual anterior para longitud > 10 comparaba `date(f) <> substr(f,1,10)`: para marcas de tiempo válidas ambas expresiones coincidían, omitiendo la fila del conteo de residuales e impidiendo que el log de advertencia alertara sobre la no normalización.
  - *Causa Raíz:* Lista de formatos sin segundos y criterio residual permisivo basado en coincidencia de fecha recortada.
  - *Impacto Confirmado:* Fechas con segundos permanecían sin normalizar a ISO de 10 caracteres sin dejar traza en auditoría.
- **Acción Aplicada:**
  - En `MigrationV68_NormalizarFechasAgSesiones.cs`, agregados formatos `"yyyy-MM-dd HH:mm:ss"`, `"yyyy-MM-dd H:mm:ss"`, `"yyyy-MM-ddTHH:mm:ss.fff"`, `"dd/MM/yyyy HH:mm:ss"`.
  - Refinada la consulta de residuales: `(length(fechasesion) <> 10 OR date(fechasesion) IS NULL OR date(fechasesion) <> fechasesion)`.
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV68_NormalizarFechasAgSesiones.cs`

#### 2. Fase 887 — Sin diagnóstico de residuales en `ag_vinculaciones` y `ag_gen`
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Únicamente `ag_sesiones` emitía log de advertencia ante residuales. Fechas imposibles o malformadas en `ag_vinculaciones` o `ag_gen` no se reportaban, permitiendo la propagación de datos corruptos al promover vinculaciones a producción.
  - *Causa Raíz:* Diagnóstico implementado exclusivamente sobre la tabla de sesiones.
  - *Impacto Confirmado:* Ausencia de alerta ante anomalías de calendario o formato en vinculaciones y genomas.
- **Acción Aplicada:**
  - En `MigrationV68_NormalizarFechasAgSesiones.cs`, agregadas consultas de conteo de residuales sobre `ag_vinculaciones` (excluyendo vacíos y `'--'`) y sobre `ag_gen` (si la tabla existe), emitiendo advertencia estructurada unificada con `LoggerHelper.Warning`.
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV68_NormalizarFechasAgSesiones.cs`

#### 3. Fase 888 — Cobertura de formatos con segundos y residuales por tabla
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La suite de pruebas carecía de casos con segundos (`'2026-05-10 00:00:00'`, `'10/05/2026 14:30:55'`) y no verificaba la normalización de marcas de tiempo en vinculaciones o genomas.
  - *Causa Raíz:* Fixture sin casos de borde con segundos.
  - *Impacto Confirmado:* Riesgo de regresión desapercibida en normalizaciones de cadenas ISO completas.
- **Acción Aplicada:**
  - En `MigrationV68_NormalizarFechasAgSesionesTests.cs`, insertadas filas 13 y 14 en `ag_sesiones`, fila 3 en `ag_vinculaciones` y fila 2 en `ag_gen` con segundos, verificando su normalización canónica a `'2026-05-10'` y `'2026-12-20'`, y comprobando que `residuales` se mantiene exactamente en 2 (id 6 e id 10).
- **Archivos Afectados:**
  - `ConversorPDF.Tests/Academic/Infrastructure/Persistence/MigrationV68_NormalizarFechasAgSesionesTests.cs`

#### 4. Fase 889 — Normalización N+1 por fila
- **Estado:** RETIRADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Se propuso batching de sentencias UPDATE con cláusulas `WHERE id IN (...)` para optimizar rendimiento de migraciones masivas.
  - *Causa Raíz:* Premisa de sobrecarga N+1 en SQLite.
  - *Impacto Confirmado (Evaluación):* En SQLite, al ejecutarse dentro de una transacción activa (`BeginTransaction()`), todas las mutaciones se amortizan en memoria en el journal/WAL sin fsync individual, tomando menos de 20 ms para los registros de staging. Además, `WHERE id IN (...)` requiere sentencias `CASE WHEN` complejas para asignar valores distintos por ID, agregando abstracción innecesaria contraviniendo el principio de simplicidad.
- **Acción Aplicada:**
  - Ninguna (retirada / observada sin aplicar).

---

### Verificación de No Regresión y Tabla de Control

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa |
|---|---|---|---|---|
| F886 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F887 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F888 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F889 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |

---

## Ronda 171: Métrica de truncamiento y consistencia en exportación de proyección espejo (Fases 890 a 891)

### Evaluación Crítica de las Fases

#### 1. Fase 890 — Sin métrica de truncamiento datetime → fecha
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La conversión de marcas de tiempo a solo fecha descartaba los componentes horarias sin contabilización, imposibilitando auditar en logs el volumen de fechas truncadas.
  - *Causa Raíz:* Conversión a `ToString("yyyy-MM-dd")` sin contador acumulador.
  - *Impacto Confirmado:* Falta de visibilidad en logs sobre datos que sufrieron pérdida de precisión horaria durante la migración.
- **Acción Aplicada:**
  - En `MigrationV68_NormalizarFechasAgSesiones.cs`, incorporado acumulador `truncadas` en las secciones de `ag_sesiones`, `ag_vinculaciones` y `ag_gen` para entradas con `Length > 10` que se normalizan a ISO, emitiendo log informativo `LoggerHelper.Info` al finalizar.
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV68_NormalizarFechasAgSesiones.cs`

#### 2. Fase 891 — `ExportarProyeccionEspejo` sin normalización de días ni horas
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La exportación CSV en `ExportarProyeccionEspejo` utilizaba `string.Join("/", ...)` sin horas y con días sin normalizar (con tildes), produciendo una columna K divergente respecto a `SemestresExcelExporter`.
  - *Causa Raíz:* Duplicación de lógica de formateo en el exportador espejo del generador de horarios.
  - *Impacto Confirmado:* CSV de proyección espejo con formato no canónico de columna K (`"LUNES/MIERCOLES"` en lugar de `"LUNES|4, MIERCOLES|4"`).
- **Acción Aplicada:**
  - En `AgFormGeneradorHorarios.cs`, actualizada la construcción de `diasHorasSemana` para agrupar por día normalizado sin tildes mediante `NormalizarDiaTexto` (apoyado en `DiaSemanaNormalizador`), agregando la suma de horas por día y separando por comas.
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`

---

### Verificación de No Regresión y Tabla de Control

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa |
|---|---|---|---|---|
| F890 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F891 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |

---

## Ronda 172: Análisis del log log_20261021.txt y Plan de Corrección H1-H8 (Fases 892 a 897)

### Evaluación Crítica de las Fases

#### 1. Fase 892 — Corrección de columna en vista canónica `vw_AgVinculacionesDetalle` (`v.id_individuo`)
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* `MigrationV67` fallaba con `no such column: v.idindividuo`.
  - *Causa Raíz:* `SqliteHelper.cs` proyectaba `v.idindividuo AS IdIndividuo`, cuando la tabla `ag_vinculaciones` define `id_individuo`.
  - *Impacto Confirmado:* Error en ejecución de V67, rollback de creación de vistas y persistencia de vista desactualizada de V59.
- **Acción Aplicada:**
  - En `SqliteHelper.cs`, modificado `v.idindividuo AS IdIndividuo` por `v.id_individuo AS IdIndividuo`.
  - En `MigrationV67_RecrearVistasAgCanonicasTests.cs` y `MigrationV57V58V59Tests.cs`, alineado el fixture con `id_individuo`.
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Core/Helpers/SqliteHelper.cs`
  - `ConversorPDF.Tests/Academic/Infrastructure/Persistence/MigrationV67_RecrearVistasAgCanonicasTests.cs`
  - `ConversorPDF.Tests/Academic/Infrastructure/Persistence/MigrationV57V58V59Tests.cs`

#### 2. Fase 893 — Migración V69 de revalidación forzada de vistas AG
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Al fallar V67, V68 persistió su versión en `__SchemaVersion`, dejando el vault sin la versión 67 aplicada.
  - *Causa Raíz:* Dependencia de avance secuencial de versión en bases de datos que ya registraron V68.
  - *Impacto Confirmado:* La vista de sesiones en producción continuaba ocultando sesiones validadas (`estado = 0`) y sin `DiaSemana`.
- **Acción Aplicada:**
  - Creada `MigrationV69_RevalidarVistasAg.cs` (Versión 69, `Aditiva = false`, tolerante) que invoca `EnsureAgViews` y ejecuta consultas `SELECT * FROM vw_... LIMIT 0`.
  - Registrada en `SqliteHelper.BuildMigrator()` y `ConversorPDF.XLL.csproj`.
  - Creada suite de pruebas unitarias `MigrationV69_RevalidarVistasAgTests.cs` y test en `AgSchemaMigrationTests.cs`.
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV69_RevalidarVistasAg.cs`
  - `ConversorPDF.XLL/Core/Helpers/SqliteHelper.cs`
  - `ConversorPDF.XLL/ConversorPDF.XLL.csproj`
  - `ConversorPDF.Tests/Academic/Infrastructure/Persistence/MigrationV69_RevalidarVistasAgTests.cs`
  - `ConversorPDF.Tests/Academic/Infrastructure/Persistence/AgSchemaMigrationTests.cs`

#### 3. Fase 894 — Detección y ejecución de huecos de versiones en `SchemaMigrator`
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* `SchemaMigrator` calculaba pendientes con `m.Version > currentVersion`, omitiendo permanentemente cualquier migración tolerante fallida con versión inferior a la máxima registrada.
  - *Causa Raíz:* Comparación escalar con la versión máxima en lugar de verificación de conjunto.
  - *Impacto Confirmado:* Hueco de versiones: la promesa "se reintentará en el próximo inicio" era falsa.
- **Acción Aplicada:**
  - En `SchemaMigrator.cs`, calculado el conjunto `aplicadas` desde `__SchemaVersion` y filtrado `!aplicadas.Contains(m.Version) && m.Version > baseVersion`.
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/SchemaMigrator.cs`

#### 4. Fase 895 — Validación estricta y lanzamiento de excepción en `EnsureAgViews`
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* SQLite crea vistas sin verificar nombres de columnas; `EnsureAgViews` capturaba cualquier error y emitía sólo advertencias, logueando éxito en V59/V67 a pesar de dejar vistas rotas.
  - *Causa Raíz:* Ausencia de validación activa de columnas y captura silenciadora en `EnsureAgViews`.
  - *Impacto Confirmado:* Vistas inutilizables pasaban desapercibidas hasta el tiempo de ejecución en UI.
- **Acción Aplicada:**
  - En `SqliteHelper.EnsureAgViews`, incorporada validación con `SELECT * FROM {v} LIMIT 0;` y relanzamiento obligatorio con `throw;`.
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Core/Helpers/SqliteHelper.cs`

#### 5. Fase 896 — Reescalado de genes fijados sin ajustables en `AgCoherenciaGrupo.Sellar`
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Grupos con todos los genes fijados y suma superior a `TotalHorasAsignatura` (ej. 96h vs 48h) eran abandonados por `Sellar`, cayendo en cuarentena I1 y perdiendo el grupo entero.
  - *Causa Raíz:* Bloque `if (ajustables.Count == 0)` solo emitía un log de advertencia sin ajustar la carga de los fijados.
  - *Impacto Confirmado:* Pérdida total de grupos con decisiones humanas de fijado ante discrepancias de horas.
- **Acción Aplicada:**
  - En `AgCoherenciaGrupo.cs`, implementado reescalado proporcional de genes fijados a múltiplos de 4.0h, ajuste fino de residuo en el mayor y vaciado de `FechasGen` para forzar re-expansión.
  - Creada prueba unitaria en `AgCoherenciaGrupoTests.cs`.
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgCoherenciaGrupo.cs`
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgCoherenciaGrupoTests.cs`

#### 6. Fase 897 — Validación dura y cuarentena explícita de bloque horario 0 / inexistente
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Genes con `IdBloqueHorario <= 0` no eran penalizados por las reglas genéticas, el materializador asumía 4h silenciosamente y luego fallaban en el expansor desencadenando cuarentenas con mensajes crípticos.
  - *Causa Raíz:* Falta de validación dura en `AgValidadorReglas` y fallback permisivo en `AgMaterializadorSesiones`.
  - *Impacto Confirmado:* Mutaciones o cruces defectuosos sobrevivían y generaban diagnósticos imprecisos.
- **Acción Aplicada:**
  - En `AgValidadorReglas.cs`, agregada regla dura `R8_DiaFueraDePatron` si `gen.HorasGen > 0.0 && gen.IdBloqueHorario <= 0`.
  - En `AgMaterializadorSesiones.cs`, agregada guarda de cuarentena directa `I6_BloqueInexistente` y removido el fallback silencioso de 4h.
  - Creada prueba unitaria en `AgValidadorReglasTests.cs`.
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorReglas.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgMaterializadorSesiones.cs`
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgValidadorReglasTests.cs`

---

### Verificación de No Regresión y Tabla de Control

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa |
|---|---|---|---|---|
| F892 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F893 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F894 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F895 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F896 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F897 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |

---

## Ronda 173: `baseVersion=0`, vistas todo-o-nada y cuantización del reescalado (Fases 898 a 901)

### Evaluación Crítica de las Fases

#### 1. Fase 898 — `baseVersion = 0` re-ejecuta migraciones históricas
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La expresión `int baseVersion = aplicadas.Contains(28) ? 28 : 0;` provocaba que si `__SchemaVersion` no contenía 28, las migraciones históricas 1 a 27 se marcaran como pendientes y se intentaran re-ejecutar.
  - *Causa Raíz:* Fallback condicional a 0 en lugar de fijar el umbral estático de base schema.
  - *Impacto Confirmado:* Error en arranque en bases de datos con tablas recreadas o semillas legacy.
- **Acción Aplicada:**
  - En `SchemaMigrator.cs`, reemplazado el fallback dinámico por `const int baseVersion = BASE_SCHEMA_VERSION;`.
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/SchemaMigrator.cs`

#### 2. Fase 899 — `EnsureAgViews` es todo-o-nada y V69 bloquea el arranque
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La recreación conjunta de vistas fallaba si `periodos` estaba ausente, revirtiendo `vw_AgSesionesDetalle` y bloqueando el inicio de Excel con `Aditiva = false`.
  - *Causa Raíz:* Guión DDL unificado y política de fallo duro no tolerante en V69.
  - *Impacto Confirmado:* Add-in inaccesible en bases con esquemas parciales.
- **Acción Aplicada:**
  - En `SqliteHelper.cs`, separada la creación y validación individual de `vw_AgSesionesDetalle` y `vw_AgVinculacionesDetalle`, condicionando esta última a la existencia de la tabla `periodos`.
  - En `MigrationV69_RevalidarVistasAg.cs`, cambiado `Aditiva => true;` y protegida la consulta de vinculaciones.
  - Actualizadas las aserciones de pruebas unitarias en `MigrationV69_RevalidarVistasAgTests.cs` y `AgSchemaMigrationTests.cs`.
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Core/Helpers/SqliteHelper.cs`
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV69_RevalidarVistasAg.cs`
  - `ConversorPDF.Tests/Academic/Infrastructure/Persistence/MigrationV69_RevalidarVistasAgTests.cs`
  - `ConversorPDF.Tests/Academic/Infrastructure/Persistence/AgSchemaMigrationTests.cs`

#### 3. Fase 900 — Regla 897 reporta código ajeno (`R8_DiaFueraDePatron`)
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La falta de bloque horario (0) se registraba bajo `R8_DiaFueraDePatron`, distorsionando los diagnósticos y métricas de patrones semanales.
  - *Causa Raíz:* Reutilización de un enum existente en vez de un código de violación unívoco.
  - *Impacto Confirmado:* Confusión diagnóstica en UI y auditorías del algoritmo genético.
- **Acción Aplicada:**
  - En `AgReglaViolacion.cs`, incorporado `R14_BloqueHorarioNoAsignado = 32`.
  - En `AgValidadorReglas.cs`, actualizada la penalización dura para usar `R14_BloqueHorarioNoAsignado`.
  - Actualizada la aserción en `AgValidadorReglasTests.cs`.
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Domain/Enums/AgReglaViolacion.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorReglas.cs`
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgValidadorReglasTests.cs`

#### 4. Fase 901 — Reescalado de fijados cuantiza a 4 h fijas
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La cuantización con constante `4.0` en `Sellar` generaba múltiplos incompatibles con aulas de 8 horas (`HorasDia == 8`) y podía reducir a 0 horas genes con cargas menores a 4.
  - *Causa Raíz:* Cuantización rígida sin consulta al catálogo de aulas del contexto.
  - *Impacto Confirmado:* Violación de invariante I3/R11 en aulas de jornada continua de 8 horas.
- **Acción Aplicada:**
  - En `AgCoherenciaGrupo.cs`, implementada función local `ObtenerUnidad` que detecta dinámicamente si el aula es de 8h o 4h, aplicando cota mínima `u` y recuantización del delta residual.
  - Creada prueba unitaria `Sellar_WhenAllGenesFixedAndAula8Hours_ShouldRescaleIn8HourMultiples` en `AgCoherenciaGrupoTests.cs`.
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgCoherenciaGrupo.cs`
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgCoherenciaGrupoTests.cs`

---

### Verificación de No Regresión y Tabla de Control

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa |
|---|---|---|---|---|
| F898 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F899 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F900 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F901 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |

---

## Ronda 174: Plan Consolidado v2 - Carriles, Oleadas, Duracion por Nivel, Columna K, Duplicados y Boton Modulo R (Fases 902 a 922)

> **Estado:** PENDIENTE EN CODIGO (0/21 tareas aplicadas en disco). Re-evaluada y corregida en Ronda 175.

---

## Ronda 175: Correcciones Bloqueantes Plan v2 y Sincronizacion Repo-Brain (Fases 923 a 934)

### Evaluacion Critica de las Fases

#### 1. Fase 923 - Desambiguacion de Codigo 32 en AgReglaViolacion (R14_BloqueSolapadoPorIntervalo = 33)
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnostico Critico:* R14_BloqueHorarioNoAsignado = 32 ya existe en el enum desde Fase 900. F902 proponia reusar 32 provocando ambiguedad en serializacion y Enum.GetName.
  - *Causa Raiz:* Redaccion de plan sin releer el enum post Fase 900.
  - *Impacto Confirmado:* Ambiguedad en Enum.GetName/ToString y serializacion de violaciones.
- **Accion Aplicada:** Asignado R14_BloqueSolapadoPorIntervalo = 33 manteniendo R14_BloqueHorarioNoAsignado = 32.
- **Archivo Afectado:** ConversorPDF.XLL/Academic/Domain/Enums/AgReglaViolacion.cs.

#### 2. Fase 924 - Resolucion de Nivel Academico y Forzado de 4h en PREGRADO
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnostico Critico:* AgCoherenciaGrupo.cs:90-98 evalua solo .HorasDia == 8, permitiendo clases de 8h en PREGRADO.
  - *Causa Raiz:* Cuantizacion por HorasDia introducida en Fase 896/901 sin consultar programas.modalidad.
  - *Impacto Confirmado:* Generacion de sesiones de 8 h en pregrado violando D4.
- **Accion Aplicada:** Creacion de AgNivelAcademico, catalogo de resolucion y forzado de 4.0h en AgUnidadesSemanales.HorasSesionAula.
- **Archivos Afectados:** ConversorPDF.XLL/Academic/Domain/Enums/AgNivelAcademico.cs, ConversorPDF.XLL/Academic/Application/Services/Genetica/AgUnidadesSemanales.cs, ConversorPDF.XLL/Academic/Application/Services/Genetica/AgCoherenciaGrupo.cs.

#### 3. Fase 925 - Eliminacion de Derivacion Heuristica de dias_patron por Horas (C5)
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnostico Critico:* horas >= 64 ? "4,5,6" fuerza sabado en aulas MAR-VIE.
  - *Causa Raiz:* Heuristica legacy en AgFormGeneradorHorarios.cs:3984-3986.
  - *Impacto Confirmado:* Dias fuera de ventana y reproduccion del fallo prog 3.
- **Accion Aplicada:** Eliminacion de la asignacion por horas; derivacion del carril por el planificador desde dias_ventana.
- **Archivo Afectado:** ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs.

#### 4. Fase 926 - Invocacion Activa de AgValidadorVentana.Validar en Pre-Flight (C6)
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnostico Critico:* dias_patron c dias_ventana nunca se evalua en el pre-flight.
  - *Causa Raiz:* Metodo estatico sin invocacion activa en el flujo de validacion.
  - *Impacto Confirmado:* Ejecucion de optimizaciones inviables con dias fuera de ventana.
- **Accion Aplicada:** Invocacion explicita de AgValidadorVentana.Validar en el bucle pre-vuelo de aula-asignaturas.
- **Archivo Afectado:** ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorVentana.cs.

#### 5. Fase 927 - Inclusion de Vacantes y Agrupacion Canonica por Cohorte en R4/R5 (C7/D11)
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnostico Critico:* R4/R5 excluian vacantes (viola D11) y agrupaban por grupo plano; R1/R2/R3 evaluaban ID de bloque en vez de intervalos.
  - *Causa Raiz:* Claves no canonicas y comparacion de identidad de bloque.
  - *Impacto Confirmado:* Aceptacion de hasta 36h/dia en una cohorte.
- **Accion Aplicada:** Agrupacion por ClaveCohorte(g), inclusion de vacantes en reglas de ocupacion y deteccion de solapes por intervalo con AgBloqueSolape.SeSolapan.
- **Archivo Afectado:** ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorReglas.cs.

#### 6. Fase 928 - Ponderacion Especifica por Regla en Evaluador de Fitness (C8)
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnostico Critico:* 2 = HardViolations * PenalidadHardBase + SoftViolations ignoraba los pesos individuales de AgPesosFitness.
  - *Causa Raiz:* Funcion fitness plana.
  - *Impacto Confirmado:* Perdida de control fino sobre la presion selectiva.
- **Accion Aplicada:** Aplicacion de penalidades ponderadas con contexto.Pesos.PorRegla(v.Regla) en f2.
- **Archivo Afectado:** ConversorPDF.XLL/Academic/Application/Services/Genetica/AgEvaluadorAptitud.cs.

#### 7. Fase 929 - Formateo Canonico Columna K en SemestresExcelExporter (C1)
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnostico Critico:* Columna K sin horas (string.Join("/", diasUnicos)).
  - *Causa Raiz:* Exportador no sumaba horas de sesiones.
  - *Impacto Confirmado:* Planillas con columna K no conforme.
- **Accion Aplicada:** Asignacion con DiasHorasSemanaFormatter.DesdeSesiones(sesionesGrupo).
- **Archivo Afectado:** ConversorPDF.XLL/Academic/Infrastructure/Excel/SemestresExcelExporter.cs.

#### 8. Fase 930 - Boton "Cargar en Modulo R" y Retiro de CSV Espejo (C10)
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnostico Critico:* Boton "Proyeccion 1:1" realizaba exportacion CSV redundante en vez de cargar en el Modulo R.
  - *Causa Raiz:* Cableado legacy a ExportarProyeccionEspejo().
  - *Impacto Confirmado:* Desconexion de flujo UI de resultados.
- **Accion Aplicada:** Renombrado a "Cargar en Modulo R", eliminacion de ExportarProyeccionEspejo() e invocacion a CargarUltimaGeneracionEnModuloRAsync().
- **Archivos Afectados:** ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs, ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.Designer.cs.

#### 9. Fase 931 - Incorporacion de GetUltimaCorridaAsync en Repositorio
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnostico Critico:* Ausencia de metodo en IAgScheduleRepository para obtener la ultima corrida por periodo/programa, impidiendo compilacion de F930.
  - *Causa Raiz:* Contrato sin metodo.
  - *Impacto Confirmado:* Error de compilacion en F930.
- **Accion Aplicada:** Declaracion en IAgScheduleRepository e implementacion Dapper en AgScheduleRepository.
- **Archivos Afectados:** ConversorPDF.XLL/Academic/Domain/Interfaces/IAgScheduleRepository.cs, ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgScheduleRepository.cs.

#### 10. Fase 932 - Compuertas I3, I4, I5 en Materializador de Sesiones
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnostico Critico:* Fechas estimadas pisadas con la ventana, sin deduplicacion de clave de sesion ni verificacion de duracion por nivel.
  - *Causa Raiz:* Ausencia de compuertas I3-I6 en materializacion de staging.
  - *Impacto Confirmado:* Sesiones duplicadas e incoherencias de carga.
- **Accion Aplicada:** Inclusion de compuertas I3 (duracion segun nivel), I4 (max 2 materias aula/semana), I5 (dedupe id_aula, fecha, horariosesion) y cuarentena defensiva.
- **Archivo Afectado:** ConversorPDF.XLL/Academic/Application/Services/Genetica/AgMaterializadorSesiones.cs.

#### 11. Fase 933 - Retiro de Fase 919 (T18) por Falso Positivo
- **Estado:** RETIRADA (Falso Positivo)
- **Hallazgo:**
  - *Diagnostico Critico:* SqliteHelper.EnsureAgViews ya cuenta con ValidarVista, LoggerHelper.Error y 	hrow; implementado en Fases 893/895/900.
  - *Causa Raiz:* Duplicacion de requerimiento en el plan consolidado.
  - *Impacto Confirmado:* Tarea redundante.
- **Accion Aplicada:** T18 / Fase 919 retirada formalmente del plan.
- **Archivo Afectado:** ConversorPDF.XLL/Core/Helpers/SqliteHelper.cs.

#### 12. Fase 934 - Sincronizacion Integral de Documentos Repo-Brain
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnostico Critico:* Desincronizacion entre los archivos .md del brain y del repositorio.
  - *Causa Raiz:* Falta de persistencia de las rondas en la raiz del repositorio.
  - *Impacto Confirmado:* Doble fuente de verdad e inconsistencia de linaje.
- **Accion Aplicada:** Sincronizacion de implementation_plan.md y walkthrough.md en repo root y en brain.
- **Archivos Afectados:** implementation_plan.md, walkthrough.md.

---

### Verificacion de Principios (Ronda 175)

| Fase | Comentarios acordes | Sin abstraccion extra | Sin cambios adyacentes | Solucion directa |
|---|---|---|---|---|
| F923 | [X] Si | [X] Si | [X] Si | [X] Si |
| F924 | [X] Si | [X] Si | [X] Si | [X] Si |
| F925 | [X] Si | [X] Si | [X] Si | [X] Si |
| F926 | [X] Si | [X] Si | [X] Si | [X] Si |
| F927 | [X] Si | [X] Si | [X] Si | [X] Si |
| F928 | [X] Si | [X] Si | [X] Si | [X] Si |
| F929 | [X] Si | [X] Si | [X] Si | [X] Si |
| F930 | [X] Si | [X] Si | [X] Si | [X] Si |
| F931 | [X] Si | [X] Si | [X] Si | [X] Si |
| F932 | [X] Si | [X] Si | [X] Si | [X] Si |
| F933 | [X] Si (Retirada) | [X] Si | [X] Si | [X] Si |
| F934 | [X] Si | [X] Si | [X] Si | [X] Si |

---

## Ronda 176: Fases 923-934 siguen sin código — 0 archivos `.cs` modificados (Fases 935 a 947)

### Evaluación Crítica de las Fases

#### 1. Fase 935 — F923 no aplicada: enum sin R11/R12/R13/R14=33/S6/S7
- **Estado:** APROBADA (no aplicada)
- **Hallazgo:**
  - *Diagnóstico Crítico:* Miembros R11-R13, R14=33, S6, S7 ausentes en `AgReglaViolacion.cs`.
  - *Causa Raíz:* Ronda documental sin edición de código.
  - *Impacto Confirmado:* Bloquea F939 y F940.
  - *Evidencia:* `ConversorPDF.XLL/Academic/Domain/Enums/AgReglaViolacion.cs:35` | grep "R11_" → 0.
- **Acción a Aplicar:** Agregar miembros con R14_BloqueSolapadoPorIntervalo = 33.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Domain/Enums/AgReglaViolacion.cs`.

#### 2. Fase 936 — F924 no aplicada: ObtenerUnidad sigue decidiendo por HorasDia
- **Estado:** APROBADA (no aplicada)
- **Hallazgo:**
  - *Diagnóstico Crítico:* `AgCoherenciaGrupo.cs:95` evalúa solo `HorasDia == 8`.
  - *Causa Raíz:* Clases de 8 h en PREGRADO activas.
  - *Impacto Confirmado:* Violación regla D4.
  - *Evidencia:* `AgCoherenciaGrupo.cs:95` | grep "AgUnidadesSemanales" → 0.
- **Acción a Aplicar:** Implementar `AgUnidadesSemanales.HorasSesionAula` y consultar modalidad.
- **Archivos Afectados:** `AgCoherenciaGrupo.cs`, `AgUnidadesSemanales.cs`.

#### 3. Fase 937 — F925 no aplicada y corrección de fallback en DiasKDe
- **Estado:** APROBADA (no aplicada)
- **Hallazgo:**
  - *Diagnóstico Crítico:* L3984-3985 asigna `"4,5,6"`. Vaciado sin corregir `DiasKDe` L40 forzaría `{2,3}` por fallback.
  - *Causa Raíz:* Dependencia no resuelta en `AgPlanificadorPatronSemanal.cs`.
  - *Impacto Confirmado:* Concentración en martes/miércoles.
  - *Evidencia:* `AgPlanificadorPatronSemanal.cs:40` | grep "new List<int> { 2, 3 }" → 1.
- **Acción a Aplicar:** Eliminar heurística por horas y derivar fallback desde `aa.DiasVentana`.
- **Archivos Afectados:** `AgFormGeneradorHorarios.cs`, `AgPlanificadorPatronSemanal.cs`.

#### 4. Fase 938 — F926 no aplicada: AgValidadorVentana.Validar sin invocarse
- **Estado:** APROBADA (no aplicada)
- **Hallazgo:**
  - *Diagnóstico Crítico:* `dias_patron ⊆ dias_ventana` no evaluado.
  - *Causa Raíz:* Código muerto.
  - *Impacto Confirmado:* Inviabilidad no filtrada.
  - *Evidencia:* `AgValidadorVentana.cs:148` | grep invocación pre-flight → 0.
- **Acción a Aplicar:** Invocar validación en bucle pre-vuelo.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorVentana.cs`.

#### 5. Fase 939 — F927 no aplicada: vacantes excluidas y sin solape de intervalo
- **Estado:** APROBADA (no aplicada)
- **Hallazgo:**
  - *Diagnóstico Crítico:* R4/R5 omiten vacantes y agrupan por `Grupo` plano; solapes por ID de bloque.
  - *Causa Raíz:* Claves no canónicas.
  - *Impacto Confirmado:* Coexistencia de 36 h/día en cohorte.
  - *Evidencia:* `AgValidadorReglas.cs:286` | grep "EsVacante == 0" → 2.
- **Acción a Aplicar:** `ClaveCohorte(g)`, vacantes en ocupación, solape con `AgBloqueSolape.SeSolapan`.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorReglas.cs`.

#### 6. Fase 940 — F928 no aplicada: fitness plano
- **Estado:** APROBADA (no aplicada)
- **Hallazgo:**
  - *Diagnóstico Crítico:* f2 plano no aplica pesos individuales de regla.
  - *Causa Raíz:* Función fitness obsoleta.
  - *Impacto Confirmado:* Presión selectiva uniforme.
  - *Evidencia:* `AgEvaluadorAptitud.cs:33` | grep "PorRegla" → 0.
- **Acción a Aplicar:** Ponderar sumando `contexto.Pesos.PorRegla(v.Regla)`.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgEvaluadorAptitud.cs`.

#### 7. Fase 941 — F929 no aplicada: columna K sin horas
- **Estado:** APROBADA (no aplicada)
- **Hallazgo:**
  - *Diagnóstico Crítico:* L389 exporta solo nombres de días.
  - *Causa Raíz:* Falta de formateador con horas.
  - *Impacto Confirmado:* Columna K no institucional.
  - *Evidencia:* `SemestresExcelExporter.cs:389` | grep "DiasHorasSemanaFormatter" → 0.
- **Acción a Aplicar:** Asignar `DiasHorasSemanaFormatter.DesdeSesiones(sesionesGrupo)`.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/SemestresExcelExporter.cs`.

#### 8. Fase 942 — F930 no aplicada: botón «Proyección 1:1»
- **Estado:** APROBADA (no aplicada)
- **Hallazgo:**
  - *Diagnóstico Crítico:* Texto y handler original intactos.
  - *Causa Raíz:* Sin edición en UI.
  - *Impacto Confirmado:* Generación de CSV obsoleto.
  - *Evidencia:* `AgFormGeneradorHorarios.Designer.cs:840` | grep "Módulo R" → 0.
- **Acción a Aplicar:** Renombrar botón a «📥 Cargar en Módulo R», asociar a `CargarUltimaGeneracionEnModuloRAsync`, borrar CSV espejo.
- **Archivos Afectados:** `AgFormGeneradorHorarios.cs`, `.Designer.cs`.

#### 9. Fase 943 — F931 no aplicada: contrato sin GetUltimaCorridaAsync
- **Estado:** APROBADA (no aplicada)
- **Hallazgo:**
  - *Diagnóstico Crítico:* Método ausente en `IAgScheduleRepository`.
  - *Causa Raíz:* Contrato incompleto.
  - *Impacto Confirmado:* Bloqueante de compilación para F942.
  - *Evidencia:* `IAgScheduleRepository.cs:14` | grep "GetUltimaCorridaAsync" → 0.
- **Acción a Aplicar:** Declarar firma en contrato e implementar en repositorio.
- **Archivos Afectados:** `IAgScheduleRepository.cs`, `AgScheduleRepository.cs`.

#### 10. Fase 944 — F932 no aplicada: materializador sin I3/I4/I5
- **Estado:** APROBADA (no aplicada)
- **Hallazgo:**
  - *Diagnóstico Crítico:* Fechas tomadas de ventana y sin dedupe.
  - *Causa Raíz:* Staging sin compuertas.
  - *Impacto Confirmado:* Sesiones duplicadas.
  - *Evidencia:* `AgMaterializadorSesiones.cs:99` | grep "HashSet<(int, DateTime?, string)>" → 0.
- **Acción a Aplicar:** Integrar compuertas I3, I4, I5 en staging.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgMaterializadorSesiones.cs`.

#### 11. Fase 945 — Clave de agregación de R12 (aula vs cohorte)
- **Estado:** APROBADA (no aplicada)
- **Hallazgo:**
  - *Diagnóstico Crítico:* Ambigüedad en R12; debe evaluar aula + semana ISO, no cohorte.
  - *Causa Raíz:* Definición cruzada en plan.
  - *Impacto Confirmado:* Falsos negativos si 2 aulas de una cohorte dictan materias simultáneas.
  - *Evidencia:* `AgValidadorReglas.cs` | clave R12 pendiente de implementación.
- **Acción a Aplicar:** `porAulaSemana = genes.GroupBy(g => (g.IdAula, AgSemanaIso.De(g.FechasGen)))`.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorReglas.cs`.

#### 12. Fase 946 — Eliminación de Artefacto `scratch\sync_docs.ps1`
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Script auxiliar residual en `scratch\`.
  - *Causa Raíz:* Falta de limpieza post-ejecución.
  - *Impacto Confirmado:* Contaminación de workspace.
  - *Evidencia:* `Test-Path 'scratch\sync_docs.ps1'` → False.
- **Acción Aplicada:** Archivo eliminado mediante `Remove-Item`.
- **Archivo Afectado:** `scratch\sync_docs.ps1`.

#### 13. Fase 947 — Estandarización de Estado Documental con Evidencia
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Fases previas catalogadas como ACEPTADA sin código en disco.
  - *Causa Raíz:* Verificación de redacción en vez de presencia de símbolos.
  - *Impacto Confirmado:* Desalineación plan/realidad.
  - *Evidencia:* `walkthrough.md` actualizado con `APROBADA (no aplicada)` y métrica de símbolos.
- **Acción Aplicada:** Corrección de estados en `walkthrough.md`.
- **Archivo Afectado:** `walkthrough.md`.

---

### Verificación de Principios (Ronda 176)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F935 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgReglaViolacion.cs:35` (grep "R11_" → 0) |
| F936 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgCoherenciaGrupo.cs:95` (grep "AgUnidadesSemanales" → 0) |
| F937 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorPatronSemanal.cs:40` (grep "{ 2, 3 }" → 1) |
| F938 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgValidadorVentana.cs:148` (grep "Validar" en pre-flight → 0) |
| F939 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgValidadorReglas.cs:286` (grep "EsVacante == 0" → 2) |
| F940 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEvaluadorAptitud.cs:33` (grep "PorRegla" → 0) |
| F941 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `SemestresExcelExporter.cs:389` (grep "DiasHorasSemanaFormatter" → 0) |
| F942 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.Designer.cs:840` (grep "Módulo R" → 0) |
| F943 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `IAgScheduleRepository.cs:14` (grep "GetUltimaCorridaAsync" → 0) |
| F944 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgMaterializadorSesiones.cs:99` (grep "HashSet" dedupe → 0) |
| F945 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgValidadorReglas.cs` (clave R12 aula pendiente) |
| F946 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `Test-Path scratch\sync_docs.ps1` → False |
| F947 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `walkthrough.md` alineado con estados reales |

---

## Ronda 177: F935-F945 y F947 sin aplicar — y 6 defectos de API/estructura que impedirían compilar (Fases 948 a 953)

### Evaluación Crítica de las Fases

#### 1. Fase 948 — AgPesosFitness sin método PorRegla (CS1061)
- **Estado:** APROBADA (no aplicada)
- **Hallazgo:**
  - *Diagnóstico Crítico:* `AgPesosFitness` carece de método `PorRegla(AgReglaViolacion)` y propiedades para R11-R14, S6, S7.
  - *Causa Raíz:* F940 asumió API inexistente.
  - *Impacto Confirmado:* Error de compilación CS1061 al compilar F940.
  - *Evidencia:* `AgPesosFitness.cs:11` | grep "PorRegla" → 0.
- **Acción a Aplicar:** Implementar `PorRegla(AgReglaViolacion)` con switch expression cuidando de no duplicar case para enums de idéntico valor numérico (CS0152).
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgPesosFitness.cs`.

#### 2. Fase 949 — F937 referencia aa.DiasVentana inexistente en AgAulaAsignatura (CS1061)
- **Estado:** APROBADA (no aplicada)
- **Hallazgo:**
  - *Diagnóstico Crítico:* `AgAulaAsignatura` no expone `DiasVentana` (reside en `AgAula`).
  - *Causa Raíz:* Confusión de modelos POCO en F937.
  - *Impacto Confirmado:* Error de compilación CS1061.
  - *Evidencia:* `AgAulaAsignatura.cs:18` | grep "DiasVentana" → 0; `AgAula.cs:35` | grep "DiasVentana" → 1.
- **Acción a Aplicar:** Adaptar firma de `DiasKDe` a `DiasKDe(AgAulaAsignatura aa, AgAula aula = null)` y consultar `aula?.DiasVentana`.
- **Archivos Afectados:** `AgAulaAsignatura.cs`, `AgPlanificadorPatronSemanal.cs`.

#### 3. Fase 950 — AgSemanaIso y AgBloqueSolape inexistentes (CS0246)
- **Estado:** APROBADA (no aplicada)
- **Hallazgo:**
  - *Diagnóstico Crítico:* F937 (`AgSemanaIso.DesdeVentana`), F939 (`AgBloqueSolape.SeSolapan`) y F945 (`AgSemanaIso.De`) dependen de tipos no creados.
  - *Causa Raíz:* Inversión de dependencias; parches propuestos antes de crear tipos base.
  - *Impacto Confirmado:* Error de compilación CS0246 en cascada.
  - *Evidencia:* `ConversorPDF.XLL` | Test-Path `AgSemanaIso.cs` → False; Test-Path `AgBloqueSolape.cs` → False.
- **Acción a Aplicar:** Implementar clases estáticas `AgSemanaIso` y `AgBloqueSolape` antes de los consumidores.
- **Archivos Afectados:** `AgSemanaIso.cs`, `AgBloqueSolape.cs`.

#### 4. Fase 951 — Fechas vacías en genes sin expandir anulan R12
- **Estado:** APROBADA (no aplicada)
- **Hallazgo:**
  - *Diagnóstico Crítico:* `AgGen.FechasGen` nace vacío; derivar semana antes de expansión agrupa en clave nula o satura el grupo.
  - *Causa Raíz:* Validación ejecutada antes de poblar fechas reales.
  - *Impacto Confirmado:* R12 nunca penaliza o genera falsos positivos.
  - *Evidencia:* `AgGen.cs:68` | `FechasGen = new List<DateTime>()`.
- **Acción a Aplicar:** Filtrar genes con `FechasGen != null && FechasGen.Count > 0` y agrupar por `AgSemanaIso.Clave(g.FechasGen[0])`.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorReglas.cs`.

#### 5. Fase 952 — Fábrica colapsa carriles de cohorte a una sola aula
- **Estado:** APROBADA (no aplicada)
- **Hallazgo:**
  - *Diagnóstico Crítico:* `aulasValidas.GroupBy(ClaveCohorteAula).Select(g => g.First())` descarta la segunda aula de carriles paralelos.
  - *Causa Raíz:* Deduplicación legacy de Fases 454/475/478 para prevenir cohortes duplicadas.
  - *Impacto Confirmado:* Destruye el modelo de carriles antes de iniciar la optimización.
  - *Evidencia:* `AgFabricaIndividuos.cs:94` | `GroupBy(ClaveCohorteAula).Select(g => g.First())`.
- **Acción a Aplicar:** Agrupar considerando jornada y ventana: `GroupBy(a => (ClaveCohorteAula(a), a.Jornada ?? "", a.DiasVentana ?? ""))`.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`.

#### 6. Fase 953 — Unificación de definición de cohorte
- **Estado:** APROBADA (no aplicada)
- **Hallazgo:**
  - *Diagnóstico Crítico:* Fábrica define cohorte por `AulaRef` o `prog|sem|grp`; validador usa `sede|prog|sem|grp`.
  - *Causa Raíz:* Función local no promovida a dominio.
  - *Impacto Confirmado:* Discrepancias diagnósticas entre módulos.
  - *Evidencia:* `AgFabricaIndividuos.cs:88` vs `AgValidadorReglas.cs:208`.
- **Acción a Aplicar:** Crear clase estática `AgClaveCohorte` en dominio y reutilizar en fábrica y validador.
- **Archivos Afectados:** `AgClaveCohorte.cs`, `AgFabricaIndividuos.cs`, `AgValidadorReglas.cs`.

---

### Verificación de Principios (Ronda 177)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F948 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPesosFitness.cs:11` (grep "PorRegla" → 0) |
| F949 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgAulaAsignatura.cs:18` (grep "DiasVentana" → 0) |
| F950 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | Test-Path `AgSemanaIso.cs` / `AgBloqueSolape.cs` → False |
| F951 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgGen.cs:68` (`FechasGen = new List<DateTime>()`) |
| F952 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFabricaIndividuos.cs:94` (`GroupBy(ClaveCohorteAula).Select(g => g.First())`) |
| F953 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFabricaIndividuos.cs:88` (divergencia con AgValidadorReglas) |

---

## Ronda 178: F948-F953 sin aplicar — F938 y F943 fallarían en ejecución (Fases 954 a 959)

### Evaluación Crítica de las Fases

#### 1. Fase 954 — F938 inerte con DiasPatron vacío (guard L151)
- **Estado:** APROBADA (no aplicada)
- **Hallazgo:**
  - *Diagnóstico Crítico:* `AgValidadorVentana.Validar` (L151) requiere `!string.IsNullOrWhiteSpace(diasPatron)`. Con `diasPatron = string.Empty` la contención D14 nunca se evalúa.
  - *Causa Raíz:* Validación ejecutada sobre el campo sin resolver en lugar de los días efectivos.
  - *Impacto Confirmado:* Pre-flight aprobaría silenciosamente asignaciones fuera de ventana.
  - *Evidencia:* `AgValidadorVentana.cs:151` | `if (!string.IsNullOrWhiteSpace(diasVentana) && !string.IsNullOrWhiteSpace(diasPatron))`.
- **Acción a Aplicar:** Resolver días efectivos con `AgPlanificadorPatronSemanal.DiasKDe(aa, aula)` y pasarlos como `string.Join(",", diasEfectivos)`.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorVentana.cs`.

#### 2. Fase 955 — Corrección de Nombre de Tabla en Consulta Dapper (`ag_corrida`)
- **Estado:** APROBADA (no aplicada)
- **Hallazgo:**
  - *Diagnóstico Crítico:* El DDL crea la tabla `ag_corrida` (singular); una consulta a `ag_corridas` dispara error SQLite `no such table`.
  - *Causa Raíz:* Redacción de especificación sin verificar el DDL de `MigrationV49`.
  - *Impacto Confirmado:* Excepción runtime en `GetUltimaCorridaAsync`.
  - *Evidencia:* `MigrationV49_AddAgHorariosTables.cs:68` | `CREATE TABLE IF NOT EXISTS ag_corrida`.
- **Acción a Aplicar:** Implementar SQL explícito contra `ag_corrida` con mapeo a DTO `AgCorrida`.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgScheduleRepository.cs`.

#### 3. Fase 956 — Migración V70 para Columna `vacantes` en `ag_corrida`
- **Estado:** APROBADA (no aplicada)
- **Hallazgo:**
  - *Diagnóstico Crítico:* `AgCorrida.Vacantes` existe en POCO pero carece de columna en la tabla `ag_corrida` y en el `INSERT`.
  - *Causa Raíz:* Adición de propiedad en el modelo sin migración de base de datos asociada.
  - *Impacto Confirmado:* Error `no such column` en lecturas directas o pérdida silenciosa del valor en persistencia.
  - *Evidencia:* `AgCorrida.cs:28` | `public int Vacantes { get; set; }`; `MigrationV49:68-88` sin columna `vacantes`.
- **Acción a Aplicar:** Crear `MigrationV70_AddAgCorridaVacantes` con `ALTER TABLE ag_corrida ADD COLUMN vacantes INTEGER NOT NULL DEFAULT 0;` y actualizar `INSERT` / `SELECT`.
- **Archivos Afectados:** `MigrationV70_AddAgCorridaVacantes.cs`, `AgScheduleRepository.cs`.

#### 4. Fase 957 — Detección Temporal de Solape en `AgBloqueSolape.SeSolapan` por Intervalo
- **Estado:** APROBADA (no aplicada)
- **Hallazgo:**
  - *Diagnóstico Crítico:* Comparar la columna de texto `Franja` ("MAÑANA", "TARDE") produce falsos solapes entre bloques disjuntos del mismo día y falsos negativos en horarios cruzados.
  - *Causa Raíz:* Evaluación cualitativa de franja en lugar de parsing horario de `HoraInicio`/`HoraFin`.
  - *Impacto Confirmado:* Detección errática de solapes temporales.
  - *Evidencia:* `AgBloqueHorario.cs:32-35` | columnas `HoraInicio`, `HoraFin`, `Franja`.
- **Acción a Aplicar:** Evaluar coincidencia de `DiaSemanaNumero` e intersección temporal matemática `i1 < f2 && i2 < f1` parseando con `TimeSpan`.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgBloqueSolape.cs`.

#### 5. Fase 958 — Delegación de `AgSemanaIso.DesdeVentana` en `AgDiasVentana.Dias`
- **Estado:** APROBADA (no aplicada)
- **Hallazgo:**
  - *Diagnóstico Crítico:* Parser canónico de ventanas ya existe en `AgDiasVentana.Dias(ventana)`; duplicarlo genera redundancia de código.
  - *Causa Raíz:* Desconocimiento de helper preexistente en F950.
  - *Impacto Confirmado:* Divergencia potencial de reglas de parsing.
  - *Evidencia:* `AgDiasVentana.cs:28` | `public static IReadOnlyList<int> Dias(string ventana)`.
- **Acción a Aplicar:** Implementar `AgSemanaIso.DesdeVentana` como delegación directa: `AgDiasVentana.Dias(diasVentana)`.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgSemanaIso.cs`.

#### 6. Fase 959 — Resolución de Capacidad en `AgValidadorVentana` con Aula Efectiva
- **Estado:** APROBADA (no aplicada)
- **Hallazgo:**
  - *Diagnóstico Crítico:* `CapacidadDe(a)` en L95 y reporte en L114 invocan `DiasKDe(a)` sin pasar `aula`, cayendo en fallback `{2,3}` y subestimando o sobreestimando capacidad real.
  - *Causa Raíz:* Consumidores no adaptados a la nueva firma de `DiasKDe(aa, aula)`.
  - *Impacto Confirmado:* Bloqueos espurios de optimización por falso cálculo de insuficiencia de capacidad.
  - *Evidencia:* `AgValidadorVentana.cs:95` | `AgPlanificadorPatronSemanal.DiasKDe(a)`.
- **Acción a Aplicar:** Pasar `AulaDe(a.IdAula)` en `CapacidadDe` y en la generación de mensajes diagnósticos.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorVentana.cs`.

---

### Verificación de Principios (Ronda 178)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F954 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgValidadorVentana.cs:151` (`!string.IsNullOrWhiteSpace(diasPatron)`) |
| F955 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `MigrationV49_AddAgHorariosTables.cs:68` (tabla `ag_corrida` singular) |
| F956 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgCorrida.cs:28` (`Vacantes`) / DDL sin columna `vacantes` |
| F957 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgBloqueHorario.cs:32-35` (`HoraInicio`, `HoraFin` vs `Franja`) |
| F958 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgDiasVentana.cs:28` (`Dias` ya implementado) |
| F959 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgValidadorVentana.cs:95,114` (llamadas a `DiasKDe` sin aula) |

---

## Ronda 179: F954-F959 sin aplicar — 4 defectos nuevos en la cadena de siembra/materialización

### Evaluación Crítica de las Fases

#### 1. Fase 960 — Blindaje ante `fechasesion TEXT NOT NULL` en `ag_sesiones`
- **Estado:** APROBADA (no aplicada)
- **Hallazgo:**
  - *Diagnóstico Crítico:* `MigrationV49:207` define `fechasesion TEXT NOT NULL`. `AgSesion.FechaSesion` es `DateTime?`. Una sesión sin fecha lanza excepción SQLite 19.
  - *Causa Raíz:* DDL sin valor por defecto y mapping directo de propiedad nullable.
  - *Impacto Confirmado:* Fallo catastrófico en `InsertSesionesBulkAsync` y pérdida total del lote de sesiones.
  - *Evidencia:* `MigrationV49_AddAgHorariosTables.cs:207` | `fechasesion TEXT NOT NULL`; `AgSesion.cs:15` | `public DateTime? FechaSesion { get; set; }`.
- **Acción a Aplicar:** Descartar sesiones sin fecha en `AgMaterializadorSesiones` (`I3 gen {s.IdGen}: sesión sin fecha`) y en `AgScheduleRepository` aplicar fallback `(s.FechaSesion ?? DateTime.MinValue).ToString("yyyy-MM-dd")`.
- **Archivos Afectados:** `AgScheduleRepository.cs`, `AgMaterializadorSesiones.cs`.

#### 2. Fase 961 — Bloques de 4 h en Pregrado Durante la Siembra
- **Estado:** APROBADA (no aplicada)
- **Hallazgo:**
  - *Diagnóstico Crítico:* `AgFabricaIndividuos.cs:270,324` llama `AgFranjaJornada.Resolve(aula.Jornada, aula.HorasDia)` usando `aula.HorasDia` (8h) directamente, asignando bloques de 8h a aulas de PREGRADO.
  - *Causa Raíz:* Omisión del nivel académico en los dos puntos de siembra de la fábrica.
  - *Impacto Confirmado:* Violación de regla D4′ y rechazo sistemático por compuerta I3 en `AgMaterializadorSesiones`.
  - *Evidencia:* `AgFabricaIndividuos.cs:270,324` | `AgFranjaJornada.Resolve(aula.Jornada, aula.HorasDia)`.
- **Acción a Aplicar:** Resolver `nivel = AgNivelAcademicoCatalogo.Desde(...)` y usar `AgUnidadesSemanales.HorasSesionAula(aula, nivel)` (4h en pregrado) en L270 y L324.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`.

#### 3. Fase 962 — Discriminación por Día en `AgFranjaJornada.Resolve` (Regla D6)
- **Estado:** APROBADA (no aplicada)
- **Hallazgo:**
  - *Diagnóstico Crítico:* `Resolve(jornada, horasDia)` carece de día ISO. D6 (sábado POSGRADO-8h = bloque único `08:00–17:00`) no puede imponerse: con `TARDE`+8h devuelve `14:00–22:00` en sábado; además `TARDE` 8h y `NOCHE` 8h comparten `14:00–22:00` idénticos.
  - *Causa Raíz:* Firma sin discriminador de día de la semana.
  - *Impacto Confirmado:* Horarios imposibles en sábado POSGRADO y solapes falsos en `AgBloqueSolape`.
  - *Evidencia:* `AgFranjaJornada.cs:19-35` | `Resolve(string jornada, int horasDia)`.
- **Acción a Aplicar:** Extender firma a `Resolve(string jornada, int horasDia, int diaIso = 0)` retornando `08:00–17:00` MAÑANA si `diaIso == 6 && horasDia == 8`.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgFranjaJornada.cs`.

#### 4. Fase 963 — Persistencia de `dias_patron` en `InsertAulaAsignaturasAsync`
- **Estado:** APROBADA (no aplicada)
- **Hallazgo:**
  - *Diagnóstico Crítico:* Columna `dias_patron` existe en `ag_aula_asignatura` (`MigrationV55:33`), pero `AgScheduleRepository.cs:190` (`InsertAulaAsignaturasAsync`) no la incluye en su `INSERT`.
  - *Causa Raíz:* Repositorio desincronizado con migración V55.
  - *Impacto Confirmado:* Registros creados por esta ruta guardan `dias_patron=''` anulando D14.
  - *Evidencia:* `AgScheduleRepository.cs:190` | `INSERT OR IGNORE INTO ag_aula_asignatura (id_aula, id_asignatura, exclusivo, prioridad, subunidad)`; `MigrationV55:33` | columna `dias_patron`.
- **Acción a Aplicar:** Actualizar SQL `INSERT OR IGNORE INTO ag_aula_asignatura (..., dias_patron, ...) VALUES (..., COALESCE(@DiasPatron,''), ...)`.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgScheduleRepository.cs`.

#### 5. Fase 964 — Cuantización Estricta de Sesiones en `AgExpansorSesiones`
- **Estado:** APROBADA (no aplicada)
- **Hallazgo:**
  - *Diagnóstico Crítico:* `Math.Min(bloque.DuracionHoras, horasFaltantes)` genera sesiones residuales de duración fraccionaria si `HorasGen` no es múltiplo de la unidad de bloque, provocando rechazo masivo por compuerta I3 en `AgMaterializadorSesiones` o sesiones duplicadas en sábado D6.
  - *Causa Raíz:* Expansor tolera residuos en lugar de verificar cuantización discreta.
  - *Impacto Confirmado:* Sesiones de duraciones anómalas (ej. 2h en bloque de 4h) y envío del grupo a cuarentena.
  - *Evidencia:* `AgExpansorSesiones.cs:91` | `double horasSesion = Math.Min(bloque.DuracionHoras, horasFaltantes);`.
- **Acción a Aplicar:** Exigir `horasObjetivoGen % bloque.DuracionHoras == 0` con guard fail-fast y expandir cada sesión con `horasSesion = bloque.DuracionHoras` fijo sin `Math.Min`.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgExpansorSesiones.cs`.

---

### Verificación de Principios (Ronda 179)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F960 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `MigrationV49:207` (`fechasesion TEXT NOT NULL`) / `AgSesion.cs:15` |
| F961 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFabricaIndividuos.cs:270,324` (`HorasDia` crudo sin nivel) |
| F962 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFranjaJornada.cs:19-35` (`Resolve` sin parámetro `diaIso`) |
| F963 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgScheduleRepository.cs:190` (`INSERT` sin `dias_patron`) |
| F964 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgExpansorSesiones.cs:91` (`Math.Min` con residuo fraccionario) |

---

## Ronda 181: F964 aplicada con error de compilación + 4 defectos en código nuevo

### Evaluación Crítica de las Fases

#### 1. Fase 965 — `horasAcumuladas` eliminada por F964 → CS0103 ×4
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La variable acumuladora `horasAcumuladas` fue eliminada en `AgExpansorSesiones.cs:84` al aplicar la cuantización de F964, dejando 4 referencias sin resolver e impidiendo compilar (CS0103).
  - *Causa Raíz:* Parche sobreescribió la línea de declaración del acumulador.
  - *Impacto Confirmado:* Error de compilación bloqueante en todo el proyecto XLL.
  - *Evidencia:* `AgExpansorSesiones.cs:84` | ausencia de `double horasAcumuladas = 0.0;`.
- **Acción Aplicada:** Declarada `double horasAcumuladas = 0.0;` tras la instanciación de `sesiones`.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgExpansorSesiones.cs`.

#### 2. Fase 966 — `GetUltimaCorridaAsync` declarada sin implementar → CS0535
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* `IAgScheduleRepository` declaró `GetUltimaCorridaAsync`, pero `AgScheduleRepository` no contenía su implementación, quebrando la satisfacción del contrato de interfaz (CS0535).
  - *Causa Raíz:* Modificación confinada exclusivamente a la interfaz en F955.
  - *Impacto Confirmado:* Error de compilación en el implementador de infraestructura.
  - *Evidencia:* `IAgScheduleRepository.cs:23` | `Task<Result<AgCorrida>> GetUltimaCorridaAsync(...)`.
- **Acción Aplicada:** Implementado método `GetUltimaCorridaAsync` con Dapper `QuerySingleOrDefaultAsync<AgCorrida>` proyectando `vacantes as Vacantes`, `parametros_json`, filtrado por `idperiodo`, `idprograma` y `subunidad` con `ORDER BY id DESC LIMIT 1`.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgScheduleRepository.cs`.

#### 3. Fase 968 — Coherencia cuantiza con `4.0` fijo → posgrado 8 h viola F964 y aborta el individuo
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* `AgCoherenciaGrupo.cs` empleaba la constante hardcodeada `4.0` para calcular `totalSesiones`, proyectar la suma y ajustar las horas de los genes. En posgrados con sesiones de 8h, generaba cargas no múltiplos de 8h (ej. 20h), activando el guard fail-fast de F964 y abortando la generación.
  - *Causa Raíz:* F936 limitó la resolución dinámica de unidad a la rama sin ajustables.
  - *Impacto Confirmado:* Aborto total del expansor de sesiones en todos los programas de posgrado.
  - *Evidencia:* `AgCoherenciaGrupo.cs:76-90,140,155,172`.
- **Acción Aplicada:** Promovida función local `ObtenerUnidad(AgGen g)` a nivel de grupo resolviendo `AgUnidadesSemanales.HorasSesionAula(a, nivel)` y reemplazados todos los factores `4.0` por `unidad` dinámica.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgCoherenciaGrupo.cs`.

#### 4. Fase 967 — `AgBloqueSolape` y `R14_BloqueSolapadoPorIntervalo` sin emisor → regla inerte
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La regla `R14_BloqueSolapadoPorIntervalo` (código 33) y el helper matemático `AgBloqueSolape.SeSolapan` no tenían ningún punto de llamada en `AgValidadorReglas.cs`, permitiendo que bloques solapados pasaran la compuerta.
  - *Causa Raíz:* Utilidad creada en F950/F957 sin punto de integración en el validador.
  - *Impacto Confirmado:* Solapamientos horarios indetectados durante la evaluación de aptitud y factibilidad.
  - *Evidencia:* `AgBloqueSolape.cs` con 0 callers; `AgValidadorReglas.cs` sin emisión de regla 33.
- **Acción Aplicada:** Conectado bucle de evaluación de solapamiento temporal en `AgValidadorReglas.Evaluar` sobre genes con bloques cargados en `contexto.BloquesPorId`, verificando recursos compartidos (aula, docente o cohorte/grupo) e invocando `AgBloqueSolape.SeSolapan(b1, b2)` para registrar `AgViolacion(R14_BloqueSolapadoPorIntervalo)` e incrementar `hardCount`.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorReglas.cs`.

#### 5. Fase 969 — Un gen inválido cancela la materialización completa del individuo
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* `AgExpansorSesiones.ExpandirIndividuo` retornaba `Result.Failure` inmediato al encontrar el primer gen no cuantizado, descartando todas las sesiones válidas sin aislar el fallo ni reportar qué genes fallaron.
  - *Causa Raíz:* Bucle de expansión sin acumulación ni aislamiento de fallos individuales.
  - *Impacto Confirmado:* Pérdida total de sesiones materializadas ante un único gen atípico.
  - *Evidencia:* `AgExpansorSesiones.cs:156-161`.
- **Acción Aplicada:** Declarada `var todasSesiones = new List<AgSesion>();`, acumulados errores por gen con `continue`, registrada advertencia con `LoggerHelper.Warning` y retornado `Failure` únicamente si no se pudo expandir ninguna sesión.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgExpansorSesiones.cs`.

---

### Verificación de Principios (Ronda 181)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F965 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgExpansorSesiones.cs:84` (`double horasAcumuladas = 0.0;`) |
| F966 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgScheduleRepository.cs:369-400` (`GetUltimaCorridaAsync`) |
| F968 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgCoherenciaGrupo.cs:78-89,142,155,172` (`ObtenerUnidad`, `unidad`) |
| F967 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgValidadorReglas.cs:462-498` (`AgBloqueSolape.SeSolapan`) |
| F969 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgExpansorSesiones.cs:151-176` (`todasSesiones`, `errores`) |

---

## Ronda 182: F965-F969 aplicadas — 3 defectos nuevos en F967/F968

### Evaluación Crítica de las Fases

#### 1. Fase 970 — F967 compara cohortes entre programas/sedes distintas → R14 hard masivo
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La condición de cohorte previa en `mismoRecurso` (`g1.Grupo == g2.Grupo && g1.Semestre == g2.Semestre`) omitía el programa (`IdPrograma`) y la sede (`IdSede`). En bases con múltiples programas y sedes que comparten nomenclatura de grupos (ej. Semestre "1", Grupo "1"), se detectaban solapes irreales como infracciones duras de `R14_BloqueSolapadoPorIntervalo`, tornando casi cualquier individuo en `Infactible` e impidiendo la convergencia del algoritmo genético.
  - *Causa Raíz:* Comparación de cohorte sin aislar `IdPrograma` ni `IdSede`.
  - *Impacto Confirmado:* Falsos positivos masivos en `R14` y bloqueo total de convergencia en el GA.
  - *Evidencia:* `AgValidadorReglas.cs:480-484`.
- **Acción Aplicada:** Agregada la validación estricta de cohorte requiriendo `g1.IdPrograma == g2.IdPrograma && g1.IdSede == g2.IdSede && !string.IsNullOrWhiteSpace(g1.Grupo) && !string.IsNullOrWhiteSpace(g1.Semestre) && g1.Grupo == g2.Grupo && g1.Semestre == g2.Semestre`.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorReglas.cs`.

#### 2. Fase 971 — F967 evalúa O(n²) sobre todos los genes → bloqueo en poblaciones reales
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Doble bucle cuadrático $O(n^2)$ comparando todos los pares de genes con bloque sin particionamiento temporal previo. Con 2.000 genes, esto generaba ~2M de comparaciones por individuo por generación ($\approx 6\times 10^9$ por corrida), causando congelamiento del hilo de ejecución en Excel/WinForms.
  - *Causa Raíz:* Ausencia de partición previa por día antes de comparar combinatoria de pares.
  - *Impacto Confirmado:* Congelamiento total de la ejecución y bloqueo de la UI en poblaciones reales.
  - *Evidencia:* `AgValidadorReglas.cs:470-496`.
- **Acción Aplicada:** Particionado previo de genes con bloque agrupando por día (`conBloque.GroupBy(x => x.Bloque.DiaSemanaNumero)`), reduciendo el espacio de búsqueda por día e incorporando corte temprano del bucle interno (`cortado = true; break;`) tras la primera violación detectada por gen.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorReglas.cs`.

#### 3. Fase 972 — F968 usa unidad de 8 h sin validar múltiplo → I1 queda irresoluble
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La selección de unidad por mayoría estadística (8h en posgrado) asignaba `unidad = 8.0` a grupos cuya carga total esperada no era múltiplo de 8h (ej. 60h, resultando en 8 sesiones $\times$ 8h = 64h o residuo no divisible de -4h). Esto producía una discrepancia insalvable $\Sigma \neq I1$ en la coherencia de grupo y convertía la solución en permanentemente infactible.
  - *Causa Raíz:* Selección de unidad modal sin comprobación de divisibilidad de la carga total.
  - *Impacto Confirmado:* Violación dura irreversible de la regla I1 en asignaturas de posgrado con carga horaria no divisible entre 8.
  - *Evidencia:* `AgCoherenciaGrupo.cs:89-91`.
- **Acción Aplicada:** Incorporada verificación de divisibilidad `Math.Abs(esperado / unidad - Math.Round(esperado / unidad)) > 0.001`; si no es múltiplo exacto de 8h, se degrada a `unidad = 4.0` (regla D4′ divisible).
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgCoherenciaGrupo.cs`.

---

### Verificación de Principios (Ronda 182)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F970 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgValidadorReglas.cs:486-488` (`IdPrograma == IdPrograma && IdSede == IdSede`) |
| F971 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgValidadorReglas.cs:469-509` (`GroupBy(DiaSemanaNumero)`) |
| F972 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgCoherenciaGrupo.cs:91-95` (`esperado / unidad - Round > 0.001 -> unidad = 4.0`) |

---

## Ronda 183: F970-F972 aplicadas correctamente — 2 divergencias unidad↔carga + 1 warning

### Evaluación Crítica de las Fases

#### 1. Fase 973 — Fábrica asigna bloque de 8 h sin validar que la carga sea múltiplo → F964 rechaza y F969 omite el gen
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La fábrica seleccionaba franjas de 8h (`horasEfectivas = HorasSesionAula(aula, nivelAula)`) para asignaturas en posgrado basándose solo en el tipo de aula, sin verificar si la carga total (`totalHoras`) o la carga del día en multi-día (`totalHoras / diasK.Count`) era divisible entre 8h. En asignaturas de 60h, se asignaba un bloque de 8h pero la coherencia degradaba la unidad a 4h, provocando que `AgExpansorSesiones` (F964) rechazara la cuantización (`60 / 8 = 7.5`), F969 omitiera el gen y la corrida fallara por cobertura incompleta de grupo.
  - *Causa Raíz:* Asignación de duración de bloque en la fábrica sin comprobación de divisibilidad de la carga.
  - *Impacto Confirmado:* F964 devolvía `Failure`, F969 omitía el gen y se quebraban invariantes I1/I2 en staging.
  - *Evidencia:* `AgFabricaIndividuos.cs:264-273` y `322-331`.
- **Acción Aplicada:** Evaluada divisibilidad `cargaUnidad / horasEfectivas` tanto en la rama de día único (`totalHoras`) como en la de multi-día (`totalHoras / diasK.Count`), degradando a `horasEfectivas = 4` cuando `horasEfectivas > 4` y la carga no es múltiplo exacto de 8h.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`.

#### 2. Fase 974 — F972 valida `esperado` pero no `horasRestantes` → residuo de fijados rompe I1
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* En grupos con genes fijados donde la carga total `esperado` es divisible entre 8 (ej. 64h), si los fijados suman una cantidad no múltiplo de 8 (ej. 20h), el remanente `horasRestantes` (44h) no puede cuantizarse en bloques de 8h. `AgCoherenciaGrupo` mantenía `unidad = 8.0`, generando una discrepancia de 4h que el ajuste residual no lograba reconciliar, resultando en violación dura permanente de I1.
  - *Causa Raíz:* Comprobación de divisibilidad de F972 restringida a `esperado` sin validar el remanente `horasRestantes`.
  - *Impacto Confirmado:* Individuo infactible permanente en grupos mixtos (fijados + ajustables) de posgrado.
  - *Evidencia:* `AgCoherenciaGrupo.cs:91-95`.
- **Acción Aplicada:** Extendida la guarda de degradación a `unidad = 4.0` para que evalúe si `esperado` o `horasRestantes` no son múltiplos exactos de `unidad`.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgCoherenciaGrupo.cs`.

#### 3. Fase 975 — `cortado` + `continue` es código muerto en F971
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La variable `bool cortado` y la sentencia `if (cortado) continue;` al final del bucle `for i` en la validación R14 constituían código muerto sin efecto operacional, pues `break` ya interrumpía el bucle interno `j` e `i` continuaba naturalmente al siguiente elemento.
  - *Causa Raíz:* Parche F971 incluyó guardia redundante.
  - *Impacto Confirmado:* Ruido sintáctico y advertencias de análisis estático.
  - *Evidencia:* `AgValidadorReglas.cs:478,504,509`.
- **Acción Aplicada:** Eliminada la variable `cortado` y la condición `if (cortado) continue;`, preservando únicamente el `break;` dentro del bucle de pares.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorReglas.cs`.

---

### Verificación de Principios (Ronda 183)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F973 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFabricaIndividuos.cs:269-274,326-331` (`cargaUnidad / horasEfectivas`) |
| F974 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgCoherenciaGrupo.cs:91-95` (`horasRestantes / unidad`) |
| F975 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgValidadorReglas.cs:474-508` (`break;` limpio sin `cortado`) |

---

## Ronda 184: F973-F975 aplicadas — 2 vestigios de 4 h fijo en la fábrica (rama multi-día)

### Evaluación Crítica de las Fases

#### 1. Fase 976 — `HorasSemanales` fijo en 4.0 contradice bloque de 8 h → carga por gen subestimada
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La fábrica de individuos asignaba bloques de 8h a programas de posgrado (D4′) pero mantenía `HorasSemanales = 4.0` codificado como invariante rígido. Esto causaba que `SesionesDelGen()` y la proyección `sesionesEsteDia * HorasSemanales` calcularan con 4h en lugar de 8h (ej. 8 sesiones de 8h proyectaban 32h en vez de 64h), originando ajustes residuales arbitrarios que podían introducir asignaciones no divisibles entre 8h, provocando rechazo fail-fast en F964 y omisión de genes en F969.
  - *Causa Raíz:* Invariante histórico rígido de sesión elemental de 4.0h no sincronizado con bloques de 8h.
  - *Impacto Confirmado:* Subestimación masiva de carga horaria inicial por gen y posterior distorsión del reparto I1.
  - *Evidencia:* `AgFabricaIndividuos.cs:241` y `396-403`.
- **Acción Aplicada:** Eliminado `HorasSemanales = 4.0` del inicializador de objeto en día único y sincronizada la propiedad `gen.HorasSemanales = bloque.DuracionHoras > 0 ? bloque.DuracionHoras : 4.0` dinámicamente tras la asignación del bloque tanto en la rama de día único como en la de multi-día.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`.

#### 2. Fase 977 — Fallback `totalHoras / diasK.Count` sin cuantizar → F964 rechaza el gen
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Cuando un gen no contaba con sesiones calculadas por fechas o histograma (`sesionesEsteDia == 0`), la rama de fallback dividía linealmente `totalHoras / diasK.Count`. Con cargas como 20h repartidas en 2 días (10h por día) sobre bloques de 4h u 8h, resultaban valores continuos no múltiplos de la unidad de bloque (10 / 4 = 2.5), provocando rechazo en F964 y exclusión del gen por F969.
  - *Causa Raíz:* Reparto continuo de horas sin cuantización al tamaño discreto de bloque.
  - *Impacto Confirmado:* Fallo en la compuerta de cuantización discreta F964 y omisión del gen en staging.
  - *Evidencia:* `AgFabricaIndividuos.cs:434-436`.
- **Acción Aplicada:** Cuantización discreta de la carga del fallback `cargaDia = Math.Max(unidadDia, Math.Round(cargaDia / unidadDia) * unidadDia)` con base en `unidadDia = gen.HorasSemanales > 0 ? gen.HorasSemanales : 4.0`, garantizando que `HorasGen` sea siempre múltiplo entero y positivo de la unidad.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`.

---

### Verificación de Principios (Ronda 184)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F976 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFabricaIndividuos.cs:298-305,396-403` (`gen.HorasSemanales = DuracionHoras`) |
| F977 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFabricaIndividuos.cs:430-440` (`Math.Round(cargaDia / unidadDia) * unidadDia`) |

---

## Ronda 185: F976-F977 aplicadas — 2 bloques falsos en `AgValidadorVentana` por 4 h fijo

### Evaluación Crítica de las Fases

#### 1. Fase 978 — `sesionesRequeridas = ceil(hAsig / 4.0)` bloquea posgrado 8 h
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* En programas de posgrado (`HorasDia = 8`), una asignatura de 64h requiere 8 sesiones de 8h, mientras que la capacidad real se mide en ocurrencias de días K semanales dentro de la ventana. Con el divisor rígido `4.0`, el validador calculaba $64 / 4 = 16$ sesiones requeridas. En una ventana de 12 semanas (capacidad real = 12 ocurrencias), $16 > 12$ generaba un bloqueo falso `16 sesiones > capacidad real 12` y abortaba la corrida, a pesar de que 8 sesiones caben perfectamente en las 12 ocurrencias.
  - *Causa Raíz:* Divisor fijo de 4.0h en `sesionesRequeridas` sin aplicar la regla de unidad efectiva por nivel académico (D4′).
  - *Impacto Confirmado:* Bloqueo falso e infactibilidad artificial de asignaturas de posgrado en ventanas semestrales válidas.
  - *Evidencia:* `AgValidadorVentana.cs:93`.
- **Acción Aplicada:** Derivada `unidadAsig` a partir del aula y su nivel académico con `AgUnidadesSemanales.HorasSesionAula(aulaAsig, AgNivelAcademicoCatalogo.Desde(progAsig?.Modalidad))`; cálculo de `sesionesRequeridas = (int)Math.Ceiling(hAsig / unidadAsig)`.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorVentana.cs`.

#### 2. Fase 979 — Validación D14 usa `auCfg.HorasDia` crudo → exige bloque de 8 h inexistente
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La validación de contención y existencia de bloque D14 invocaba `Validar` con `auCfg.HorasDia` sin filtrar por nivel. En aulas de pregrado con `horas_dia = 8`, esto forzaba la búsqueda de un bloque horario de 8h (`horas_dia = 8`), el cual no existe en el catálogo para pregrado (donde los bloques son de 4h según D4′), provocando un error bloqueante `No existe ningún bloque horario activo para la jornada ... con duración 8 horas`.
  - *Causa Raíz:* Parámetro `horasDia` evaluado directamente desde `auCfg.HorasDia` crudo en lugar de la duración efectiva.
  - *Impacto Confirmado:* Bloqueo falso de pre-vuelo en aulas de pregrado con atributo `horas_dia = 8`.
  - *Evidencia:* `AgValidadorVentana.cs:128`.
- **Acción Aplicada:** Resuelto `horasEfectivasCfg = (int)AgUnidadesSemanales.HorasSesionAula(auCfg, AgNivelAcademicoCatalogo.Desde(progCfg?.Modalidad))` y transmitido a `Validar` en sustitución de `auCfg.HorasDia`.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorVentana.cs`.

---

### Verificación de Principios (Ronda 185)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F978 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgValidadorVentana.cs:94-101` (`unidadAsig`, `sesionesRequeridas`) |
| F979 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgValidadorVentana.cs:128-132` (`horasEfectivasCfg`) |

---

## Ronda 186: F978-F979 aplicadas — unidad resuelta solo desde `aa` y capacidad subestimada por día

### Evaluación Crítica de las Fases

#### 1. Fase 980 — Unidad 4 h cuando `aa == null` → F978 inoperante en primera corrida
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La resolución de `unidadAsig` dependía exclusivamente del mapeo preasignado `aa` (`AulaAsignaturasPorId`). En una primera corrida o cuando una asignatura no posee aula previamente asociada (`aa == null`), caía en fallback rígido de 4.0h. En posgrado (64h), esto volvía a calcular 16 sesiones requeridas en lugar de 8, reactivando los bloqueos falsos de capacidad ($16 > 12$) antes de evaluar las configuraciones reales de aula (`configs`).
  - *Causa Raíz:* Resolución prematura de unidad antes de identificar el aula gobernante/ganadora de `configs`.
  - *Impacto Confirmado:* Bloqueo falso recurrente de pre-vuelo en primeras corridas para programas de posgrado.
  - *Evidencia:* `AgValidadorVentana.cs:97-104`.
- **Acción Aplicada:** Extraído el helper `UnidadDe(AgAula au)` y diferido el cálculo de `unidadAsig` y `sesionesRequeridas` para ejecutarse tras determinar `mejorCfg` (aula ganadora de `configs` o fallback a `aa`), garantizando consistencia entre la unidad de demanda y la oferta de capacidad.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorVentana.cs`.

#### 2. Fase 981 — Capacidad cuenta 1 sesión por día ignorando AM/PM
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* `CapacidadDe` contabilizaba 1 sesión por cada ocurrencia de día K en la ventana semestral. En aulas de jornada continua (`HorasDia = 8`) con sesiones de 4h (pregrado con franjas mañana y tarde, reglas D2/D3), un solo día aloja 2 sesiones de 4h. Ignorar este factor subestimaba la capacidad real a la mitad, provocando que asignaturas válidas fueran bloqueadas por aparente falta de sesiones en la ventana.
  - *Causa Raíz:* Conteo de capacidad en ocurrencias brutas de calendario en vez de unidades de sesión por día (`au.HorasDia / unidadAu`).
  - *Impacto Confirmado:* Falsos bloqueos de capacidad en asignaturas de pregrado con aulas de 8h/día.
  - *Evidencia:* `AgValidadorVentana.cs:106-108`.
- **Acción Aplicada:** Incorporado el factor `sesionesPorDia = au != null ? Math.Max(1, (int)Math.Round(au.HorasDia / unidadAu)) : 1` dentro de `CapacidadDe`, multiplicando las ocurrencias del día K por la cantidad de sesiones discretas que admite la jornada del aula.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorVentana.cs`.

---

### Verificación de Principios (Ronda 186)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F980 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgValidadorVentana.cs:121-125` (`AulaGanadora`, `UnidadDe`) |
| F981 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgValidadorVentana.cs:105-111` (`sesionesPorDia = au.HorasDia / unidadAu`) |

---

## Ronda 187: F980-F981 aplicadas — capacidad por día sobre-permisiva (ignora franjas reales)

### Evaluación Crítica de las Fases

#### 1. Fase 982 — `sesionesPorDia` apuesta 2 sesiones/día sin verificar franjas AM/PM existentes
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* `sesionesPorDia = round(au.HorasDia / unidadAu)` asumía disponibilidad incondicional de franjas matutinas y vespertinas. Si un aula con `HorasDia = 8` operaba en un catálogo donde solo existía una franja de 4h activa (ej. solo MAÑANA), se le atribuían 2 sesiones/día ficticias, permitiendo que asignaturas sobrepasaran la capacidad real del catálogo de bloques e introduciendo conflictos diferidos durante la ejecución del algoritmo genético.
  - *Causa Raíz:* Cálculo de concurrencia diaria sin cruzar con el inventario real de franjas activas en `contexto.Bloques`.
  - *Impacto Confirmado:* Sobreestimación ficticia de capacidad diaria y aplazamiento de conflictos de asignación horaria.
  - *Evidencia:* `AgValidadorVentana.cs:110-113`.
- **Acción Aplicada:** Limitado `sesionesPorDia` mediante `Math.Min(..., Math.Max(1, franjasReales))`, donde `franjasReales` cuenta las franjas activas y distintas de duración `unidadAu` presentes en `contexto.Bloques`.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorVentana.cs`.

---

### Verificación de Principios (Ronda 187)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F982 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgValidadorVentana.cs:110-123` (`Math.Min(au.HorasDia / unidadAu, franjasReales)`) |

---

## Ronda 188: F982 aplicada — carriles no se generan y patrón semanal inflado (4 h fijo en `Construir`)

### Evaluación Crítica de las Fases

#### 1. Fase 983 — `Construir` divide por 4.0 fijo → posgrado 8 h declara patrón infactible
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* `AgPlanificadorPatronSemanal.Construir` calculaba $S = \lceil totalHoras / 4.0 \rceil$ de forma invariable. En programas de posgrado (asignatura de 64h con sesiones de 8h), proyectaba $S = 16$ sesiones en lugar de 8. Con un día por semana (`diasK = [1]`) en una ventana de 12 semanas, la exigencia de 16 semanas mínimas superaba la ventana ($16 > 12$), retornando `Failure`, degradando la asignación a un día forzado sin patrón (`patron = null`) y descalabrando la coherencia y reparto I1.
  - *Causa Raíz:* Divisor rígido de 4.0h en `Construir` sin parametrizar la unidad real de sesión del programa/aula.
  - *Impacto Confirmado:* Falsos rechazos de factibilidad de patrón semanal y pérdida de carriles multi-día en posgrado.
  - *Evidencia:* `AgPlanificadorPatronSemanal.cs:110` e invocación en `AgFabricaIndividuos.cs:314`.
- **Acción Aplicada:** Agregado parámetro opcional `double unidadSesion = 4.0` a `Construir`, calculando `S = (int)Math.Ceiling(totalHoras / unidad)` con `unidad = unidadSesion > 0 ? unidadSesion : 4.0`. En `AgFabricaIndividuos.cs`, resuelta `unidadPatron` desde el aula/nivel y suministrada a `Construir`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorPatronSemanal.cs` y `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`.

#### 2. Fase 984 — Sin overload `DiasKDe(diasPatron)` → carriles `[M,X]`/`[J,V]` inalcanzables
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La única variante de `DiasKDe(aa, aula)` dependía de una instancia `AgAulaAsignatura` persistida. Ante patrones textuales declarados o semillas de UI ('1,3' / '4,5'), no existía mecanismo puro para resolver e intersectar los días K contra la ventana del aula, impidiendo la siembra y expresión del modelo de carriles paralelos / oleadas T5-T7.
  - *Causa Raíz:* Carencia de overload funcional desacoplado de la entidad `AgAulaAsignatura`.
  - *Impacto Confirmado:* Imposibilidad de construir y resolver carriles multi-día alternos en la fábrica.
  - *Evidencia:* `AgPlanificadorPatronSemanal.cs:27-53`.
- **Acción Aplicada:** Implementado overload puro `DiasKDe(string diasPatron, string diasVentana = null, AgAula aula = null)` que parsea y valida días ISO [1..7], intersecta contra la ventana y recurre defensivamente al comportamiento canónico si no hay días válidos.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorPatronSemanal.cs`.

---

### Verificación de Principios (Ronda 188)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F983 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorPatronSemanal.cs:100-108`, `AgFabricaIndividuos.cs:314-319` (`unidadPatron`) |
| F984 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorPatronSemanal.cs:56-85` (`DiasKDe(string, string, AgAula)`) |

---

## Ronda 189: F983-F984 aplicadas — cruce degrada posgrado y fusiona carriles

### Evaluación Crítica de las Fases

#### 1. Fase 987 — Cruce reemplaza `dst.DiasPatron` con el patrón de otra aula → fusión de carriles
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* `diasK` se derivaba invariablemente de `aaSrc` (aula del progenitor) y luego sobreescribía `dst.DiasPatron = string.Join(",", diasK)`. Si el hijo pertenecía a un carril `[4,5]` y el progenitor a `[1,3]`, el hijo migraba forzosamente de carril, fusionando carriles concurrentes en uno solo y dejando el segundo sin sesiones.
  - *Causa Raíz:* El cruce no verificaba si el descendiente ya poseía un patrón de días asignado (`dst.DiasPatron`).
  - *Impacto Confirmado:* Fusión y colapso de carriles paralelos en oleadas T5-T7.
  - *Evidencia:* `AgOperadorCruce.cs:161-166`.
- **Acción Aplicada:** Declarada función local `AulaSrcDst(idAula)` e inicializado `diasK` priorizando `dst.DiasPatron` mediante `AgPlanificadorPatronSemanal.DiasKDe(dst.DiasPatron, null, AulaSrcDst(dst.IdAula))`; solo hereda de `aaSrc` si `dst.DiasPatron` está vacío.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgOperadorCruce.cs`.

#### 2. Fase 985 — Cruce reconstruye patrón con `unidad=4.0` → posgrado 8 h queda sin `PatronSemanal`
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* `_planificadorPatron.Construir` se invocaba sin el parámetro opcional `unidadSesion`, tomando por defecto `4.0`. En asignaturas de posgrado de 64h requiriendo 8 sesiones de 8h, `Construir` calculaba $S = \lceil 64 / 4 \rceil = 16$; en ventanas cortas retornaba `Failure`, dejando `dst.PatronSemanal = string.Empty` y desviando `SesionesDelGen()` al cálculo genérico.
  - *Causa Raíz:* Regla D4′ no propagada al operador de cruce.
  - *Impacto Confirmado:* Desconfiguración del histograma de sesiones semanales, subestimación de carga e inconsistencias en la compuerta I1.
  - *Evidencia:* `AgOperadorCruce.cs:199-208`.
- **Acción Aplicada:** Obtenida `unidadCruce` resolviendo `AgUnidadesSemanales.HorasSesionAula(aulaDst, nivel)` desde el aula y modalidad del programa destino, suministrándola como 6º argumento a `_planificadorPatron.Construir`.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgOperadorCruce.cs`.

#### 3. Fase 986 — Selección de bloque fija `DuracionHoras == 4.0` → posgrado recibe bloque de 4 h
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La heurística de selección de bloque (`bloquePatron`) imponía `b.DuracionHoras == 4.0` fijo. En posgrado, donde los bloques son de 8h, nunca había coincidencia y el fallback elegía cualquier bloque del día, asignando bloques de 4h a posgrados y reconfigurando `HorasSemanales = 4.0`.
  - *Causa Raíz:* Duración hardcodeada a 4.0 en lugar de utilizar la unidad efectiva del aula.
  - *Impacto Confirmado:* Violación silenciosa de la regla D4′ al generar 16 sesiones de 4h en lugar de 8 sesiones de 8h.
  - *Evidencia:* `AgOperadorCruce.cs:170-173`.
- **Acción Aplicada:** Modificado `bloquePatron` para buscar prioritariamente bloques con `Math.Abs(b.DuracionHoras - unidadCruce) < 0.001`, manteniendo fallback secundario en 4.0 y fallback terciario de cualquier bloque del día.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgOperadorCruce.cs`.

---

### Verificación de Principios (Ronda 189)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F987 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgOperadorCruce.cs:164-172` (`diasK` condicionado a `dst.DiasPatron`) |
| F985 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgOperadorCruce.cs:175-179,214-220` (`unidadCruce` en `Construir`) |
| F986 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgOperadorCruce.cs:181-189` (`bloquePatron` con `unidadCruce`) |

---

## Ronda 190: F985-F987 aplicadas — R8 y columna K incompatibles con carriles/8 h

### Evaluación Crítica de las Fases

#### 1. Fase 988 — `R8_DiaFueraDePatron` usa `gr.First().DiasPatron` → todo carril secundario es violación dura
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La regla R8 comparaba todos los días de las asignaciones del grupo contra el patrón del primer gen (`gr.First().DiasPatron`). En cohorte con carriles alternos (ej. carril 1 `[1,3]` y carril 2 `[4,5]`), los días 4 y 5 se marcaban fuera de patrón con `esDura = true`, sentenciando al individuo con `hardCount > 0` e imposibilitando la supervivencia del segundo carril en la selección genética.
  - *Causa Raíz:* Validación mono-carril a nivel de grupo agregador en lugar de evaluar a nivel de gen/carril individual.
  - *Impacto Confirmado:* Bloqueo total de la supervivencia de oleadas y carriles paralelos T5-T7.
  - *Evidencia:* `AgValidadorReglas.cs:418-433`.
- **Acción Aplicada:** Modificado `AgValidadorReglas.cs` para validar individualmente cada gen con `DiasPatron` explícito (`gr.Where(g => !string.IsNullOrWhiteSpace(g.DiasPatron))`), verificando si su propio `DiaSemana` pertenece a los días permitidos de su propio carril.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorReglas.cs`.

#### 2. Fase 989 — Columna K cae a 4 h cuando falta el bloque → posgrado reportado mal
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* En `AgMaterializadorSesiones.cs`, si `idBloque <= 0` o el bloque no se encontraba en `contexto.BloquesPorId`, `horas` tomaba el valor literal `4.0`, generando cadenas `"DIA|4|"` en la columna K aun en programas de posgrado con sesiones de 8h, propagando horas distorsionadas a la vista y a `RebuildSelectedDays`.
  - *Causa Raíz:* Fallback escalar rígido en 4.0h en lugar de consultar la duración asignada en `HorasSemanales` del gen.
  - *Impacto Confirmado:* Distorsión de horas reportadas en la vinculación académica de posgrado.
  - *Evidencia:* `AgMaterializadorSesiones.cs:122-129`.
- **Acción Aplicada:** Reemplazado el literal `4.0` por `subGrp.Select(x => x.HorasSemanales).Where(h => h > 0).DefaultIfEmpty(4.0).Max()`, preservando la duración del gen si no hay bloque resoluble.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgMaterializadorSesiones.cs`.

---

### Verificación de Principios (Ronda 190)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F988 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgValidadorReglas.cs:418-438` (`genFuera` por carril) |
| F989 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgMaterializadorSesiones.cs:123-125` (`HorasSemanales` fallback) |

---

## Ronda 191: F988-F989 aplicadas — columna K pierde franja PM y formatter canónico es inerte/incompatible (Fases 990 y 991)

