# Plan Consolidado de Corrección y Evolución — v1.2.49.0
## Basado en Auditoría Interna + Revisiones Externas (Kimi · GML5 · Qwen)

---

## Estado del Sistema

| Dimensión | Estado | Detalle |
|---|---|---|
| Arquitectura Vault Efímero | 🟢 Sólida | Patrón RAM→Cifrado→Disco correcto |
| Integridad de Datos en Guardado | 🔴 Crítico | 7 sesiones generadas, 6 persistidas |
| Coherencia de Vinculación | 🔴 Crítico | IdTercero cambia de 3→2 al recargar |
| `LeerBytesConRetryStatic` | 🔴 Crítico | Retorna `null` en vez de lanzar `IOException` |
| Temporales en Disco | 🟠 Alto | 15+ archivos por sesión de uso |
| Tiempo de Apertura | 🟠 Alto | ~9s por 5 contextos desencriptados independientes |
| DDL Diferido | 🟡 Medio | `ALTER TABLE` puede fallar si flag se reactiva |
| Semilla de Catálogos | 🟡 Medio | Doble fuente de verdad (filesystem vs recurso embebido) |
| `SecureString` | 🟡 Medio | Contraseña permanece en heap del assembly |
| Bug `path1 null` | 🟢 Bajo | Controlado por retry, no impacta usuario |
| Documentación del Plan | 🟡 Medio | `implementation_plan.txt` desactualizado |

---

## Hallazgos Consolidados — Análisis Cruzado

### 🔴 BUG-1: Pérdida de sesión al guardar (7 → 6)

**Evidencia en log:**
```
[10:28:34] btnGenerar: Generación exitosa. Sesiones: 7. Total: 48h.
[10:28:36] Preparación UI completada en 8ms. Sesiones a procesar: 6
[10:28:42] GetSesionesByVinculacion(id=1): 6 sesiones encontradas.
```

**Causa raíz (consenso GML5 + análisis de código):** En `btnGenerar_Click`, el log reporta `dgvSesiones.Rows.Count` (que incluye la fila fantasma `IsNewRow` de WinForms). La séptima sesión sí fue generada y agregada a `_sesionesGeneradas`, pero en `PrepararSesionesDesdeGrid()` hay una condición de carrera: si la celda `ColChecked` de la última fila no ha tenido su `CommitEdit` implícito antes de que el hilo UI entregue control al `Task.Run`, la fila se filtra como no-checked.

**Fix:** Agregar `dgvSesiones.EndEdit()` como primera línea de `PrepararSesionesDesdeGrid()`.

---

### 🔴 BUG-2: Corrupción de IdTercero al recargar (3 → 2)

**Evidencia en log:**
```
[10:27:59] CargarEstadoPrevio: IdTercero=3, Sede=3, Asignatura=15
[10:28:42] CargarEstadoPrevio: IdTercero=2, Sede=6, Asignatura=56
```

**Causa raíz:** El usuario cambió `cmbTerceros` de Id=3 a Id=2 entre la generación y el guardado. `ValidarParametrosCriticos()` debería haber bloqueado el guardado porque el hash `_ultimoEstadoGenerado` incluye `T={idTercero}`. Sin embargo, la vinculación en `PrepararVinculacionDesdeForm()` toma el valor *actual* del combo (Id=2), mientras que las sesiones en el grid fueron generadas con el Tercero Id=3 y no tienen `row.Tag` asignado para recuperar el IdTercero original. Resultado: la vinculación queda con IdTercero=2 pero las sesiones tienen el nombre del Tercero=3 visible en pantalla.

**Fix (dos capas):**
1. En `PrepararSesionesDesdeGrid`: si `row.Tag == null`, lanzar `InvalidOperationException` en lugar de inferir del combo.
2. En `PrepararVinculacionDesdeForm`: verificar que `IdTercero` del combo coincide con el `IdTercero` que está en los `row.Tag` de las sesiones antes de proceder.

---

### 🔴 BUG-3: `LeerBytesConRetryStatic` retorna `null` silenciosamente

**Código actual:**
```csharp
while (intentos < maxIntentos) { ... }
return null; // ← Viola el walkthrough
```

**Walkthrough dice:** "...la aplicación lanza una `IOException` explícita en lugar de fallar silenciosamente."

**Impacto:** `GuardarAsync` ya maneja el null con su propio `throw`, pero cualquier otro consumidor futuro del método recibirá un `null` y fallará con `NullReferenceException` en el cifrado.

**Fix:** Reemplazar `return null` final por `throw new IOException("No se pudo leer el archivo temporal tras agotar reintentos.")`.

---

### 🟠 PROBLEMA-4: Explosión de archivos temporales (15+ por sesión)

**Evidencia en log:** Cada `InvokeRepository`, `UpdateAsignaturas`, y `CargarEstadoPrevio` crea un nuevo `EncryptedSQLiteContext` con un GUID único → desencripta el vault completo → crea archivo temporal. En la sesión auditada se identifican al menos 6 archivos distintos en 90 segundos de uso.

**Causa raíz:** El diseño actual de `InvokeRepository` instancia y destruye un contexto por cada operación, sin reutilización durante la vida del formulario.

**Fix:** Implementar un contexto compartido de lectura durante toda la sesión del formulario (ver Fase B).

---

### 🟠 PROBLEMA-5: Tiempo de apertura excesivo (~9 segundos)

**Evidencia en log:**
```
[10:27:51] Iniciando SesionesAcademicas
[10:28:00] LoadInitialData: Carga completada con éxito
```

**5 desencriptados independientes en `OnShown`:** `ExecuteSeedData` + `CatalogCacheService.EnsureLoaded` + `UpdateAsignaturas` (×2 por eventos de combo) + `CargarEstadoPrevio`.

**Fix:** Consolidar en un único contexto de lectura compartido al abrir el formulario.

---

### 🟡 PROBLEMA-6: DDL diferido no es idempotente

En `GuardarAsync`, el `ALTER TABLE` se ejecuta si `_esquemaRequiereMigracion == true`. Si el flag se reactiva (condición de carrera, reinicio parcial), el `ALTER TABLE` fallará porque la columna ya existe. `ExecuteRawSql` no tiene manejo de "columna ya existente".

**Fix:** Envolver el DDL en verificación `PRAGMA table_info` antes de ejecutar, o capturar la excepción específica de "duplicate column".

---

### 🟡 PROBLEMA-7: Doble fuente de verdad para datos de semilla

`ExecuteSeedData` lee del filesystem (`AcademicSeedData.sql` en disco), mientras que `SeedCatalogosTransaccional` usa el recurso embebido. Si el archivo en disco diverge del embebido, los catálogos en memoria difieren de los que se persisten al Vault.

**Fix:** `ExecuteSeedData` debe llamar a `SeedCatalogosTransaccional` directamente, eliminando la lectura del filesystem.

---

### 🟡 PROBLEMA-8: Comentarios obsoletos en `AcademicRepository.cs`

```csharp
/// Los métodos de escritura deben llamar context.NotifyChanges() para re-encriptar al cerrar.
```

`NotifyChanges()` fue deprecado. El comentario del patrón de uso es incorrecto y confunde a futuros desarrolladores.

---

### 🟡 PROBLEMA-9: `ValidarParametrosCriticos` tiene semántica invertida

Retorna `true` cuando hay error, lo que produce código contraintuitivo:
```csharp
if (ValidarParametrosCriticos()) return; // "if hay error, salir"
```

---

### 🟡 PROBLEMA-10: `GetAwaiter().GetResult()` en `InitializeDatabase`

```csharp
byte[] initialBytes = EncryptedSQLiteContext
    .LeerBytesConRetryStatic(context.TempDbPath)
    .GetAwaiter().GetResult(); // Bloqueo síncrono de async
```

En .NET 4.8 dentro de un AppDomain COM (Excel-DNA), esto puede causar deadlock si el contexto de sincronización está capturado.

---

### 🟡 PROBLEMA-11: Logger duplicado en `CreateConnectionWithRetry`

```csharp
LoggerHelper.Warning($"... Estabilizando AppDomain...");
LoggerHelper.Warning($"...");  // ← Duplicado exacto sin el sufijo descriptivo
```

---

### 🟢 PROBLEMA-12: `GetAsignaturasByPrograma` filtra en memoria

La query trae **todas** las asignaturas del programa (56 en el log) y filtra por semestre en C#. Ineficiente para programas con muchas asignaturas.

---

### 🟢 PROBLEMA-13: `implementation_plan.txt` desactualizado

Las fases 1, 2 y 3 aparecen sin marcar (`[ ]`) pero ya están implementadas en el código desplegado. Genera confusión sobre el estado real del proyecto.

---

## Plan de Acción v1.2.49.0

### FASE A — Corrección de Bugs Críticos de Datos [P0 · Estimado: 3h]

> ⚠️ **No desplegar a producción sin completar esta fase.** Los bugs A.1, A.2 y A.3 causan corrupción silenciosa de registros académicos.

| ID | Tarea | Archivo | Cambio concreto |
|---|---|---|---|
| **A.1** | Fix pérdida de sesión | `FormSesionesAcademicas.cs` | Primera línea de `PrepararSesionesDesdeGrid()`: `dgvSesiones.EndEdit();` |
| **A.2** | Fix corrupción de IdTercero | `FormSesionesAcademicas.cs` | En `PrepararSesionesDesdeGrid`: si `row.Tag == null`, lanzar `InvalidOperationException("Sesión sin modelo asociado. Regenerar.")`. Eliminar el fallback al combo. |
| **A.3** | Validación cruzada Tercero-Sesiones | `FormSesionesAcademicas.cs` | En `PrepararVinculacionDesdeForm`: antes de construir el objeto, verificar que el `IdTercero` del combo coincide con `((Sesion)row.Tag).IdTercero` de todas las filas. Si diverge, abortar con mensaje claro. |
| **A.4** | Fix `LeerBytesConRetryStatic` | `EncryptedSQLiteContext.cs` | Reemplazar `return null;` final por `throw new IOException("No se pudo leer el archivo temporal tras agotar reintentos.");` |
| **A.5** | Log de auditoría de guardado | `FormSesionesAcademicas.cs` | En `GuardarAsync`, loguear explícitamente el objeto `Vinculacion` (IdTercero, IdSede, IdAsignatura, TotalSesiones) justo antes del `Commit()`. |

---

### FASE B — Optimización de Vault e I/O [P1 · Estimado: 4h]

| ID | Tarea | Archivo | Cambio concreto |
|---|---|---|---|
| **B.1** | Contexto único de lectura en formulario | `FormSesionesAcademicas.cs` | Refactorizar `OnShown` para abrir **un solo** `EncryptedSQLiteContext` que pase la conexión a `ExecuteSeedData`, `CatalogCacheService.EnsureLoaded`, `UpdateAsignaturas` y `CargarEstadoPrevio`. Objetivo: 1 desencriptado en apertura en vez de 5. |
| **B.2** | Limpieza proactiva de temporales | `FormSesionesAcademicas.cs` | Invocar `EncryptedSQLiteContext.CleanupOrphanedTempFiles()` en `OnShown` (antes de abrir el contexto compartido) y en `OnFormClosing`. |
| **B.3** | Timeout en Semaphore | `EncryptedSQLiteContext.cs` | Cambiar `_lock.Wait()` por `_lock.Wait(TimeSpan.FromSeconds(30))` con manejo explícito del timeout para evitar bloqueo indefinido. |
| **B.4** | DDL idempotente | `FormSesionesAcademicas.cs` | Antes de ejecutar el `ALTER TABLE`, verificar con `PRAGMA table_info(vinculaciones)` si la columna ya existe. Ejecutar solo si no existe. |
| **B.5** | Unificar semilla al recurso embebido | `FormSesionesAcademicas.cs` | `ExecuteSeedData` debe delegar en `SeedCatalogosTransaccional()` del repositorio. Eliminar la lectura de `AcademicSeedData.sql` del filesystem. |

---

### FASE C — Completar Deuda Técnica del Plan v1.2.48.0 [P1 · Estimado: 2h]

| ID | Tarea | Archivo | Cambio concreto |
|---|---|---|---|
| **C.1** | Eliminar logger duplicado | `EncryptedSQLiteContext.cs` | Remover la segunda llamada a `LoggerHelper.Warning` en `CreateConnectionWithRetry`. |
| **C.2** | Eliminar `NotifyChanges()` | `AcademicRepository.cs` | Actualizar el comentario de clase (eliminar referencia a `NotifyChanges()`). Verificar que no existan llamadas residuales en ningún archivo. |
| **C.3** | Renombrar `ValidarParametrosCriticos` | `FormSesionesAcademicas.cs` | → `ParametrosEstanDesincronizados()`. Actualizar todas las referencias. |
| **C.4** | Fix `.GetAwaiter().GetResult()` | `SqliteHelper.cs` | Crear método sincrónico `LeerBytesConRetrySynchronous(string path)` que use `Thread.Sleep` en vez de `Task.Delay`. Usarlo en `InitializeDatabase` en lugar del bloqueo async. |
| **C.5** | Actualizar `implementation_plan.txt` | `implementation_plan.txt` | Marcar con `[x]` las tareas ya implementadas. Archivar el documento en `/docs/archive/` y crear `PLAN_v1.2.49.0.md` como fuente de verdad activa. |

---

### FASE D — Robustez y Calidad [P2 · Estimado: 3h]

| ID | Tarea | Archivo | Cambio concreto |
|---|---|---|---|
| **D.1** | Filtrar asignaturas en SQL | `AcademicRepository.cs` | Agregar parámetro `semestre` a `GetAsignaturasByPrograma` y filtrar en la query: `WHERE idprograma = @idPrograma AND semestre = @semestre`. Eliminar el filtro en memoria de `UpdateAsignaturas`. |
| **D.2** | Validación de fechas en generación | `FormSesionesAcademicas.cs` | En `btnGenerar_Click`: si `dtpFinal.Value < dtpInicio.Value`, mostrar error y retornar antes de iniciar el bucle. |
| **D.3** | Cultura invariante en fechas | `FormSesionesAcademicas.cs` | Cambiar `null` por `CultureInfo.InvariantCulture` en todos los `DateTime.TryParseExact` y `ToString("dd/MM/yyyy")` de `PrepararSesionesDesdeGrid`. |
| **D.4** | Verificación post-Vault | `SqliteHelper.cs` | Tras `UpdateVaultDirectly`, verificar que el archivo `.db.enc` resultante tiene tamaño > 0 y no difiere dramáticamente del tamaño anterior. Si falla, alertar y conservar el `.bak`. |
| **D.5** | Revisar `SecureString` | `SqliteHelper.cs` | Agregar `using` para el `SecureString` en `GetSecurePassword()`. Evaluar si la contraseña puede derivarse de un identificador del equipo en lugar de string embebido. |

---

### FASE E — Verificación y Cierre [Estimado: 1.5h]

| ID | Test | Criterio de Aceptación |
|---|---|---|
| **E.1** | Smoke Test limpio | Vault inexistente → abrir formulario → catálogos cargan sin error en < 2s |
| **E.2** | Consistencia de sesiones | Generar N sesiones → guardar → reabrir → verificar que log muestra "Sesiones a procesar: N" y se recargan exactamente N sesiones con los mismos metadatos |
| **E.3** | Bloqueo por desincronización | Generar con Tercero=3 → cambiar combo a Tercero=2 → intentar guardar → debe bloquearse con mensaje claro |
| **E.4** | Rendimiento de apertura | Apertura del formulario < 2 segundos en máquina estándar (reducción de 5 desencriptados a 1) |
| **E.5** | Stress de I/O | Mientras se guarda, bloquear archivo temporal con otro proceso → debe reintentar y completar (no NullRef) |
| **E.6** | Conteo de temporales | Sesión típica (abrir, cambiar combos 3 veces, guardar) → no más de 3 archivos temporales creados |
| **E.7** | Bug `path1 null` | Verificar que el retry sigue manejándose silenciosamente sin propagarse al usuario |

---

## Métricas de Éxito

| Métrica | Baseline Actual | Objetivo |
|---|---|---|
| Sesiones generadas = Sesiones persistidas | ❌ 7 generadas / 6 guardadas | ✅ N = N siempre |
| IdTercero consistente pre/post guardado | ❌ 3 → 2 | ✅ Sin cambio |
| Archivos temporales por sesión típica | ~15+ | ≤ 3 |
| Tiempo de apertura del formulario | ~9s | < 2s |
| Tiempo total de guardado | 3249ms | < 4000ms (mantener) |
| Excepciones no controladas al usuario | 0 (actual) | 0 (mantener) |

---

## Roadmap de Largo Plazo (Post v1.2.49.0)

Las siguientes mejoras son deseables pero no bloquean el despliegue. Se priorizan en sprints futuros:

**Escalabilidad (1 mes):** Sistema de migración versionada con `PRAGMA user_version` y scripts `V1__*.sql`. Actualmente solo se detecta la columna `idtercero` hardcodeada.

**Modernización (2-3 meses):** Migrar de `System.Data.SQLite` a `Microsoft.Data.Sqlite` para eliminar las dependencias COM y el bug `path1 null` de raíz. Introducir inyección de dependencias para `EncryptedSQLiteContext` e `IAcademicRepository`.

**Desacoplamiento UI (3+ meses):** Extraer la lógica de generación, validación y mapeo de sesiones a una clase de presentación (`SesionViewModel`). El formulario WinForms quedaría reducido a bindings y eventos.

**Observabilidad (continuo):** Logging estructurado con métricas pre/post cifrado (`plainDatabaseBytes.Length`, `newEncryptedBlob.Length`, `sw.ElapsedMilliseconds`) para detectar degradaciones en equipos con antivirus agresivos.

---

## Resumen de Prioridades

```
INMEDIATO (esta sesión):
  A.1 + A.2 + A.3 + A.4 → Bugs de integridad de datos — riesgo de corrupción en producción

ESTA SEMANA:
  A.5 + B.1..B.5 + C.1..C.5 → Deuda técnica + optimización de I/O

PRÓXIMA SEMANA:
  D.1..D.5 + E.1..E.7 → Calidad, robustez y verificación completa

FUTURO:
  Roadmap largo plazo → escalabilidad y modernización arquitectónica
```
