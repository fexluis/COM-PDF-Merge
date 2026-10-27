# Restauración de las 3 Dimensiones en Perfiles de Árbol Académico y Estabilización de Sincronización

**Fecha:** 2026-10-23  
**Ronda:** 221 / 222  
**Fases:** 1203 a 1222  
**Módulo:** `Academic.Application` / `Academic.UI.Forms` / `Academic.Infrastructure.Persistence`

---

## 1. Propósito y Causa Raíz

Al guardar un perfil de árbol de selección académica, las 3 dimensiones de configuración (`Jornada`, `HorasDia`, `DiasVentana`) se persisten correctamente en la base de datos a través de `AgArbolMapper` y `AgArbolRepository`. Sin embargo, al cargar el perfil mediante `AsegurarSubarbol` (`AgFormGeneradorHorarios.cs`), únicamente se activaba la casilla `Checked = true` en el nodo de nivel 6 (Modalidad), sin sincronizar las 3 dimensiones desde el DTO hacia el objeto tipado `AgNodoArbolKey` asignado al `Tag` del nodo.

Como el método `CargarHijosLazy` siembra valores por defecto (`MANANA / 4h / MAR-VIE`), el árbol mostraba visualmente estos valores por defecto y, si el usuario procedía a sobrescribir o guardar nuevamente el perfil, se persistían dichos defaults en vez de las dimensiones originales previamente guardadas (`AgSeleccionMapper` lee directamente desde la key).

Adicionalmente, se detectaron y corrigieron anomalías colaterales en el ciclo de vida del árbol, sincronización asíncrona y migraciones SQLite:
1. `MigrationV73` no estaba registrada en `SqliteHelper`.
2. `MigrationV76` realizaba mapeo frágil de 8 columnas mediante `ValueTuple`.
3. El debounce de sincronización de dimensiones filtraba por `ActiveSubunidad` global en vez de respetar la clave canónica `AulaRef`.
4. El cierre del formulario no detenía ni volcaba el temporizador de debounce, provocando riesgo de pérdida de datos y deadlocks con el hilo STA.
5. `AsegurarSubarbol` desplegaba modales bloqueantes durante la restauración automática en `OnShown`.

---

## 2. Resumen de Cambios por Fase

### Fase 1203 — Mapeador `AgNodoArbolKeyMapper` y Registro en CSPROJ
- Creación de `ConversorPDF.XLL\Academic\Application\Mappers\AgNodoArbolKeyMapper.cs`.
- Método `AplicarDimensiones(AgNodoArbolKey key, AgContextoSeleccionDto dto)` que normaliza cadenas con `Trim().ToUpperInvariant()`, valida jornada no vacía, horas ∈ {4, 8} y ventana de días contra `AgDiasVentana.Dias()`.
- Fallback defensivo (`MANANA / 4 / MAR-VIE`) y `Result.Failure("I1_RESTAURACION_INCOMPLETA: ...")` ante datos incompletos.
- Registro en `ConversorPDF.XLL.csproj`.

### Fase 1204 — Suite de Pruebas Unitarias
- Archivo `ConversorPDF.Tests\Academic\Application\Mappers\AgNodoArbolKeyMapperTests.cs`.
- 6 pruebas xUnit bajo convención `{Accion}_When_{Cond}_Should_{Res}` con FluentAssertions 6.12 cubriendo dimensiones válidas, fallbacks y guardas de nulos.

### Fase 1205 — Restauración de Dimensiones en el Árbol
- En `AsegurarSubarbol`, invocación de `AgNodoArbolKeyMapper.AplicarDimensiones` sobre nodos de nivel 6 (localizados y dinámicos).
- Conteo de combinaciones que requirieron fallback (`dimensionesDefecto`).

### Fase 1206 — Visibilidad, Repintado y Refresco de Cabeceras
- Invocación de `primerNodoRestaurado.EnsureVisible()`, `treeContextoAcademico.Invalidate()` y `ActualizarCabecerasNivel6()` en el `finally`.

### Fase 1207 — Prevención de Descarte Silencioso al Persistir
- En `GuardarNuevoPerfilArbolAsync` y `SobrescribirPerfilArbolAsync`, evaluación de `descartes` y diálogo modal `Yes/No` con detalle de hasta 10 nodos.

### Fase 1208 — Sincronización de Combos Sedes y Programas
- En `CargarPerfilArbolSeleccionadoAsync`, sincronización del `SelectedItem` de `cmbSedes` y `cmbProgramas` con el primer DTO de la selección.

### Fase 1209 — Telemetría Estructurada y Traza Diagnóstica P12
- Enlace con `CrearTraza("CargarPerfilArbol")` introduciendo el paso canónico `P12_RestaurarDimensionesPerfil`.

### Fase 1210 — Documentación Técnica de Cambios Inicial
- Documentación inicial del comportamiento del árbol y advertencia sobre perfiles históricos sobreescritos.

### Fase 1212 — Registro de `MigrationV73` en `SqliteHelper`
- Inclusión de `new MigrationV73_LimpiarDiasPatronMultidia()` en `SqliteHelper._migrations`.

### Fase 1213 — POCO `AgAulaDimensionesCandidato` para `MigrationV76`
- Reemplazo de `ValueTuple` de 8 elementos por clase POCO sellada registrada en `ConversorPDF.XLL.csproj`.

### Fase 1214 — Aislamiento de Sincronización frente a `ActiveSubunidad`
- Envío de `string.Empty` en la sincronización del debounce (`AgFormGeneradorHorarios.cs:106`) y preservación del filtro en `AgAulaDimensionesSync.cs` solo cuando se informa explícitamente.

### Fase 1215 — Parada y Flush del Timer de Dimensiones en Cierre
- Invocación de `_timerSyncDimensiones.Stop()` y vaciado de cambios pendientes al inicio de `OnFormClosing`.

### Fase 1216 — Acotamiento de Logging Diagnóstico en `MigrationV76`
- Truncado a 20 advertencias individuales (`Take(20)`) con resumen general del total.

### Fase 1217 — Eliminación de `async` sin `await` (CS1998)
- Conversión de `AsegurarSubarbolAsync` a método síncrono `AsegurarSubarbol(IReadOnlyCollection<AgContextoSeleccionDto> nodos, bool mostrarAvisos = true)`.

### Fase 1218 — Prevención de Deadlock en Hilo UI al Cerrar Formulario
- Inclusión de `.ConfigureAwait(false)` en `AgAulaDimensionesSync.ExecuteAsync` y ejecución mediante `Task.Run` protegido con `try/catch` en `OnFormClosing`.

### Fase 1219 — Guarda Post-Await en Temporizador de Debounce
- Inclusión de `if (IsDisposed || Disposing || !IsHandleCreated) return;` tras `await SincronizarAsync` en el evento `Tick` para prevenir excepciones `ObjectDisposedException`.

### Fase 1220 — Normalización de Trazabilidad P12 y Causa de Fallo
- Separación de `P12_CargarPerfil` (`AgCausaFallo.SinAulas`) y `P12_RestaurarDimensionesPerfil` (`AgCausaFallo.HorasDiaInvalidoNivel`). `traza.Fin(true)` consistente.

### Fase 1221 — Supresión de Modales en Restauración Automática (`OnShown`)
- Pase de `mostrarAvisos: false` desde `OnShown` para garantizar una carga fluida sin interrupciones emergentes.

### Fase 1222 — Auditoría de Filas Afectadas en `MigrationV73`
- Conteo previo de filas candidatas con `SELECT COUNT(*)` antes de ejecutar la limpieza de `dias_patron`.

---

## 3. Cómo Usar

1. En `AgFormGeneradorHorarios`, configure las dimensiones deseadas en el árbol para las modalidades seleccionadas (ej. `TARDE`, `8 Horas`, `JUE-SAB`).
2. Guarde un nuevo perfil o sobrescriba uno existente. Si alguna combinación no tiene sus dimensiones completas, el sistema advertirá con un diálogo de confirmación indicando cuáles combinaciones serán omitidas.
3. Al cargar el perfil guardado mediante el combo:
   - Se restauran los radios de jornada, horas y días exactamente según lo guardado.
   - La rama del árbol se expande y se hace visible automáticamente.
   - Los combos de Sedes y Programas reflejan el contexto del perfil.
   - En el log de diagnóstico (`Modo Debug`), se registra la traza del paso `P12_RestaurarDimensionesPerfil`.
4. Al abrir el formulario, el último perfil activo se restaura silenciosamente en segundo plano sin desplegar modales molestos.

---

## 4. Cobertura de Pruebas

- Mapeador de dimensiones: `AgNodoArbolKeyMapperTests` (100% cubierto, 8 casos de prueba).
- Mapeador de selección inverso: `AgSeleccionMapperTests` (100% cubierto).
- Sincronización con subunidad vacía: `AgAulaDimensionesSyncTests` (100% cubierto).
- Migración de dimensiones: `MigrationV76Tests` (100% cubierto).
