<!-- historico congelado - generado por .agents/tools/split-history.ps1 - commit d0ee6fe - no editar -->
# implementation_plan - Fases 270-461 (Rondas 42, 43, 44)

> Indice: [../../../implementation_plan.md](../../../implementation_plan.md) - Escision: 2026-11-16

---
### 104. **Fase 270 — Persistencia Inmediata al Cierre del Formulario (`OnFormClosing`)**
- **Archivo Modificado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico y Causa Raíz:**
  La invocación de `MarkDirty()` programa un temporizador adaptativo de *debounce* de 5 segundos antes de re-encriptar el vault cifrado. Si el usuario guarda un perfil o genera aulas y cierra la ventana o Excel antes de que expire ese temporizador, o si `RibbonController` destruye la sesión compartida antes de completar el ciclo asíncrono, la base de datos temporal se cerraba sin completar la escritura en el archivo `.db.enc`.
- **Solución Implementada:**
  Siguiendo el patrón arquitectónico establecido en `FormVinculaciones.cs` (L427-450), se sobreescribió `OnFormClosing` en `AgFormGeneradorHorarios.cs`:
  ```csharp
  protected override void OnFormClosing(FormClosingEventArgs e)
  {
      try
      {
          var scheduler = SyncContext.Scheduler;
          if (scheduler != null)
          {
              var sharedCtx = SqliteHelper.GetContext();
              if (sharedCtx != null)
              {
                  try
                  {
                      scheduler.ForcePersistVaultAsync(sharedCtx).GetAwaiter().GetResult();
                  }
                  catch (Exception ex)
                  {
                      LoggerHelper.Error("[AgFormGeneradorHorarios] Error al persistir el Vault en el cierre", ex);
                  }
              }

              scheduler.Flush();
          }
      }
      catch (Exception ex)
      {
          LoggerHelper.Error("[AgFormGeneradorHorarios] Error en secuencia de cierre del formulario", ex);
      }

      base.OnFormClosing(e);
  }
  ```

### 105. **Fase 271 — Notificación de Mutaciones en `AgScheduleRepository`**
- **Archivo Modificado:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgScheduleRepository.cs`
- **Diagnóstico y Solución:**
  Se implementó el método auxiliar `NotifyMarkDirty()` y se vinculó tras cada operación de escritura exitosa en el repositorio genético:
  - `InsertAulaAsync`
  - `InsertAulaAsignaturasAsync`
  - `InsertCorridaAsync`
  - `UpdateCorridaAsync`
  - `SaveIndividuoAsync`
  - `SaveBitacoraAsync`
  - `GuardarSesionesStagingAsync`
  Esto asegura que el historial de corridas, individuos generados y staging se mantengan sincronizados con el vault cifrado de forma homogénea.

### 106. **Fase 272 — Protocolo de Verificación de Persistencia Post-Cierre (Sin Cambios de Código)**
- **Verificaciones Operativas:**
  1. Guardar perfil en la UI $\rightarrow$ confirmar en el log: `ForcePersistVaultAsync … persistido exitosamente`.
  2. Cerrar Excel por completo $\rightarrow$ reabrir $\rightarrow$ verificar que el `Perfil Árbol` aparezca en el combo y el árbol jerárquico se reconstruya automáticamente con marcas (Fase 269).
  3. Verificar mediante consulta SQLite que `SELECT COUNT(*) FROM ag_arbol` y `SELECT COUNT(*) FROM ag_aula` mantengan sus registros intactos en la nueva sesión.
  4. Confirmar que `SyncScheduler.Flush()` ejecute la sincronización de cambios sin omitir escrituras.

---

## 📦 Entrega de Release v1.6.3.28 (Instalador Oficial)

- **Versión Oficial:** `1.6.3.28`
- **Archivos de Versión Sincronizados:**
  - `ConversorPDF.XLL\Core\Helpers\App.cs` (`APP_VERSION = "1.6.3.28"`)
  - `ConversorPDF.XLL\ConversorPDF.XLL.csproj` (`<Version>1.6.3.28</Version>`)
  - `installer\InstaladorConversorPDF.wxs` (`Version="1.6.3.28"`)
- **Verificación de Pruebas Unitarias:**
  - Suite completa: **431 / 431 pruebas aprobadas (0 errores)** (`dotnet test ConversorPDF.Tests\ConversorPDF.Tests.csproj -c Release`).
- **Compilación de Add-In:**
  - `ConversorPDF.XLL` (Release|x64): **0 Errores**.
- **Paquete Instalador Generado:**
  - Archivo: `installer\InstaladorConversorPDFv1.6.3.28.msi`
  - Tamaño: **5,402,624 bytes**
  - Estado: Listo para distribución y smoke test en copia del vault.

---

## Ronda 42 — Soporte de Fallback en Contexto Genético y Diagnóstico Accionable de Catálogo

### 107. **Fase 274 — Fallback de Semestre en `ConstruirContextoAsync` (Desbloqueo de Pre-Vuelo)**
- **Archivo Modificado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico y Causa Raíz:**
  En la Ronda 42, el proceso de generación de aulas reportó éxito en persistencia (`creadas: 0 | preexistentes omitidas: 193`), pero el diagnóstico pre-vuelo abortó con:
  `[ERROR PRE-VUELO] El aula o el programa no tiene asignaturas dependientes activas.` (para la combinación `2|3|2|I|A|PRESENCIAL`).
  Al analizar el origen de datos, `ConstruirContextoAsync` filtraba estrictamente por `semestre = @Semestre` (`'I'`). Si en el catálogo de base de datos la tabla `asignaturas` almacenaba el semestre en formato arábigo (`'1'`), vacío o nulo, la consulta devolvía 0 registros. A diferencia de `GenerarEncabezadosAulasAsync` (Fase 266), `ConstruirContextoAsync` carecía de fallback.
- **Solución Aplicada:**
  Se agregó la consulta de respaldo para `ctx.Asignaturas` y `ctx.PoolDocentes`:
  Si la consulta filtrada por semestre devuelve 0 filas, se ejecuta inmediatamente un fallback acotado a `idprograma = @IdPrograma` y `subunidad = @Sub` (sin restringir semestre). Esto garantiza que el motor genético disponga del conjunto de asignaturas del programa incluso ante discrepancias de nomenclatura entre la rama del árbol y el pensum.

### 108. **Fase 275 — Diagnóstico Pre-Vuelo Accionable y Contador Explícito de Aulas sin Asignaturas**
- **Archivo Modificado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:**
  Los mensajes de error previos eran genéricos e impedían discernir si el bloqueo se debía a un filtro de subunidad, falta de pensum cargado o discrepancia de semestre. Asimismo, `GenerarEncabezadosAulasAsync` informaba "éxito" aun cuando las aulas quedaban sin asignaturas vinculadas.
- **Solución Aplicada:**
  1. **Pre-vuelo accionable (`EjecutarDiagnosticoPreVueloAsync` y `btnOptimizar_Click`):**
     Se sustituyó el mensaje genérico por un detalle contextual explícito:
     `[ERROR PRE-VUELO] Sin asignaturas para idprograma={_contextoActual.IdPrograma} semestre='{semestreRama}' subunidad='{_contextoActual.Subunidad}'. Verifique el catálogo 'asignaturas' (idprograma/semestre/subunidad) o cargue el pensum.`
     Permite al usuario/administrador identificar de inmediato el parámetro faltante.
  2. **Contador explícito `sinAsignaturas` en `GenerarEncabezadosAulasAsync`:**
     Se introdujo el acumulador `int sinAsignaturas` que se incrementa si tras la consulta principal y el fallback la lista `idsAsig` permanece vacía.
     El diálogo `MessageBox.Show` y el log en `txtLogAG` notifican con advertencia si alguna de las aulas quedó sin vinculaciones en `ag_aula_asignatura`.

### 109. **Fase 276 — Verificación Operativa del Catálogo y Vault (Sin Código)**
- **Protocolo de Comprobación:**
  1. `SELECT COUNT(*) FROM asignaturas WHERE idprograma=2;` $\rightarrow$ si devuelve 0, indica dato faltante en el catálogo del vault (se requiere cargar pensum académico).
  2. `SELECT DISTINCT semestre FROM asignaturas WHERE idprograma=2;` $\rightarrow$ verificar formato de nomenclatura (`'I'`, `'II'` vs `'1'`, `'2'` vs `''`).
  3. `SELECT COUNT(*) FROM asignaturas WHERE idprograma=2 AND (subunidad='VALLE' OR subunidad='' OR subunidad IS NULL);` $\rightarrow$ comprobar visibilidad de subunidad.
  4. `SELECT COUNT(*) FROM ag_aula_asignatura WHERE id_aula IN (SELECT id FROM ag_aula WHERE idprograma=2);` $\rightarrow$ confirmar vinculaciones activas tras reejecutar `Generar Aulas`.
  5. Confirmar que el diagnóstico pre-vuelo avance sin errores de bloqueo.

---

## Ronda 43 — Traza Diagnóstica Integral y Verificación de Esquema en Módulo S (Fases 277 - 283)

### 110. **Fase 277 — `AgDiagnosticoTraza` (Motor de Trazabilidad Estructurada)**
- **Archivo Creado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgDiagnosticoTraza.cs`
- **Diagnóstico y Motivación:**
  Previamente, los fallos y anomalías en el Módulo S (`AgFormGeneradorHorarios`) se manifestaban de forma tardía, después de transacciones abortadas o en forma de errores genéricos sin visibilidad sobre las sentencias SQL ejecutadas, parámetros enlazados, conteos de filas obtenidas ni tiempos relativos de ejecución.
- **Implementación:**
  - Clase sellada `AgDiagnosticoTraza` con `Stopwatch` para temporización relativa continua en milisegundos (`t=...ms`).
  - Secuenciación determinista de pasos: `[P01]`, `[P02]`... `[Pnn]`.
  - Métodos estructurados:
    - `Inicio(detalle)` / `Fin(ok, detalle)`: Delimitación de ciclos operativos con identificador de correlación (ej. `S-HHmmss-Op`).
    - `Paso(nombre, detalle)`: Avance de hito con emisión en nivel `DEBUG`.
    - `Dato(clave, valor)`: Emisión de variables de contexto y configuración.
    - `Sql(nombre, sql, parametros)`: Normalización de sentencias SQL (colapsando saltos de línea y espacios duplicados) y serialización segura de parámetros anónimos sin exponer PII.
    - `Conteo(etiqueta, obtenidos, minimo)`: Detección proactiva de consultas vacías emitiendo `DEBUG` si `obtenidos >= minimo` o `WARNING` con marca `ALERTA (<minimo)` si el resultado es insuficiente.
    - `Contexto(AgContexto c)`: Volcado exhaustivo de dimensiones de corrida (Ids, Subunidad, Ventana, Aulas, Bloques, Docentes, Asignaturas, Festivos, etc.).
    - `Causa(AgCausaFallo causa, detalle)`: Registro explícito del código de causa raíz.
    - `Fallo(paso, ex)`: Captura y emisión de excepciones con stack trace a través de `LoggerHelper.Error`.
  - Integración dual: Emite a `LoggerHelper` (`Debug`, `Info`, `Warning`, `Error`, `Success`) y sincroniza con el delegado UI (`txtLogAG`) cuando la traza está habilitada.

### 111. **Fase 278 — `AgEsquemaDiagnostico` (Validación Preventiva de Esquema SQLite)**
- **Archivo Creado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgEsquemaDiagnostico.cs`
- **Diagnóstico y Motivación:**
  Los errores de derivación de esquema en bases de datos locales o shadow copy (tales como `no such column: activo` o `no such column: idaula`) solían manifestarse durante la ejecución de comandos Dapper en tiempo de corrida.
- **Implementación:**
  - Clase estática `AgEsquemaDiagnostico.Verificar(IDbConnection db, AgDiagnosticoTraza t, string tabla, params string[] requeridas)`.
  - Ejecuta `PRAGMA table_info(tabla)` para obtener el conjunto canónico de columnas existentes en minúsculas e invariante a cultura.
  - Compara contra las columnas requeridas por el código. Si detecta columnas faltantes, emite advertencia explícita en la traza y en `LoggerHelper.Warning` con el listado exacto de campos ausentes antes de que se ejecute la lógica de negocio.

### 112. **Fase 279 — Instrumentación Paso a Paso P01 - P11 en `AgFormGeneradorHorarios`**
- **Archivo Modificado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Implementación:**
  Se integró la traza `AgDiagnosticoTraza` a lo largo de todo el ciclo de vida del Módulo S mediante los siguientes hitos secuenciales:
  - **P01 — `CargarCatalogosAsync`:** Verificación previa de esquema (`AgEsquemaDiagnostico`) en tablas clave (`periodos`, `programas`, `sedes`, `ag_arbol`, `ag_arbol_nodo`, `ag_aula`, `ag_aula_asignatura`, `asignaturas`, `terceros_asignaturas`, `ag_corrida`) y reporte de conteos de catálogos cargados.
  - **P02 — `ActualizarVentanaSemestralAsync`:** Volcado estructurado de parámetros temporales (período, fechas límite, días de duración, modo manual vs automático).
  - **P03 — `PoblarArbolContextoAcademico`:** Conteo de nodos jerárquicos (raíz, sedes, programas, semestres, grupos, modalidades) y hojas generadas con formato `AulaRef`.
  - **P04 — `GenerarEncabezadosAulasAsync`:** Traza de combinaciones seleccionadas, inspección de SQL de asignaturas, conteos de muestra, registro de activación de fallback y balance final (`insertadas`, `omitidas`, `sinAsignaturas`).
  - **P05 — `ConstruirContextoAsync`:** Inspección de cada consulta preparatoria (`ag_aula`, `ag_bloque_horario`, `ag_aula_asignatura`, `asignaturas`, `terceros_asignaturas`, `terceros`, `festivos`, `presenciales`) registrando sentencias SQL normalizadas, parámetros y conteos.
  - **P06 — `AgContexto.ConstruirIndices`:** Invocación de `traza.Contexto(ctx)` y reporte de cardinalidad de índices cruzados.
  - **P07 — `EjecutarDiagnosticoPreVueloAsync`:** Registro de evaluaciones con tipificación mediante `AgCausaFallo` ante bloqueos de ventana, aulas, asignaturas o pool docente.
  - **P08 — `IniciarOptimizacionAsync`:** Registro de inserción de corrida, parámetros del motor genético (población, generaciones, mutación, élite) y evolución generacional.
  - **P09 — `CargarModuloResultadosAsync`:** Traza de individuos generados y persistencia en staging.
  - **P10 — `PromoverAProduccionAsync`:** Traza de inserción y promoción transaccional a tablas operativas de vinculaciones y sesiones.
  - **P11 — `OnFormClosing`:** Trazabilidad de persistencia forzada en el vault (`ForcePersistVaultAsync`) y reporte de duración de sincronización.

### 113. **Fase 280 — Códigos de Causa Raíz `AgCausaFallo`**
- **Archivo Creado:** `ConversorPDF.XLL/Academic/Domain/Enums/AgCausaFallo.cs`
- **Implementación:**
  - Enumeración tipificada que mapea las causas exactas de rechazo o inconsistencia:
    `Ninguna`, `SinPeriodo`, `SinVentana`, `SinAulas`, `SinAnclajeAula`, `SinAsignaturas`, `SinAulaAsignaturas`, `SinPoolDocente`, `SinBloques`, `VentanaInfactible`, `EsquemaIncompleto`, `TransaccionAbortada`.
  - Cumple estrictamente la regla *Single Type Per File* al alojarse en su propio archivo dentro del espacio de nombres `ConversorPDF.XLL.Academic.Domain.Enums`.

### 114. **Fase 281 — Control de Verbosidad y Checkbox de Modo Debug en UI**
- **Archivos Modificados:**
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.Designer.cs`
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Implementación:**
  - Se agregó el control `chkModoDebug` en la barra superior de acciones (`pnlBotonesSim`) de `AgFormGeneradorHorarios`.
  - Permite al usuario activar/desactivar la emisión en tiempo real de trazas de nivel `DEBUG` en el visor `txtLogAG`.
  - Al instanciarse cada traza con `CrearTraza(operacion)`, la propiedad `Habilitado` se inicializa a partir de `chkModoDebug.Checked` (o configuración local). Los errores y advertencias continúan escribiéndose en el archivo de log físico sin penalizar la interfaz gráfica.

### 115. **Fase 282 — Registro en Solución y Pruebas Unitarias**
- **Archivos Modificados y Creados:**
  - `ConversorPDF.XLL/ConversorPDF.XLL.csproj`: Registrados `<Compile Include>` para `AgCausaFallo.cs`, `AgDiagnosticoTraza.cs` y `AgEsquemaDiagnostico.cs`.
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgDiagnosticoTrazaTests.cs`:
    - Prueba de secuenciación de pasos `P01..Pnn`.
    - Prueba de silencio en UI cuando `Habilitado = false`.
    - Prueba de activación de `ALERTA` en `Conteo()` al recibir valores bajo el umbral mínimo.
    - Prueba de normalización de cadenas SQL multilínea y serialización de parámetros.
    - Prueba de asignación y persistencia de `UltimaCausa`.
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgEsquemaDiagnosticoTests.cs`:
    - Prueba con SQLite in-memory verificando reporte exitoso cuando todas las columnas requeridas existen.
    - Prueba con SQLite in-memory verificando reporte exacto de columnas faltantes.
    - Prueba ante tablas inexistentes.

### 116. **Fase 283 — Documentación Técnica de Diagnóstico**
- **Archivos Modificados:**
  - `implementation_plan.md`: Detalle histórico de las Fases 277 a 283.
  - `docs/articles/changes/2026-ag-horarios-academicos.md`: Inclusión de la sección 8.15 que documenta la arquitectura de trazabilidad, catálogo de códigos de causa raíz y el flujo de pasos P01 a P11.

### 117. **Fase 284 — Incorporación de `traza.Causa(...)` y Propiedad `UltimaCausa`**
- **Archivo Modificado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgDiagnosticoTraza.cs`
- **Diagnóstico:**
  Durante la instrumentación de Módulo S en la Fase 279, se invocaba `traza.Causa(AgCausaFallo.SinAulaAsignaturas, "...")`, pero la clase `AgDiagnosticoTraza` sólo exponía la enumeración de causa dentro de la sobrecarga de `Fallo(paso, ex, causa)`. Esto generaba un error de compilación CS1061 por método inexistente.
- **Solución Aplicada:**
  - Se implementó el método público `public void Causa(AgCausaFallo causa, string detalle = null)` que registra la causa tipificada, actualiza `UltimaCausa` y emite una advertencia estructurada `WARNING` con formato `causa={causa}`.
  - Se expuso la propiedad de lectura `public AgCausaFallo UltimaCausa { get; private set; }` para trazabilidad e inspección en pruebas unitarias.

### 118. **Fase 285 — Corrección de Columnas Requeridas en Verificación de Esquema (Eliminación de Falsos Positivos)**
- **Archivo Modificado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:**
  La verificación preventiva de esquema introducida en la Fase 278 evaluaba:
  1. `sedes` con la columna `"nombre"`: En el esquema canónico (`AcademicSchema.sql` L116-128), la tabla `sedes` no posee columna `nombre` (sus columnas descriptivas son `ciudad`, `subsede`, `departamento`).
  2. `ag_corrida` con la columna `"id_aula"`: En la migración canónica (`MigrationV49_AddAgHorariosTables.cs` L403), la columna añadida para el anclaje de aula se denomina `"idaula"` (sin guion bajo).
  Esto producía avisos ruidosos `[ESQUEMA] FALTAN: nombre` y `FALTAN: id_aula` en cada inicialización de catálogos.
- **Solución Aplicada:**
  Se actualizaron las llamadas en `CargarCatalogosAsync`:
  - `AgEsquemaDiagnostico.Verificar(db, traza, "sedes", "id", "ciudad", "subsede", "activo");`
  - `AgEsquemaDiagnostico.Verificar(db, traza, "ag_corrida", "id", "idperiodo", "idprograma", "idaula", "estado");`
  Garantizando reporte de 0 columnas faltantes en el arranque.

### 119. **Fase 286 — Verificación Estática y Cobertura de Pruebas**
- **Archivos Modificados:**
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgDiagnosticoTrazaTests.cs`
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgEsquemaDiagnosticoTests.cs`
- **Resultados:**
  - Validación de compilación estática sin advertencias ni errores CS1061.
  - Concordancia total entre los nombres de columnas verificados por `AgEsquemaDiagnostico` y los esquemas canónicos `AcademicSchema.sql` y `MigrationV49`.
  - Pruebas unitarias configuradas para validar la numeración correlativa `[P01..Pnn]`, la tolerancia ante consultas vacías, el registro de causas y la verificación de esquema SQLite.
- **Resultados de Release v1.6.3.29:**
  - Pruebas unitarias ejecutadas: **440 superadas**, 0 fallidas (100% éxito).
  - Ensamblados generados en `Release|x64`: `ConversorPDF.XLL.dll`, `ConversorPDF.XLL-AddIn64.xll`.
  - Instalador generado: `installer\InstaladorConversorPDFv1.6.3.29.msi` (5,410,816 bytes).

---

## Ronda 44 — Carga de AulaAsignaturas, Resolución de Sede en Staging y Cobertura Docente (Fases 293 - 296)

### 120. **Fase 293 — Carga Relacional de `ctx.AulaAsignaturas` en `ConstruirContextoAsync`**
- **Archivo Modificado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico y Causa Raíz:**
  En `ConstruirContextoAsync`, `ctx.AulaAsignaturas` nunca se cargaba desde la base de datos, manteniéndose como lista vacía (`Count = 0`). Esto dejaba `ctx.AulaAsignaturasPorId` desprovisto de asignaturas vinculadas, distorsionaba el diagnóstico de factibilidad y generaba la alerta `AulaAsignaturas: 0 ALERTA (<1)`.
- **Solución Aplicada:**
  - Se agregó la consulta a `ag_aula_asignatura` para el aula anclada `ctx.IdAulaCorrida` previo a `ctx.ConstruirIndices()`, poblando `ctx.AulaAsignaturas` con columnas canónicas (`id_aula`, `id_asignatura`, `exclusivo`, `prioridad`, `horas_objetivo`, `orden_curricular`, `requerida`, `subunidad`).
  - Se añadieron compuertas defensivas en `EjecutarDiagnosticoPreVueloAsync` e `IniciarOptimizacionAsync` que bloquean de inmediato con `AgCausaFallo.SinAulaAsignaturas` si el aula anclada no posee asignaturas vinculadas.

### 121. **Fase 294 — Resolución de `IdSede` y Salvaguarda Fail-Fast en Staging**
- **Archivos Modificados:**
  - `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgContexto.cs`
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`
  - `ConversorPDF.XLL/Academic/Application/Mappers/AgGenToAgVinculacionMapper.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgMaterializadorSesiones.cs`
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgMaterializadorSesionesTests.cs`
- **Diagnóstico y Causa Raíz:**
  La tabla `asignaturas` no posee columna `idsede`. Al instanciar `AgGen`, `asig.IdSede` era siempre 0, lo que se propagaba a `AgVinculacion` y provocaba que las vinculaciones materializadas en staging tuvieran `idsede = 0`, comprometiendo la consistencia de datos en producción.
- **Solución Aplicada:**
  - Se incorporó la propiedad `public int IdSede { get; set; }` en `AgContexto` y se asigna desde `sedeSel.Id` en `ConstruirContextoAsync`.
  - En `AgFabricaIndividuos.cs`, `IdSede` se resuelve a partir del aula anclada (`aSede.IdSede`) o del contexto (`contexto.IdSede`).
  - En `AgGenToAgVinculacionMapper.cs` `MapGrupo`, se implementó fallback para resolver `IdSede` desde el aula o contexto si `g0.IdSede <= 0`.
  - En `AgMaterializadorSesiones.cs`, se implementó validación fail-fast: si tras la resolución `agVinc.IdSede <= 0`, la materialización aborta inmediatamente con `Result.Failure`.

### 122. **Fase 295 — Validación de Cobertura de Pool Docente y Prevención de R5**
- **Archivos Modificados:**
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorVentana.cs`
- **Diagnóstico:**
  Si un programa carece de docentes habilitados en `terceros_asignaturas` para sus asignaturas, `AgFabricaIndividuos` recurría a docentes generales no habilitados o dejaba `idDocente = 0`, provocando violaciones duras insalvables de la regla R5 (`AgReglaViolacion.R5_HabilitacionDocentePool`).
- **Solución Aplicada:**
  - En `EjecutarDiagnosticoPreVueloAsync` e `IniciarOptimizacionAsync`, se valida la cobertura del pool. Si no existe ningún docente habilitado ni en catálogo, se bloquea con `AgCausaFallo.SinPoolDocente`. Si existen docentes pero ninguno en pool, se advierte explícitamente en el log y se registra la causa en la traza.
  - En `AgFabricaIndividuos.cs`, se previene la asignación de `idDocente <= 0` arrojando `InvalidOperationException` preventiva.
  - En `AgValidadorVentana.cs`, la ausencia total de docentes en pool y catálogo general se categoriza como bloqueo crítico.

### 123. **Fase 296 — Métricas de Cobertura en Trazabilidad Diagnóstica**
- **Archivo Modificado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Solución Aplicada:**
  - En `ConstruirContextoAsync`, `EjecutarDiagnosticoPreVueloAsync` e `IniciarOptimizacionAsync`, se calculan y emiten métricas estructuradas:
    - `traza.Dato("CoberturaDocente", $"{asignaturasConPool}/{totalAsignaturas}");`
    - `traza.Dato("AulasSinAsignaturas", aulasSinAsig);`
  - Proporciona visibilidad accionable previa a la ejecución sobre la completitud del pensum y asignaciones docentes.

### 124. **Fase 297 — UI: Panel "Modo AG (Staging)" en `tabResultados` (Módulo R)**
- **Archivo Modificado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.Designer.cs`
- **Diagnóstico y Análisis:**
  En la pestaña Módulo R (`tabResultados`), los usuarios no contaban con acceso directo a la inspección y manipulación de datos en staging (`ag_vinculaciones` / `ag_sesiones`). El único acceso disponible a nivel de form residía en `btnAbrirModoAg` (`FormSesionesAcademicas`), sin acceso estructurado a Vinculaciones completas ni a la vista agregada de Detalle.
- **Solución Aplicada:**
  - Se incorporó el panel `pnlAgModo` acoplado en `Dock = DockStyle.Top` (altura 38px, fondo suave `#FFF8F0`, borde `#FixedSingle`) ubicado inmediatamente debajo de `pnlHeaderResultados` y encima del `splitContainerResultados`.
  - Se añadieron los botones de acción e indicador visual:
    - `btnAgVinculaciones` ("🔗 Vinculaciones (AG)", 170x28px, x=10).
    - `btnAgVerVinculaciones` ("📋 Detalle (AG)", 150x28px, x=190).
    - `btnAgSesiones` ("✏️ Sesión (AG)", 150x28px, x=350).
    - `lblAgModo` ("Modo AG · staging: ag_vinculaciones / ag_sesiones", x=520, fuente Segoe UI 9pt Bold, ForeColor OrangeRed).
  - Se añadieron las declaraciones de campos privados y la adecuada secuencia en `ResumeLayout(false)` / `PerformLayout()`.

### 125. **Fase 298 — Apertura en Modo AG (Wiring desde Módulo R) y Resincronización**
- **Archivo Modificado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico y Análisis:**
  - El grid `dgvHorariosGenerados` enlaza una proyección anónima y no `List<SesionListItem>`, por lo que hacer un cast directo `as List<SesionListItem>` resultaría en `null`.
  - Para abrir `FormVinculaciones` en Modo AG se requieren 5 dependencias instanciadas canónicamente con `AgModoDatos.Ag`.
  - Para abrir `FormVerVinculaciones` se requiere consultar de forma tipada `vw_ag_sesiones_detalle` mediante `VinculacionQueryRepository.GetSesionesAsync(...)`.
  - Para abrir `FormSesionesAcademicas` con la vinculación exacta, se debe resolver `id_vinculacion` desde la fila seleccionada en `ag_sesiones`.
- **Solución Aplicada:**
  - En `ConfigurarEventos()` se cablearon los tres eventos Click de forma segura y asíncrona:
    - `btnAgVinculaciones.Click += async (s, e) => await AbrirVinculacionesModoAgAsync();`
    - `btnAgVerVinculaciones.Click += async (s, e) => await AbrirDetalleSesionesModoAgAsync();`
    - `btnAgSesiones.Click += async (s, e) => await AbrirSesionModoAgAsync();`
  - Se implementó `AbrirVinculacionesModoAgAsync()` instanciando repositorios y servicio canónicos con `AgModoDatos.Ag`, abriendo modal con `ExcelDnaUtil.Application` y refrescando con `CargarModuloResultadosAsync` al cerrar.
  - Se implementó `AbrirDetalleSesionesModoAgAsync()` validando filas activas, consultando `vw_ag_sesiones_detalle` mediante `VinculacionQueryRepository` y abriendo `FormVerVinculaciones`.
  - Se implementó `AbrirSesionModoAgAsync()` y `ObtenerIdVinculacionSeleccionadaAsync()` extrayendo `IdSesion` del grid y resolviendo `id_vinculacion` desde `ag_sesiones`.

### 126. **Fase 299 — Indicador Visual "Modo AG" y Títulos de Formularios**
- **Archivos Modificados:**
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.Designer.cs`
  - `ConversorPDF.XLL/Academic/UI/Forms/FormSesionesAcademicas.cs`
- **Diagnóstico y Análisis:**
  `FormVinculaciones` y `FormVerVinculaciones` ya actualizaban su título con `[MODO AG - STAGING]` en caso de `_modo.EsModoAg`, pero `FormSesionesAcademicas` no contaba con esta distinción visual en su barra de título tras `InitializeComponent()`.
- **Solución Aplicada:**
  - En `FormSesionesAcademicas.cs` (constructor), se añadió la verificación:
    ```csharp
    if (_modo.EsModoAg)
    {
        this.Text = $"{this.Text} [MODO AG - STAGING]";
    }
    ```
  - En `AgFormGeneradorHorarios.Designer.cs`, `lblAgModo` refleja el indicador canónico de staging: `"Modo AG · staging: ag_vinculaciones / ag_sesiones"`.

### 127. **Fase 303 — Bloqueo Estricto por Cobertura Incompleta de Pool Docente en Pre-Vuelo y Optimización**
- **Archivo Modificado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico y Causa Raíz:**
  Anteriormente, la compuerta de validación de pool docente en `EjecutarDiagnosticoPreVueloAsync` e `IniciarOptimizacionAsync` sólo bloqueaba cuando `asigConPool == 0 && Docentes.Count == 0`, emitiendo una mera advertencia cuando la cobertura era parcial (`asigSinPool > 0`). En escenarios reales donde un programa contenía asignaturas sin docentes habilitados en `terceros_asignaturas` (ej. 5 con pool y 1 sin pool), el pre-vuelo permitía continuar, generando asignaciones inválidas con `Docente #0` y colapsando en violaciones duras de la regla R5 (`R5_HabilitacionDocentePool`).
- **Solución Aplicada:**
  - Se modificaron `EjecutarDiagnosticoPreVueloAsync` e `IniciarOptimizacionAsync` para calcular `asignaturasSinPool` y evaluar `if (asigSinPool > 0)`.
  - Si existen asignaturas sin docente habilitado activo en el pool, el flujo se bloquea inmediatamente (`return;`), notificando en `txtLogAG` y en un cuadro de diálogo descriptivo con la lista detallada de asignaturas afectadas (`Id:Nombre`), y registrando `AgCausaFallo.SinPoolDocente` en la traza estructurada.

### 128. **Fase 304 — Blindaje contra Asignación de Genes con `IdTercero <= 0` en Fábrica de Individuos**
- **Archivo Analizado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`
- **Diagnóstico y Verificación:**
  Se evaluó la inicialización en `AgFabricaIndividuos.CrearAleatorio`. La fábrica ya implementaba la salvaguarda fail-fast estricta:
  ```csharp
  if (idDocente <= 0)
  {
      throw new InvalidOperationException($"Asignatura {asig.Id} ('{asig.Nombre}') sin docente asignable en pool ni catálogo general.");
  }
  ```
  La cual, combinada con el bloqueo temprano de la Fase 303 en pre-vuelo y optimización, garantiza que ningún gen se instancie con `IdTercero <= 0` ni en la población inicial ni durante el proceso evolutivo.

### 129. **Fase 305 — Inicialización en Código del Indicador `lblAgModo`**
- **Archivo Modificado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:**
  Aunque `lblAgModo` poseía texto predeterminado en el Designer, en el código-behind no se inicializaban explícitamente sus propiedades de estado, dependiendo exclusivamente del archivo de diseño.
- **Solución Aplicada:**
  - En el constructor de `AgFormGeneradorHorarios.cs`, tras `InitializeComponent()`, se añadió la inicialización programática canónica:
    ```csharp
    if (lblAgModo != null)
    {
        lblAgModo.AutoSize = true;
        lblAgModo.ForeColor = Color.OrangeRed;
        lblAgModo.Text = $"Modo AG · staging: {AgModoDatos.Ag.TablaVinculaciones} / {AgModoDatos.Ag.TablaSesiones}";
    }
    ```

### 130. **Fase 306 — Orden de Acoplamiento (Docking) de `pnlAgModo` en Módulo R**
- **Archivo Verificado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.Designer.cs`
- **Diagnóstico y Verificación:**
  Se verificó la jerarquía de acoplamiento en `tabResultados`:
  `tabResultados.Controls.Add(splitContainerResultados)` (Fill), seguido de `tabResultados.Controls.Add(pnlAgModo)` (Top), y finalmente `tabResultados.Controls.Add(pnlHeaderResultados)` (Top). WinForms procesa los controles dockeados en orden inverso de inserción, posicionando `pnlHeaderResultados` arriba, `pnlAgModo` en el medio y `splitContainerResultados` llenando el espacio restante de forma matemáticamente consistente y sin colisiones.

### 131. **Fase 307 — Encapsulamiento de Excepciones y Trazabilidad en `EjecutarOptimizacionAsync`**
- **Archivos Modificados:**
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgGeneradorHorarios.cs`
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico y Causa Raíz:**
  El método `EjecutarOptimizacionAsync` carecía de bloque `try/catch`. Cualquier excepción lanzada durante la inicialización de población (como `InvalidOperationException` desde `AgFabricaIndividuos`), mutación o evaluación de aptitud escapaba del servicio sin retornar un `Result.Failure<AgResultadoCorrida>` manejado ni registrar el error en bitácora, derivando en un crash no controlado en la capa de interfaz de usuario.
- **Solución Aplicada:**
  - En `AgGeneradorHorarios.cs`, se envolvió la ejecución completa de `EjecutarOptimizacionAsync` en un bloque `try/catch`:
    - `catch (OperationCanceledException)`: Retorna `Result.Failure<AgResultadoCorrida>("Optimización cancelada por el usuario.")`.
    - `catch (Exception ex)`: Registra en `LoggerHelper.Error` y retorna `Result.Failure<AgResultadoCorrida>($"Error durante la optimización: {ex.Message}")`.
  - En `AgFormGeneradorHorarios.cs` (`IniciarOptimizacionAsync`), se añadió en los bloques `catch` la emisión estructurada de `traza.Fallo("P08_EjecutarOptimizacionAsync", ex, AgCausaFallo.ExcepcionNoControlada)` y `traza.Fin(false, ...)`.

### 132. **Fase 308 — Protocolo de Verificación Operativa y Estabilidad del Motor Genético**
- **Alcance:** Verificación formal de robustez funcional:
  1. Con asignaturas sin pool docente: el diagnóstico pre-vuelo y la compuerta de inicio bloquean preventivamente listando `Id:Nombre`.
  2. Emisión en telemetría de `CoberturaDocente: n/m` y `AsignaturasSinDocente`.
  3. Ausencia de `Docente #0` y erradicación de violaciones duras `R5_HabilitacionDocentePool`.
  4. Indicador visual `lblAgModo` visible con `"Modo AG · staging: ag_vinculaciones / ag_sesiones"`.
  5. Contención total de excepciones imprevistas mediante `Result.Failure`, previniendo crashes en la aplicación de escritorio.

### 133. **Fase 309 — Corrección de Columna `idvinculacion` en `ag_sesiones`**
- **Archivo Modificado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico y Causa Raíz:**
  En `ObtenerIdVinculacionSeleccionadaAsync`, la consulta SQL utilizaba `id_vinculacion` (`SELECT id_vinculacion FROM ag_sesiones ...`). Según la definición canónica del DDL en `MigrationV49_AddAgHorariosTables.cs` (línea 201), el identificador es `idvinculacion`. Esto provocaba la excepción `SQLiteException: no such column: id_vinculacion` al pulsar el botón "✏️ Sesión (AG)".
- **Solución Aplicada:**
  - Se corrigió la consulta en `ObtenerIdVinculacionSeleccionadaAsync`:
    ```csharp
    const string sql = "SELECT idvinculacion FROM ag_sesiones WHERE id = @IdSesion LIMIT 1;";
    int? idVinc = await db.ExecuteScalarAsync<int?>(sql, new { IdSesion = idSesion });
    ```

### 134. **Fase 310 — Desduplicación de Título de Ventana en `FormSesionesAcademicas`**
- **Archivo Modificado:** `ConversorPDF.XLL/Academic/UI/Forms/FormSesionesAcademicas.cs`
- **Diagnóstico:**
  La asignación previa en `OnLoad` (`"Sesiones Académicas — MODO AG"`) concatenada con el sufijo del constructor (`"[MODO AG - STAGING]"`) generaba títulos duplicados redundantes: `"Sesiones Académicas — MODO AG [MODO AG - STAGING]"`.
- **Solución Aplicada:**
  - Se unificó la asignación tanto en el constructor como en `OnLoad`:
    ```csharp
    this.Text = _modo.EsModoAg ? "Sesiones Académicas [MODO AG - STAGING]" : "Sesiones Académicas";
    ```

### 135. **Fase 311 — Interpolación Dinámica de Vista en Mensaje de Detalle de Sesiones**
- **Archivo Modificado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:**
  En `AbrirDetalleSesionesModoAgAsync`, el cuadro de advertencia ante 0 resultados contenía el nombre de vista codificado en minúsculas (`vw_ag_sesiones_detalle`), en lugar de interpolar la propiedad canónica del modelo `modo.VistaSesionesDetalle` (`vw_AgSesionesDetalle`).
- **Solución Aplicada:**
  - Se interpoló dinámicamente `modo.VistaSesionesDetalle`:
    ```csharp
    MessageBox.Show(this, $"No se encontraron sesiones registradas en la vista de detalle staging ({modo.VistaSesionesDetalle}).", "Modo AG", MessageBoxButtons.OK, MessageBoxIcon.Warning);
    ```

### 136. **Fase 312 — Protocolo de Verificación Operativa de Integración Módulo R → Modo AG**
- **Alcance:** Verificación formal de integración:
  1. Resolución exitosa de `idvinculacion` al pulsar "✏️ Sesión (AG)" en la grilla y apertura modal de `FormSesionesAcademicas` con título limpio.
  2. Consulta tipada de `vw_AgSesionesDetalle` al pulsar "📋 Detalle (AG)".
  3. Refresco sincronizado de `dgvHorariosGenerados` tras el cierre de los diálogos modales.

### 137. **Fase 313 — Validación Preventiva de IdVinculacion en Apertura de Sesión Modo AG**
- **Archivo Modificado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico y Causa Raíz:**
  En `AbrirSesionModoAgAsync`, cuando `ObtenerIdVinculacionSeleccionadaAsync` no resolvía un valor válido (por fila sin `IdSesion`, celda vacía o sesión inexistente en `ag_sesiones`), se invocaba ciegamente `AbrirEnModoAg(idVinc)`. Dado que `idVinc` era nulo, `FormSesionesAcademicas` se abría como una vinculación nueva en borrador sin advertir al usuario.
- **Solución Aplicada:**
  - Se introdujo una comprobación estricta con guard clause antes de abrir el diálogo modal:
    ```csharp
    int? idVinc = await ObtenerIdVinculacionSeleccionadaAsync();
    if (!idVinc.HasValue || idVinc.Value <= 0)
    {
        MessageBox.Show(this, "Seleccione una fila de resultados con IdSesion válido para editar en Modo AG.",
                        "Modo AG", MessageBoxButtons.OK, MessageBoxIcon.Information);
        return;
    }
    AbrirEnModoAg(idVinc);
    ```

### 138. **Fase 314 — Persistencia de Estado de Corrida y Rehidratación de Resultados en Módulo R**
- **Archivo Modificado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico y Causa Raíz:**
  La variable `_ultimoResultado` únicamente se asignaba en memoria viva al término de `IniciarOptimizacionAsync`. Al reiniciar Excel o recargar una corrida histórica mediante `CargarModuloResultadosAsync(idCorrida)`:
  1. `_ultimoResultado` era `null`, provocando que `CargarModuloResultadosAsync` retornara tempranamente sin poblar la grilla ni las etiquetas.
  2. Al intentar abrir "📋 Detalle (AG)" mediante `AbrirDetalleSesionesModoAgAsync`, la guarda `dgvHorariosGenerados.Rows.Count == 0 || _ultimoResultado == null` generaba un bloqueo falso con el mensaje "Ejecute 'Optimizar Horarios'…", incluso con las 72 sesiones generadas en base de datos.
  3. Los eventos de cierre de los formularios modales de Modo AG (`AbrirVinculacionesModoAgAsync`, `AbrirDetalleSesionesModoAgAsync`, `AbrirSesionModoAgAsync`) no refrescaban la vista porque evaluaban `_ultimoResultado?.IdCorrida > 0` (falso).
  4. Los métodos de compuerta `PromoverAProduccionAsync` y `AprobarCorridaYAulaAsync` retornaban inmediatamente por condición nula de `_ultimoResultado`.
- **Solución Aplicada:**
  - Se incorporó el campo de instancia `private long _idCorridaCargada;` para rastrear persistentemente el identificador de corrida visualizada.
  - En `IniciarOptimizacionAsync` se asigna `_idCorridaCargada = idCorrida;`.
  - En `CargarModuloResultadosAsync(long idCorrida)`, se asigna `_idCorridaCargada = idCorrida;` y se añade lógica de rehidratación automática: si `_ultimoResultado?.MejorIndividuo == null`, se consulta `_scheduleRepository.GetCorridaByIdAsync(idCorrida)` y `_scheduleRepository.GetIndividuoByIdAsync(corridaDb.IdIndividuoOptimo.Value)`, reconstruyendo `_ultimoResultado` con su individuo óptimo, aptitud y métricas de convergencia.
  - Se protegió la carga de violaciones contra colecciones nulas: `(_ultimoResultado.MejorIndividuo.Violaciones ?? new List<AgViolacion>())`.
  - Se actualizaron las guardas de `AbrirDetalleSesionesModoAgAsync` a `if (dgvHorariosGenerados.Rows.Count == 0 && _idCorridaCargada <= 0)`.
  - Se migraron los callbacks de refresco modal para evaluar `if (_idCorridaCargada > 0) await CargarModuloResultadosAsync(_idCorridaCargada);`.
  - En `PromoverAProduccionAsync` y `AprobarCorridaYAulaAsync` se agregó la rehidratación transparente vía `CargarModuloResultadosAsync(_idCorridaCargada)`.

### 139. **Fase 315 — Protocolo de Verificación Operativa post-Reinicio y Modo AG**
- **Alcance:** Verificación formal de robustez funcional y operacional:
  1. Al reabrir Excel y cargar una corrida en Módulo R, el botón "📋 Detalle (AG)" abre el diálogo `FormVerVinculaciones` con las 72 sesiones proyectadas desde `vw_AgSesionesDetalle` sin bloqueos falsos.
  2. Al pulsar "✏️ Sesión (AG)" sobre una fila con `IdSesion` válido, se resuelve su `idvinculacion` en `ag_sesiones` y se abre `FormSesionesAcademicas` en `[MODO AG - STAGING]`.
  3. Al pulsar "✏️ Sesión (AG)" sin fila seleccionada o sobre una fila sin `IdSesion` resuelto, se muestra un mensaje explicativo sin instanciar un formulario en blanco.
  4. El cierre de cualquiera de los formularios de auditoría de Modo AG refresca `dgvHorariosGenerados` de forma determinista sin duplicación de registros.
  5. Las acciones de compuerta ("Aprobar Corrida y Aula" y "Promover a Producción") permanecen operables tanto en la sesión activa como tras recargar corridas históricas.

### 140. **Fase 317 — Esquema BD: Columnas de Vacante (`MigrationV54_AddVacanteColumns`)**
- **Archivos Modificados y Creados:**
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV54_AddVacanteColumns.cs` (Nuevo)
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV49_AddAgHorariosTables.cs` (Espejo en DDL base)
  - `ConversorPDF.XLL/Core/Helpers/SqliteHelper.cs` (Registro de V54)
  - `ConversorPDF.XLL/ConversorPDF.XLL.csproj` (Inclusión del archivo de migración)
- **Detalle de Esquema:**
  - `ag_vinculaciones`: `es_vacante INTEGER DEFAULT 0`, `motivo_vacante TEXT DEFAULT ''`, `candidatos_pool INTEGER DEFAULT 0`.
  - `ag_sesiones`: `es_vacante INTEGER DEFAULT 0`, `motivo_vacante TEXT DEFAULT ''`, `candidatos_pool INTEGER DEFAULT 0`.
  - `vinculaciones`: `es_vacante INTEGER DEFAULT 0`, `motivo_vacante TEXT DEFAULT ''`.
  - `sesiones`: `es_vacante INTEGER DEFAULT 0`, `motivo_vacante TEXT DEFAULT ''`.
  - Migración tolerante a fallos (`IMigrationTolerante`) e idempotente ante columnas preexistentes.

### 141. **Fase 318 — Modelos de Dominio y Tipado Fuerte de Vacantes**
- **Archivos Modificados y Creados:**
  - `ConversorPDF.XLL/Academic/Domain/Enums/AgMotivoVacante.cs`:
    - `Ninguno = 0`, `SIN_CANDIDATOS = 1`, `SIN_DISPONIBILIDAD = 2`, `NO_CUMPLE_REQUISITOS = 3`, `POOL_VACIO = 4`, `FueraDeVentanaSemestre = 5`.
  - `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgGen.cs`: Propiedades `EsVacante`, `MotivoVacante`, `CandidatosPool` y clonación en `Clonar()`.
  - `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgVinculacion.cs`: Propiedades `EsVacante`, `MotivoVacante`, `CandidatosPool`.
  - `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgSesion.cs`: Propiedades `EsVacante`, `MotivoVacante`, `CandidatosPool`.
  - `ConversorPDF.XLL/Academic/Domain/Models/Vinculacion.cs`: Propiedades `EsVacante`, `MotivoVacante`.
  - `ConversorPDF.XLL/Academic/Domain/Models/Sesion.cs`: Propiedades `EsVacante`, `MotivoVacante`.
  - `ConversorPDF.XLL/Academic/Application/DTOs/AgVacanteGridDto.cs` (Nuevo DTO de presentación para el grid de vacantes).
  - `ConversorPDF.XLL/ConversorPDF.XLL.csproj` (Registro de `AgVacanteGridDto.cs`).

### 142. **Fase 319 — Clasificación del Motivo de Vacante en Fábrica Evolutiva**
- **Archivo Modificado:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`
- **Diagnóstico y Solución:**
  - Se eliminó el `throw new InvalidOperationException` que abortaba la simulación cuando una asignatura carecía de docentes en pool o docentes habilitados en contexto.
  - Se implementó `ClasificarMotivo(AgContexto c, Asignatura a, DateTime? fecha, AgBloqueHorario bloque)`:
    1. Si `PoolPorAsignatura` es nulo o vacío: si no hay docentes en contexto retorna `POOL_VACIO`, de lo contrario `SIN_CANDIDATOS`.
    2. Si el pool contiene registros pero ninguno activo (`Activo == 1`), retorna `NO_CUMPLE_REQUISITOS`.
    3. Si todos los docentes habilitados están ocupados en la franja/fecha (`TerceroOcupado`), retorna `SIN_DISPONIBILIDAD`.
    4. De lo contrario, retorna `NO_CUMPLE_REQUISITOS`.
  - El gen se genera de forma segura con `IdTercero = 0`, `EsVacante = 1`, `MotivoVacante = motivo`, `CandidatosPool = candidatosCount`.

### 143. **Fase 320 — Relajación de Regla R5 y Métrica de Vacantes**
- **Archivos Modificados:**
  - `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgPesosFitness.cs`: Inclusión de `PesoVacante = 50.0`.
  - `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgIndividuo.cs`: Métrica calculada `Vacantes => Genes?.FindAll(g => g.EsVacante == 1 || g.IdTercero <= 0).Count ?? 0`.
  - `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgResultadoCorrida.cs`: Propiedad `Vacantes => MejorIndividuo?.Vacantes ?? 0`.
  - `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgCorrida.cs`: Propiedad `Vacantes`.
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorReglas.cs`:
    - La restricción dura `R5_HabilitacionDocentePool` ahora solo se evalúa sobre genes reales no vacantes (`!gen.EsVacante && gen.IdTercero > 0`).
    - En el cómputo de penalizaciones blandas, cada vacante suma `ctx.Pesos.PesoVacante`, incentivando al algoritmo a cubrirlas sin bloquear la corrida con violaciones duras.
  - `ConversorPDF.XLL/Academic/Application/Mappers/AgGenToAgVinculacionMapper.cs`: Propagación de `EsVacante`, `MotivoVacante`, `CandidatosPool`.
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgExpansorSesiones.cs`: Propagación de atributos de vacante al expandir sesiones.
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgScheduleRepository.cs`: Persistencia y lectura de columnas de vacante en `GuardarStagingAsync`, `GetVinculacionesStagingAsync` y `GetSesionesStagingAsync`.
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AcademicRepository.cs`: Persistencia de vacantes en producción (`InsertVinculacion`, `InsertSesion`, `InsertSesiones`).

### 144. **Fase 321 — Diagnóstico Pre-Vuelo: de Bloqueo a Consentimiento Explícito**
- **Archivo Modificado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico y Solución:**
  - Se modificaron las compuertas de pre-vuelo en `ValidarFactibilidadVentanaAsync` e `IniciarOptimizacionAsync`.
  - En lugar de detener la ejecución con mensaje de error ante asignaturas sin pool (`asigSinPool > 0`), se despliega un diálogo de confirmación `MessageBoxButtons.YesNo`:
    *"Hay {asigSinPool} de {totalAsig} asignaturas sin docente habilitado en el pool... Se marcarán como VACANTE (es_vacante=1) con penalización suave y podrá resolverlas después en el Panel de Vacantes. ¿Continuar con la optimización?"*.
  - Si el usuario consiente, la optimización procede generando genes vacantes controlados.

### 145. **Fase 322 — Panel de VACANTES en Módulo R y Acciones Contextuales**
- **Archivos Modificados:**
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.Designer.cs`:
    - Incorporación de `lblVacantes` ("Vacantes: N") en `pnlHeaderResultados`.
    - Sustitución del contenedor inferior por `tabAuditoriaResultados` con dos pestañas:
      1. `tabViolaciones` ("⚠️ Violaciones de Reglas"): Grilla de violaciones (`dgvViolaciones`).
      2. `tabVacantes` ("👤 Panel de Vacantes"): Panel de filtro (`pnlFiltroVacantes`), combo de motivos (`cmbMotivoVacante`), botón contextual (`btnAccionVacante`), contador (`lblConteoVacantesTab`) y grilla de vacantes (`dgvVacantes`).
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`:
    - Cableado de eventos: `cmbMotivoVacante.SelectedIndexChanged`, `dgvVacantes.SelectionChanged`, `btnAccionVacante.Click`.
    - Población dinámica de `dgvVacantes` a partir de `sesionesResult` filtrando por `EsVacante == 1 || IdTercero <= 0`.
    - Actualización contextual del botón `btnAccionVacante`:
      - `SIN_CANDIDATOS` o `POOL_VACIO` → "Cargar pool" (abre `FormTerceroAsignaturas` prefiltrado por programa, asignatura y semestre).
      - `SIN_DISPONIBILIDAD` → "Mover franja" (abre `FormSesionesAcademicas` en Modo AG para la vinculación seleccionada).
      - `NO_CUMPLE_REQUISITOS` → "Revisar requisitos" (abre `FormTerceroAsignaturas`).
    - Al cerrar los diálogos modales, refresco automático de resultados recalculando el contador de vacantes.
    - Advertencia preventiva en `PromoverAProduccionAsync` requiriendo confirmación si existen vacantes sin resolver.
    - Exportación CSV ampliada con columnas `Vacante`, `Motivo` y `Candidatos`.

### 146. **Fase 323 — Pruebas Unitarias de Clasificación y Fábrica de Vacantes**
- **Archivos Creados y Modificados:**
  - `ConversorPDF.Tests/Academic/Domain/AgMotivoVacanteTests.cs` (Nuevo):
    - `ClasificarMotivo_WhenPoolIsNullAndNoDocentes_ShouldReturnPoolVacio`
    - `ClasificarMotivo_WhenPoolIsNullButDocentesExist_ShouldReturnSinCandidatos`
    - `ClasificarMotivo_WhenPoolHasRowsButAllInactive_ShouldReturnNoCumpleRequisitos`
    - `ClasificarMotivo_WhenPoolActiveTeachersAreAllOccupied_ShouldReturnSinDisponibilidad`
    - `AgFabricaIndividuos_WhenNoDocenteHabilitado_ShouldCreateVacantGenWithoutThrowing`
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgValidadorReglasTests.cs`:
    - `Evaluar_WhenGenIsVacante_ShouldNotRecordR5HardViolation_AndApplySoftPenalty`

### 147. **Fase 324 — Sincronización Documental y Exportación**
- **Archivos Modificados:**
  - `docs/articles/changes/2026-ag-horarios-academicos.md`: Inclusión de la Sección 8.22 con descripción integral de las Fases 317 a 324.
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`: Formato CSV con columnas `Vacante`, `Motivo`, `Candidatos`.

### 148. **Fase 325 — Formato Unificado de Asignaturas y Semestre Romano (`AgFormatoAsignatura`)**
- **Archivos Creados y Modificados:**
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFormatoAsignatura.cs` (Nuevo)
  - `ConversorPDF.XLL/ConversorPDF.XLL.csproj` (Inclusión del nuevo archivo)
- **Detalle de Implementación:**
  - `SemestreRomano(string)`: Mapeo tolerante de entradas arábigas (`"1"`..`"12"` → `"I"`..`"XII"`), variantes en minúsculas (`"ii"` → `"II"`), cadenas vacías o nulas (`"SIN ASIGNAR"`) y passthrough para cadenas no mapeadas.
  - `Linea(Asignatura a)`: Construcción unificada `"{a.Id}: {SemestreRomano(a.Semestre)} {Flecha} {a.Nombre}"`.
  - `BulletList(IEnumerable<Asignatura>)`: Encadenamiento con viñeta `"• "` (`string.Join("\n• ", ...)`).
  - Configuración opcional de `Flecha` (`"→"` por defecto, configurable a `"->"`).

### 149. **Fase 326 — Integración en Puntos de Llamada (`AgFormGeneradorHorarios.cs`)**
- **Archivo Modificado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Detalle de Implementación:**
  - Call site 1 (L771 - `ValidarFactibilidadVentanaAsync`):
    - `string listaSin = AgFormatoAsignatura.BulletList(asignaturasSinPool);`
    - `traza.Dato("AsignaturasSinDocente", listaSin.Replace("\n", " "));`
    - `Log($"[PRE-VUELO] {asigSinPool}/{totalAsig} asignaturas SIN docente habilitado en 'terceros_asignaturas':\n• {listaSin}");`
    - Presentación en `MessageBox.Show` con formato `• {listaSin}`.
  - Call site 2 (L935 - `IniciarOptimizacionAsync`):
    - `string listaSinOpt = AgFormatoAsignatura.BulletList(asignaturasSinPoolOpt);`
    - `traza.Dato("AsignaturasSinDocente", listaSinOpt.Replace("\n", " "));`
    - `Log($"[OPTIMIZACIÓN] {asigSinPoolOpt}/{totalAsigOpt} asignaturas SIN docente habilitado en 'terceros_asignaturas':\n• {listaSinOpt}");`
    - Presentación en `MessageBox.Show` con formato `• {listaSinOpt}`.

### 150. **Fase 327 — Pruebas Unitarias de Formato (`AgFormatoAsignaturaTests`)**
- **Archivo Creado:** `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgFormatoAsignaturaTests.cs`
- **Casos de Prueba Cubiertos:**
  - `SemestreRomano`: `"2"` → `"II"`, `"ii"` → `"II"`, `"VIII"` → `"VIII"`, `""`/`null`/espacios → `"SIN ASIGNAR"`, `"XI"` → `"XI"`, números arábigos 1 al 10 a sus correspondientes romanos.
  - `Linea`: Objeto con `Id = 132`, `Semestre = "II"`, `Nombre = "MATEMATICAS II"` → `"132: II → MATEMATICAS II"`; objeto nulo → `string.Empty`.
  - `BulletList`: Colección de 2 asignaturas formateada con viñetas `"132: II → MATEMATICAS II\n• 133: II → PROYECTO DE FUTURO I"`.
  - `Flecha`: Configuración temporal a `"->"` y verificación de salida ASCII pura.

### 151. **Fase 328 — Documentación y Verificación de Salida**
- **Archivos Modificados:**
  - `docs/articles/changes/2026-ag-horarios-academicos.md`: Sección 8.23 detallando la especificación y comportamiento del formato legible.
  - `implementation_plan.md`: Registro de las Fases 325 a 328.

### 152. **Fase 329 — Normalización Canónica de `AgMotivoVacante` y Extensiones (`AgMotivoVacanteExtensions`)**
- **Archivos Modificados y Creados:**
  - `ConversorPDF.XLL/Academic/Domain/Enums/AgMotivoVacante.cs`: Eliminación de alias con duplicados de valor entero (`SinCandidatos=1`, `SinDocenteHabilitado=1`, etc.), formalizando los valores canónicos en mayúsculas: `Ninguno = 0`, `SIN_CANDIDATOS = 1`, `SIN_DISPONIBILIDAD = 2`, `NO_CUMPLE_REQUISITOS = 3`, `POOL_VACIO = 4`, `FUERA_DE_VENTANA = 5`.
  - `ConversorPDF.XLL/Academic/Domain/Enums/AgMotivoVacanteExtensions.cs` (Nuevo): Creación según la regla de un solo tipo por archivo con métodos `.Codigo()` (retorna código en mayúsculas o cadena vacía si es Ninguno) y `.Desde(string)` (parseo insensible a mayúsculas con fallback seguro a Ninguno).
  - `ConversorPDF.XLL/ConversorPDF.XLL.csproj`: Registro de `AgMotivoVacanteExtensions.cs`.
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgAsignadorTercero.cs`: Actualización para asignar y diagnosticar motivos utilizando `.Codigo()`.
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`: Asignación de `motivoVacante = motivo.Codigo()`.
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorVentana.cs`: Asignación a `AgMotivoVacante.SIN_CANDIDATOS`.
  - `ConversorPDF.XLL/Academic/Application/DTOs/AgVacanteGridDto.cs`: Proyección determinista `Motivo => MotivoVacante.Codigo()`.
  - `ConversorPDF.Tests/Academic/Domain/AgMotivoVacanteTests.cs`: Pruebas de compatibilidad y parseo en `Codigo()` y `Desde()`.
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgAsignadorTerceroTests.cs`: Ajuste de aserciones al código canónico.

### 153. **Fase 330 — Promoción Parcial con Justificación Formal de Vacantes**
- **Archivos Modificados:**
  - `ConversorPDF.XLL/Academic/Domain/Interfaces/IAgPromotorVinculaciones.cs`: Incorporación de sobrecarga `PromoverAProduccionAsync(long idCorrida, long idIndividuo, string usuario, string justificacionVacante, CancellationToken ct)`.
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPromotorVinculaciones.cs`:
    - Relajación de la compuerta dura de vacantes: si existen vacantes (`nVac > 0`) pero se suministra `justificacionVacante`, se permite la promoción parcial de las asignaturas cubiertas.
    - Registro de `es_vacante = 1` y `motivo_vacante` en entidades de producción (`Vinculacion`, `Sesion`).
    - Asiento de `justificacion_vacante` en la tabla `ag_promocion`.
    - Omisión segura de vacantes al sincronizar pool docente post-promoción (`sv.IdTercero > 0`).
  - `ConversorPDF.XLL/Academic/Application/Mappers/AgVinculacionToVinculacionMapper.cs`: Mapeo explícito de `EsVacante` y `MotivoVacante`.
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AcademicRepository.cs`: Inclusión de `es_vacante` y `motivo_vacante` en `UpdateVinculacion` y `GetVinculacionById`.
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`: Al detectar vacantes activas y confirmar la promoción, se genera la justificación formal y se transmite a `_promotor.PromoverAProduccionAsync(...)`.
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgPromotorVinculacionesTests.cs`: Nueva prueba unitaria verificando que ante vacantes con justificación, la ejecución supera la compuerta de vacantes.

### 154. **Fase 331 — Población Dinámica del Combo de Motivos y Soporte de `FUERA_DE_VENTANA`**
- **Archivos Modificados:**
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`:
    - Población dinámica de `cmbMotivoVacante` mediante `Enum.GetValues(typeof(AgMotivoVacante)).Cast<AgMotivoVacante>().Where(m => m != AgMotivoVacante.Ninguno).Select(m => m.Codigo()).Distinct()`.
    - Soporte contextual en `btnAccionVacante` para `FUERA_DE_VENTANA` ("Ajustar ventana") abriendo Modo AG para corregir fechas.
    - Parseo robusto con `AgMotivoVacanteExtensions.Desde(s.MotivoVacante)` al alimentar la colección `_todasLasVacantes`.

### 155. **Fase 332 — Verificación Visual y Espacial del Panel de Resultados**
- **Verificación en Designer:**
  - Disposición horizontal en `pnlHeaderResultados`:
    - `lblMejorFitness`: X=15, Width=155 (Right=170)
    - `lblHardViolaciones`: X=195, Width=147 (Right=342)
    - `lblVacantes`: X=375, Width=85 (Right=460)
    - `btnExportarSimulacion`: X=580, Width=150 (Right=730)
    - `btnPromoverProduccion`: X=740, Width=200 (Right=940)
  - Espacio libre entre `lblVacantes` (Right=460) y `btnExportarSimulacion` (Left=580) de 120 px. Cero solapes espaciales confirmado.

### 156. **Fase 333 — Acción Contextual y Navegación Dual para `FUERA_DE_VENTANA`**
- **Archivo Modificado:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Detalle de Implementación:**
  - `ActualizarBotonContextualVacante`: Asigna el texto `"Ajustar ventana"` cuando `item.MotivoVacante == AgMotivoVacante.FUERA_DE_VENTANA`.
  - `EjecutarAccionVacanteAsync`: Implementación de flujo interactivo dual mediante diálogo modal:
    - Si el usuario selecciona 'Sí': abre la sesión en Modo AG (`FormSesionesAcademicas`) para editar o mover las fechas específicas de la sesión fuera de ventana.
    - Si el usuario selecciona 'No': redirige inmediatamente a la pestaña de configuración del Módulo S (`tabControlPrincipal.SelectedTab = tabSimulacion`) para permitir ampliar la ventana semestral o revisar fechas de período.
    - Si cancela: descarta la acción sin alterar el estado.

---

## 157. PLAN v14 — Modelo de Proyección Semestral (Fases 378-390) · Revisión crítica

> [!IMPORTANT]
> Estado: **PROPUESTO — PENDIENTE DE APROBACIÓN**. No se ha modificado código. No compilado ni ejecutado.

### 157.1 Veredicto por fase (v14)

| Fase | Veredicto | Motivo |
|---|---|---|
| F378 | Válida con corrección | `TerceroAsignaturaKey.Calcular` es PK física de `terceros_asignaturas`; normalizarla cambia IDs ⇒ duplicados (→ F399). APIs net48 (→ F393). |
| F379 | Válida con corrección | Usa sólo `combinaciones.First()`; corrida = 1 aula ⇒ "219 filas" no es aceptación por corrida (→ F400). |
| F380 | Parcial | Parche A: cast `(DayOfWeek)7` inválido (→ F396). Parche B: semanas de α, no de la fila; tope contradictorio (→ F401). Parche C: no compilable; `R-DOC` por día = **falso positivo** (→ F401). |
| F381 | **Rechazada como está** | Fórmula no reproduce histogramas; escala días fuera de K contradiciendo H6 (→ F391, F392, F393). |
| F382 | Válida con corrección | `Asignatura.DiasSemana` no existe (→ F394); sin franja Lunes ⇒ `First()` lanza (→ F395); redefine `TotalHorasAsignatura` (→ F397); grupo colisiona en cruce (→ F398); V55 incompleta (→ F402); formato `DiaHorasSemanales` (→ F403). |
| F383 | Válida con corrección | `vin.Value` sin guarda; `AgCausaFallo.StagingVacio` inexistente (→ F404). |
| F384 | Válida | Mover antes del smoke (→ F407). |
| F385 | Válida | — |
| F386 | Válida con corrección | Opción A sin fuente de `FechaInicio` en BD (→ F406). |
| F387 | Válida con corrección | Nombres de controles erróneos (→ F405). |
| F388 | Válida | Depende de `dias_patron` (F394/F402). |
| F389 | Válida con corrección | Casos de aceptación mal etiquetados (→ F392). |
| F390 | Válida | Usar `ExcelOptimizeHelper.Optimize` (regla Excel-DNA). |
| D12 | Falso positivo parcial | `Grupo` ya viaja en tupla S/`AulaRef`; `ParsearSemestreGrupo` sólo para import Excel. Defecto real: `Grupo = "A"` hardcode. |

### 157.2 Correcciones propuestas

#### Fase 391 — R-PS: generador por recorrido de calendario (reemplaza fórmula)
- **Causa:** `64h/2d/5sem` ⇒ fórmula da `[4,4,4,4]`, no `[2,3,3,4,4]`; `48h/3d/5sem` ⇒ `[3,3,3,3]`, no `[1,2,3,3,3]`; `base/base+1` no produce dispersión >1. Histogramas son **emergentes** (semana inicial parcial + festivos).
- **Corrección:** `AgPlanificadorPatronSemanal.GenerarFechas(diasK, horas, fechaInicio, ventanaFin, esFestivo)`: recorrer días desde `FechaInicio`, aceptar si `DayOfWeek ∈ K` y no festivo, hasta `S = D/4`. Histograma = agrupación ISO-semana del resultado. Infactible ⇒ `Result.Failure` (no ampliar K).

#### Fase 392 — Casos de aceptación desde fixture real
- **Causa:** `80/4d/4sem ⇒ 4-4-4-4-4` son 5 semanas; `64/2d ⇒ 2-3-3-4-4` exige 4 días/semana con K=2 (contradice H6).
- **Corrección:** cada caso de `AgPlanificadorPatronSemanalTests` = fila real `(K, D, Ini, Fin, festivos) ⇒ fechas esperadas` del fixture `Horarios2026_02Fixture`.

#### Fase 393 — Compatibilidad .NET 4.8 / C# 7.3 / un tipo por archivo
- `Math.Clamp` ⇒ `Math.Max(1, Math.Min(x, tope))`.
- `Split('-', StringSplitOptions.RemoveEmptyEntries)` ⇒ `Split(new[] { '-' }, StringSplitOptions.RemoveEmptyEntries)`.
- `AgPatronSemanal` en archivo propio + constructor explícito; registrar ambos en `ConversorPDF.XLL.csproj`.

#### Fase 394 — Fuente de K (días del patrón)
- **Causa:** `Asignatura` no tiene `DiasSemana`.
- **Corrección:** V55 agrega `ag_aula_asignatura.dias_patron TEXT DEFAULT ''` (ISO `"4,5"`); `AgAulaAsignatura.DiasPatron`; `AgFabricaIndividuos` lee de `ctx.AulaAsignaturasPorId`; vacío ⇒ fallback a franjas activas del aula (comportamiento actual, 1 día).

#### Fase 395 — Franja Lunes inexistente
- **Causa:** V49 siembra días 2..7; patrones con LUNES (H7, ≥10 filas) ⇒ `bloquesValidos.First(...)` lanza `InvalidOperationException`.
- **Corrección:** V55 `INSERT ... (1,'18:00','22:00',4.0,'NOCHE',1,'') WHERE NOT EXISTS (dia_semana=1)`; fábrica usa `FirstOrDefault` ⇒ si null: gen vacante `MotivoVacante = SIN_DISPONIBILIDAD` + `traza.Dato("FranjaInexistente", ...)`.

#### Fase 396 — `ContarOcurrencias` con domingo
- **Causa:** `(DayOfWeek)7` no es `Sunday` ⇒ 0 ocurrencias domingo.
- **Corrección:** usar `b.DiaSemana` (propiedad ya mapea 7→Sunday en `AgBloqueHorario` L26).

#### Fase 397 — No redefinir `TotalHorasAsignatura`
- **Causa:** R6, `AgMetricas`, fitness y operadores la usan como total de asignatura.
- **Corrección:** nueva `AgGen.HorasGen` (horas del día del patrón); expansor consume `HorasGen` si `PatronSemanal != ""`; I1 = `Σ HorasGen(grupo) == TotalHorasAsignatura`.

#### Fase 398 — Clave de grupo determinista
- **Causa:** contador local `++_contadorGrupoIndividuo` colisiona entre padres en cruce.
- **Corrección:** `ClaveGrupo = $"{IdAsignatura}|{IdAula}|{Grupo}"`; cruce/mutación operan por grupo; `AgValidadorReglas` verifica mismo `IdTercero`/`IdAula`/`DiasPatron` por grupo.

#### Fase 399 — Normalizar en búsqueda, no en PK
- **Causa:** cambiar `TerceroAsignaturaKey.Calcular` invalida IDs existentes.
- **Corrección:** `Calcular` intacto; `AgContexto.EsDocenteHabilitado` prueba claves `{crudo, romano, arábigo, ""}`.

#### Fase 400 — Alcance por aula + grupo real
- **Corrección:** demanda/pool por la tupla de la corrida (`ctx.IdAulaCorrida`), no `combinaciones.First()`; aceptación 219/9340/2324 = Σ de corridas por lote. `Grupo = ctx.Grupo` (no `"A"`) en `AgGen`/`AgFabricaIndividuos`.

#### Fase 401 — Factibilidad y R-DOC
- F380-B: reutilizar `GenerarFechas` (F391) por asignatura; bloqueo si `Result.Failure`.
- `R-DOC` por (docente, día) = **falso positivo**: 98 pares pueden ser franjas distintas (Sáb MAÑANA/TARDE). Regla correcta ya existe: **R1 dura por (docente, fecha, franja)**. Eliminar Parche C; sólo advertencia informativa opcional.

#### Fase 402 — Migración V55 completa e idempotente
- `IMigrationTolerante` + verificación `PRAGMA table_info` por columna.
- Contenido: columnas `ag_gen` (`indice_dia_patron`, `patron_semanal`, `dias_patron`, `clave_grupo`, `horas_gen`), `ag_aula_asignatura.dias_patron`, seed Lunes.
- Actualizar `AgScheduleRepository` INSERT (L404) / SELECT (L487), `AgGen.Clonar`, `SqliteHelper`, `.csproj`.
- Corrección de datos de `periodos` (si se elige F406-B) ⇒ **V56** separada.

#### Fase 403 — Formato `DiaHorasSemanales`
- **Causa:** agregar `|18:00 A 22:00` rompe parsers existentes (`"JUEVES|8, VIERNES|4"`).
- **Corrección:** mantener `DIA|horas` separado por `", "`.

#### Fase 404 — Materialización segura
- Guardar `vin.IsSuccess`/`ses.IsSuccess` antes de `.Value`; agregar `AgCausaFallo.StagingVacio` (y `ColisionDocente` sólo si se usa).

#### Fase 405 — Nombres de controles F387
- `dgvHorarios` ⇒ `dgvHorariosGenerados`; `btnPromoverAProduccion` ⇒ `btnPromoverProduccion`.

#### Fase 406 — Ventana α vs Excel (H14)
- Opción A rechazada: no hay `FechaInicio` por asignatura en BD.
- Recomendado: **B** (V56 datos `periodo_inicio='2026-08-11'`) o ventana manual (prioridad 3 existente) para smoke.

#### Fase 407 — Orden de ejecución
- `F378 → F379 → F384 → F380 → F391/F381 → F382 → F383 → BUILD+TESTS+SMOKE → F385 → F386 → F387 → F388 → F389 → F390`.

#### Fase 408 — Decisiones pendientes cerradas
- Tope días/semana = **|K| (duro, H6)**; 5 sólo como cota de validación.
- `R-DOC` = **R1 dura por franja**; por día ⇒ no es regla.

### 157.3 Recomendación
- **No pasar a build/smoke.** Aplicar F391-F408 sobre el plan antes de implementar F381/F382.

---

### 157.4 Catálogo Exhaustivo de Cambios de Código y Soporte Técnico por Fase Afectada

> [!NOTE]
> Todo el diseño respeta estrictamente los principios de arquitectura de la solución: .NET Framework 4.8 / C# 7.3, tipos únicos por archivo físico, llamadas limpias a `ExcelOptimizeHelper.Optimize()`, no lanzar excepciones de flujo de negocio y preservación de aislamiento entre capas.

#### 1. Fase 378 / Fase 393 / Fase 399 — Normalización Segura de Semestre/Grupo
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFormatoAsignatura.cs`
  - `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgContexto.cs`
- **Cambios Técnicos:**
  1. En `AgFormatoAsignatura.cs`, agregar `SemestreArabigo(string semestre)` usando el mapeo canónico (1 al 12 ↔ I al XII) y passthrough defensivo.
  2. Implementar `ParsearSemestreGrupo(string valor)` con sintaxis C# 7.3 compatible: `valor.Trim().Split(new[] { '-' }, StringSplitOptions.RemoveEmptyEntries)`. Devuelve `(string Semestre, string Grupo)`.
  3. **Corrección F399:** Se mantiene intacto `TerceroAsignaturaKey.Calcular(...)` para no invalidar claves primarias ni registros históricos en la tabla física `terceros_asignaturas`.
  4. En `AgContexto.EsDocenteHabilitado(...)`, realizar búsqueda sucesiva de coincidencia en el diccionario probando: clave exacta, clave con `SemestreRomano`, clave con `SemestreArabigo` y clave con semestre vacío (docente habilitado general).

#### 2. Fase 379 / Fase 400 — Demanda Acotada por S (Eliminación de Fallback Global)
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
  - `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgContexto.cs`
- **Cambios Técnicos:**
  1. En `AgFormGeneradorHorarios.ConstruirContextoAsync(...)`, reemplazar el fallback descontrolado (L507-525) que traía 46 asignaturas de todos los semestres.
  2. Acotar la consulta SQL `sqlAsig` al programa y a la equivalencia semestral:
     `AND (@Semestre = '' OR semestre = @Semestre OR (@SemestreArabigo <> '' AND semestre = @SemestreArabigo))`
  3. **Corrección F400:** La tupla de demanda se ancla al semestre y grupo seleccionados para la corrida (`ctx.IdAulaCorrida`), asegurando que la demanda corresponda exactamente al aula en optimización.
  4. Cargar en `ctx.AulasAlcance` la totalidad de aulas del período/sede para cálculo de capacidad global en pre-vuelo.

#### 3. Fase 380 / Fase 396 / Fase 401 — Pre-vuelo: Oferta por Ocurrencias Reales y Factibilidad de Asignación
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorTemporal.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorVentana.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgGeneradorHorarios.cs`
- **Cambios Técnicos:**
  1. En `AgPlanificadorTemporal.cs`, implementar el método auxiliar determinista:
     `public static int ContarOcurrencias(DayOfWeek dia, DateTime ini, DateTime fin, Func<DateTime, bool> esFestivo)`.
  2. **Corrección F396:** En `AgValidadorVentana.cs`, al iterar franjas activas, invocar `b.DiaSemana` (que traduce correctamente Domingo=7 a `DayOfWeek.Sunday=0`) evitando el desbordamiento de `(DayOfWeek)b.DiaSemanaNumero`.
  3. Calcular la oferta real del semestre multiplicando la duración de cada franja por sus ocurrencias efectivas no festivas en la ventana $[\alpha_{ini}, \alpha_{fin}]$.
  4. Validar factibilidad por asignación mediante el algoritmo de fechas reales de F391. Si alguna asignatura requiere más sesiones de las que caben en su ventana y patrón de días, registrar bloqueo tipificado.
  5. **Corrección F401:** Se descarta `R-DOC` por día como condición bloqueante (falso positivo, ya que un docente puede tener dos franjas distintas el mismo día como Sábado mañana/tarde). La validación de colisiones docentes se mantiene a nivel de franja exacta R1 en `AgValidadorReglas`.

#### 4. Fase 381 / Fase 391 / Fase 392 / Fase 393 — Planificador de Patrón Semanal (`AgPlanificadorPatronSemanal`)
- **Archivos Nuevos y Afectados:**
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPatronSemanal.cs` (Nuevo tipo aislado)
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorPatronSemanal.cs` (Nuevo servicio)
  - `ConversorPDF.XLL/ConversorPDF.XLL.csproj` (Registro de compilación)
- **Cambios Técnicos:**
  1. `AgPatronSemanal.cs`: Clase sellada con propiedades de solo lectura `DiasSemana` (`IReadOnlyList<int>`), `SesionesPorSemana` (`IReadOnlyList<int>`), `FechasValidas` (`IReadOnlyList<DateTime>`). Constructor explícito sin Primary Constructors (regla C# 7.3).
  2. **Corrección F391:** En `AgPlanificadorPatronSemanal.cs`, descartar fórmulas de división cerrada. Implementar `GenerarFechas(...)` recorriendo día a día el calendario desde `FechaInicio` hasta `FechaFin`, verificando pertenencia a los días $K$ del patrón y exclusión de festivos hasta completar $S = D / 4$ sesiones.
  3. Derivar `SesionesPorSemana` agrupando las fechas resultantes por semana del año ISO 8601. Este recorrido genera de forma natural y exacta los 13 histogramas de $H9$ (incluidos patrones escalonados como `2,3,3,4,4` y `1,2,3,3,3`).
  4. Respetar estrictamente los días declarados $K$ (invariante $H6$). Si la ventana es insuficiente con los días dados, retornar `Result.Failure` en lugar de agregar días arbitrariamente.

#### 5. Fase 382 / Fase 394 / Fase 395 / Fase 397 / Fase 398 / Fase 402 / Fase 403 — Multi-Día 1↔N Genes, Invariantes y Migración V55
- **Archivos Afectados y Creados:**
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV55_AddAgGenPatronSemanal.cs` (Nuevo)
  - `ConversorPDF.XLL/Core/Helpers/SqliteHelper.cs` (Registro en migrador)
  - `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgGen.cs`
  - `ConversorPDF.XLL/Academic/Domain/Models/AgAulaAsignatura.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgExpansorSesiones.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgMaterializadorSesiones.cs`
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgScheduleRepository.cs`
- **Cambios Técnicos:**
  1. **Migración V55 (F402):**
     - Añadir columnas a `ag_gen`: `indice_dia_patron INTEGER NOT NULL DEFAULT 0`, `patron_semanal TEXT NOT NULL DEFAULT ''`, `dias_patron TEXT NOT NULL DEFAULT ''`, `clave_grupo TEXT NOT NULL DEFAULT ''`, `horas_gen REAL NOT NULL DEFAULT 4.0`.
     - Añadir columna a `ag_aula_asignatura`: `dias_patron TEXT NOT NULL DEFAULT ''` (almacena días ISO como `"4,5"`).
     - **Corrección F395:** Sembrar franja de Lunes en `ag_bloque_horario` (`dia_semana=1, hora_inicio='18:00', hora_fin='22:00', duracion_horas=4.0, turno='NOCHE', activo=1`) si no existe, garantizando que patrones con Lunes no lancen excepción.
     - Implementación con verificación preventiva `PRAGMA table_info` e idempotencia.
  2. **Modelo `AgGen.cs` (F397 & F398):**
     - Incorporar `IndiceDiaPatron`, `PatronSemanal`, `DiasPatron`, `ClaveGrupo` y `HorasGen`.
     - Preservar `TotalHorasAsignatura` como las horas globales de la asignatura (invariante para métricas y R6). `HorasGen` almacena las horas correspondientes a este gen específico dentro del patrón.
     - `ClaveGrupo` determinista: `${IdAsignatura}|${IdAula}|${Grupo}` (inmune a cruces evolutivos).
     - Actualizar `Clonar()` copiando los nuevos campos.
  3. **Fábrica `AgFabricaIndividuos.cs`:**
     - Leer `dias_patron` desde `AgAulaAsignatura` (F394).
     - Por cada asignatura, si posee patrón multi-día de $N$ días, generar $N$ genes vinculados bajo la misma `ClaveGrupo`, asignando a cada uno la franja horaria correspondiente a su día y sus `HorasGen`.
  4. **Expansor `AgExpansorSesiones.cs`:**
     - Expandir fechas utilizando `gen.HorasGen` si el gen forma parte de un patrón.
     - Aserción de integridad: horas acumuladas en sesiones igual a `gen.HorasGen`.
  5. **Materializador `AgMaterializadorSesiones.cs` (F403):**
     - Agrupar genes por `ClaveGrupo` para sintetizar la vinculación.
     - Formatear `DiaHorasSemanales` en la convención estándar no disruptiva: `"DIA|horas, DIA|horas"` (ej: `"JUEVES|4, VIERNES|4"`), preservando compatibilidad con los consumidores existentes.
     - Validar que `Σ gen.HorasGen == agVinc.TotalHoras == asignatura.IntHoraria` (Invariante I1).
  6. **Repositorio `AgScheduleRepository.cs`:**
     - Actualizar sentencias de inserción y lectura en `SaveIndividuoAsync` y consultas de staging para mapear las columnas añadidas en V55.

#### 6. Fase 383 / Fase 404 — Materialización Verificada y Manejo de Errores
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
  - `ConversorPDF.XLL/Academic/Domain/Enums/AgCausaFallo.cs`
- **Cambios Técnicos:**
  1. En `AgCausaFallo.cs`, incorporar los valores canónicos `SinAsignaturas` y `StagingVacio`.
  2. En `AgFormGeneradorHorarios.cs`, tras la invocación a `MaterializarEnStagingAsync`:
     - Evaluar explícitamente `matRes.IsFailure`. Si falla, registrar traza de fallo y notificar vía `MessageBox`.
     - **Corrección F404:** Consultar staging y verificar defensivamente `ses.IsSuccess && vin.IsSuccess` antes de acceder a `.Value`. Si alguna consulta falla o retorna 0 elementos, emitir alerta de inconsistencia y evitar presentar métricas erróneas.

#### 7. Fase 384 / Fase 407 — Pool por Unión (Semestre ∪ General) sin Fallback Condicional
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
  - `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgContexto.cs`
- **Cambios Técnicos:**
  1. En `AgFormGeneradorHorarios.cs`, unificar la carga de docentes en una única consulta SQL:
     `SELECT ... FROM terceros_asignaturas WHERE activo = 1 AND idprograma = @IdPrograma AND (@Semestre = '' OR semestre = @Semestre OR (@SemestreArabigo <> '' AND semestre = @SemestreArabigo) OR semestre = '' OR semestre IS NULL) AND (subunidad = @Sub OR subunidad = '' OR subunidad IS NULL)`
  2. Eliminar completamente el bloque de fallback condicional `if (ctx.PoolDocentes.Count == 0)`.
  3. En `AgContexto.ConstruirIndices()`, construir `PoolPorAsignatura` ordenando los candidatos de forma prioritaria: primero aquellos asignados específicamente al semestre actual, seguidos por prioridad y menor carga de asignaciones anteriores.

#### 8. Fase 385 — Identificación Canónica de Programas Homónimos y Aislamiento por Sede
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Cambios Técnicos:**
  1. Resolver programas en el selector comparando por su identificador único y nombre normalizado, alertando en bitácora si existen homónimos entre distintas modalidades o planes (ej. APT-52 vs APT-53).
  2. Delimitar estrictamente la oferta y demanda a la sede física seleccionada, evitando agregaciones espurias de sedes territoriales distintas.

#### 9. Fase 386 / Fase 406 — Ventana Semestral $\alpha$ y Cobertura Temporal Completa (H14)
- **Archivos Afectados / Creados:**
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV56_AjustarInicioPeriodo2026_02.cs` (Nuevo)
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgResolutorVentanaSemestre.cs`
- **Cambios Técnicos:**
  1. **Corrección F406 (Opción B):** Implementar migración de datos V56 que actualiza `periodo_inicio = '2026-08-11'` para el período `2026-02` (idperiodo = 2), garantizando que la ventana formal abarque desde el primer martes de clases observado en `Horarios.xlsx` (11-ago-2026) hasta el 12-dic-2026 (17 semanas completas).
  2. En `AgResolutorVentanaSemestre.cs`, asegurar que la prioridad 5 (`Periodos`) tome este rango y registre en bitácora el inicio, fin y semanas activas tras descontar festivos.

#### 10. Fase 387 / Fase 405 — Saneamiento y Honestidad Visual en UI
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Cambios Técnicos:**
  1. **Corrección F405:** En `CargarResultados(...)`, utilizar los identificadores reales de controles WinForms:
     - `dgvHorariosGenerados.DataSource = null;`
     - `btnPromoverProduccion.Enabled = false;`
     - `lblMejorFitness.Text = lblVacantes.Text = lblHardViolaciones.Text = "—";`
  2. Si la corrida no produjo sesiones válidas en staging, limpiar completamente el panel y advertir al usuario en consola, impidiendo que queden residuos visuales de optimizaciones previas.

#### 11. Fase 388 — Generación Asistida de Aulas Congruente
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Cambios Técnicos:**
  1. En `GenerarEncabezadosAulasAsync(...)`, aplicar la misma regla de demanda acotada de F379 al poblar `ag_aula_asignatura`.
  2. Asignar `horas_objetivo = asignaturas.int_horaria` y almacenar el patrón inicial en `dias_patron`, garantizando congruencia absoluta entre el catálogo de aulas y los requerimientos curriculares.

#### 12. Fase 389 / Fase 392 — Suite de Pruebas Unitarias de Regresión y Paridad
- **Archivos Afectados y Creados en `ConversorPDF.Tests`:**
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgFormatoAsignaturaTests.cs` (Extensión con casos de números arábigos y parseo con guiones)
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgPlanificadorPatronSemanalTests.cs` (Nuevo: validación de los 13 histogramas de $H9$ usando calendario real)
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgValidadorVentanaTests.cs` (Reescritura de oferta por ocurrencias y descarte de falso positivo `R-DOC`)
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/PoolDocenteUnionTests.cs` (Nuevo: verificación de cobertura docente con pool unificado)
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgMultiDiaGenTests.cs` (Nuevo: verificación de 1↔N genes, suma de horas e invariante I1)
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgMaterializadorVerificadoTests.cs` (Nuevo: verificación de guardas de staging)
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/Fixtures/Horarios2026_02Fixture.cs` (Fixture canónico con filas representativas del archivo `Horarios.xlsx`)

#### 13. Fase 390 — Exportación Espejo 1:1 con Layout de `Horarios.xlsx`
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.Designer.cs`
- **Cambios Técnicos:**
  1. Incorporar botón `btnExportarProyeccion` en el panel de resultados.
  2. Implementar la generación de la hoja de cálculo con la estructura espejo exacta:
     - Columnas A a K: Metadatos (`No.`, `Programa`, `Sem.`, `Horas`, `Asignatura`, `Sede`, `Residencia`, `Docente`, `Fecha Inicio`, `Fecha Final`, `Días/Horas Semana`).
     - Columnas L a DZ: 119 columnas-día consecutivas rotando fechas y etiquetas de día (`Ma, Mi, J, V, S, D, L`).
     - Columna EA: `SubTotal` (fórmula suma o cálculo idéntico).
  3. Envolver todo el proceso de exportación dentro de un bloque `using (ExcelOptimizeHelper.Optimize(app))` para maximizar el rendimiento y garantizar la liberación determinista de objetos COM.

### 157.5 Resolución Integral de Ronda 52 — Correcciones Críticas y Paridad Funcional (Fases 409 - 422)

Evaluación exhaustiva de los 14 hallazgos detectados en la revisión estática v14, validando los 2 defectos bloqueantes de compilación/migración y las 12 mejoras funcionales de preservación de invariantes, paridad de histogramas y consistencia de alcances.

#### 1. Fase 409 — Corrección de Error de Compilación CS1061 en Materialización
- **Archivo:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgMaterializadorSesiones.cs`
- **Diagnóstico:** `grp` es de tipo `IGrouping<string, AgGen>`, por lo que `grp.Key` es un `string` (`ClaveGrupo`), no un objeto compuesto. El acceso a `grp.Key.IdAula` arrojaba error de compilación CS1061.
- **Solución:** Se extrajo `int idAulaGrupo = genes.Count > 0 ? genes[0].IdAula : 0;` del primer gen ordenado por `IndiceDiaPatron`, empleando `idAulaGrupo` de forma segura para resolver la sede y emitir el mensaje de fallo.

#### 2. Fase 410 — Corrección de Columna 'franja' en Migración V55
- **Archivo:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV55_AddAgGenPatronSemanal.cs`
- **Diagnóstico:** El seed para la franja de los lunes utilizaba la columna inexistente `turno` en `INSERT INTO ag_bloque_horario`, ocasionando fallo de SQLite en el arranque (`no such column: turno`).
- **Solución:** Se corrigió el nombre de la columna a `franja`, en estricta conformidad con el esquema canónico definido en `MigrationV49_AddAgHorariosTables.cs`.

#### 3. Fase 411 — Derivación e Inserción de 'dias_patron' en Encabezados de Aulas
- **Archivo:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** La inserción en `ag_aula_asignatura` pasaba una cadena vacía constante `''` para `dias_patron`, provocando que el validador y la fábrica operaran a ciegas sin datos de multi-día.
- **Solución:** Se implementó la regla de derivación empírica según carga horaria:
  - $\ge 80$ h $\rightarrow$ `"2,3,4,5,6"` (5 días)
  - $\ge 64$ h $\rightarrow$ `"4,5,6"` (Jueves, Viernes, Sábado)
  - $< 64$ h $\rightarrow$ `"2,3"` (Martes, Miércoles)
  Pasando el parámetro `@DiasPatron` en la sentencia Dapper `INSERT OR IGNORE`.

#### 4. Fase 412 — Degradación Defensiva a Día Único ante Falla de Patrón
- **Archivo:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`
- **Diagnóstico:** Si `_planificadorPatron.Construir` retornaba `Failure`, `patron` era nulo, la propiedad `PatronSemanal` quedaba vacía y `SesionesDelGen()` evaluaba las horas totales completas para cada uno de los `diasK.Count` genes, multiplicando artificialmente la carga horaria por el número de días.
- **Solución:** Si `resPatron.IsFailure`, se degrada inmediatamente a día único: `diasK = new List<int> { diasK[0] };`, asegurando que se genere un solo gen con la carga total sin multiplicar las horas.

#### 5. Fase 413 — Reproducción Fiel de Histogramas Escalonados Empíricos
- **Archivo:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorPatronSemanal.cs`
- **Diagnóstico:** El recorrido día a día puro producía histogramas uniformes, sin reproducir los patrones escalonados observados en `Horarios.xlsx` (`1-2-3-3-3`, `2-3-3-4-4`, `3-4-4-5`, `1-2-2-2`).
- **Solución:** Se implementó la distribución por semanas con rampa empírica: si `semanas >= 4` y `baseD > 1`, se decrementa la semana 0 (efecto arranque) y se compensa en las semanas posteriores respetando el tope de 5 días por semana. Se emplearon primitivas compatibles con .NET Framework 4.8 (`Clamp`, `Math.Max`, `Math.Min`).

#### 6. Fase 414 — Asignación y Consumo de 'FechasGen' en Expansión de Sesiones
- **Archivos:**
  - `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgGen.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgExpansorSesiones.cs`
- **Diagnóstico:** `AgExpansorSesiones` regeneraba fechas desde `FechaInicio`, concentrando las sesiones al inicio del semestre y desincronizándose del histograma del patrón.
- **Solución:**
  1. Se añadió `public List<DateTime> FechasGen { get; set; } = new List<DateTime>();` a `AgGen` (con soporte en `Clonar()`).
  2. La fábrica filtra y asigna las fechas precalculadas por el planificador a cada gen según su día ISO.
  3. `AgExpansorSesiones` consume prioritariamente `gen.FechasGen` si está poblado antes de invocar `GenerarFechasValidas`.

#### 7. Fase 415 — Agrupación por Semanas Relativas Ordenadas
- **Archivo:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorPatronSemanal.cs`
- **Diagnóstico:** El cálculo previo con `GetWeekOfYear` sin ordenar rompía la secuencia ante semanas que cruzan fin de año.
- **Solución:** Se computan semanas ordinales relativas a la fecha de inicio del periodo semestral mediante `(f.Date - fechaInicio.Date).TotalDays / 7.0` con ordenamiento determinista por clave.

#### 8. Fase 416 — Cálculo Protegido y Verificación de Invariante I1 en Fábrica
- **Archivo:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`
- **Diagnóstico:** `totalHoras / diasK.Count` carecía de protección contra división por cero y no verificaba que la suma de horas de los genes coincidiera con la carga de la asignatura.
- **Solución:** Se protegió con `Math.Max(1, diasK.Count)` y se añadió la verificación post-bucle $\sum \text{HorasGen} == \text{totalHoras}$ emitiendo advertencia `[I1]` ante discrepancias.

#### 9. Fase 417 — Validación de Invariantes I1 e I2 en Materialización de Staging
- **Archivo:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgMaterializadorSesiones.cs`
- **Diagnóstico:** La persistencia en staging no validaba matemáticamente las horas ni las sesiones agrupadas antes de insertar.
- **Solución:** Se agregaron compuertas estrictas:
  - Invariante I1: `Math.Abs(horasGrupo - esperado) <= 0.001` (retorna `Result.Failure` si no coincide).
  - Invariante I2: `Math.Abs(sesionesGrupo * 4.0 - esperado) <= 0.001` (retorna `Result.Failure` si las sesiones expandidas difieren de `TotalHorasAsignatura / 4`).

#### 10. Fase 418 — Homogeneización de Alcance de Oferta vs Demanda en Validador
- **Archivo:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorVentana.cs`
- **Diagnóstico:** La oferta computada sobre `AulasAlcance` sumaba todas las aulas del programa/sede sin restringir al semestre y grupo de la demanda evaluada, distorsionando la comparación.
- **Solución:** Se alineó `aulasOferta = aulasDisponibles`, evaluando la capacidad en el mismo alcance curricular que la demanda.

#### 11. Fase 419 — Unificación de Valor por Defecto de 'diasDeclarados'
- **Archivos:**
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorVentana.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`
- **Diagnóstico:** El validador usaba 2 días fijos mientras la fábrica caía a 1 gen si `DiasPatron` venía vacío.
- **Solución:** Se estandarizó la fuente única: si `DiasPatron` no está configurado, ambos componentes adoptan por defecto 2 días (`"2,3"`: Martes y Miércoles).

#### 12. Fase 420 — Limpieza de Deuda Técnica y Visibilidad en 'AgContexto'
- **Archivo:** `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgContexto.cs`
- **Diagnóstico:** Presencia de nombres de tipo completamente cualificados en llamadas a `AgFormatoAsignatura` y setter público no restringido en `PoolPorAsignatura`.
- **Solución:** Se incorporó el `using ConversorPDF.XLL.Academic.Application.Services.Genetica;` y se restableció `public Dictionary<int, List<TerceroAsignatura>> PoolPorAsignatura { get; private set; }`.

#### 13. Fase 421 — Exportación Espejo 1:1 con Seriales OADate y Título A1
- **Archivo:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** El exportador emitía fechas en formato de texto plano `yyyy-MM-dd` sin fila de título de periodo.
- **Solución:** Se añadió el título A1 `"PROGRAMACION DOCENTE PERIODO 2026-02"` y se formatearon las 119 columnas-día utilizando números de serie OLE Automation (`f.ToOADate()`), asegurando plena interoperabilidad matemática y visual con Excel.

#### 14. Fase 422 — Pruebas Unitarias para Histogramas Escalonados e Invariante I1
- **Archivos:**
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgPlanificadorPatronSemanalTests.cs`
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgFabricaIndividuosTests.cs`
- **Diagnóstico:** La suite de pruebas no validaba las secuencias escalonadas observadas en `Horarios.xlsx` ni la conservación de suma de `HorasGen`.
- **Solución:** Se añadieron pruebas unitarias para:
  - 64h / 2 días / 5 sem $\rightarrow$ `[3, 4, 4, 5]`
  - 64h / 4 días / 5 sem $\rightarrow$ `[2, 3, 3, 4, 4]`
  - 48h / 3 días / 5 sem $\rightarrow$ `[1, 2, 3, 3, 3]`
  - 28h / 1 día / 4 sem $\rightarrow$ `[1, 2, 2, 2]`
  - Preservación estricta de $\sum \text{HorasGen} == \text{TotalHorasAsignatura}$ en soluciones generadas por `AgFabricaIndividuos`.

### 157.6 Ronda 53 — Calibración Matemática de Escalado e Invariantes (Fases 423 - 429)

Resolución de los 7 hallazgos detectados en la verificación del escalado empírico y blindaje de invariantes formales:

#### 1. Fase 423 — Conservación Estricta $\sum \text{sesiones} == S$
- **Archivo:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorPatronSemanal.cs`
- **Diagnóstico:** Si todas las semanas alcanzaban el tope (`TopeDiasSemana = 5`), la compensación de decremento en semana 0 no encontraba receptor, perdiendo 1 sesión (ej. 80h $\rightarrow$ 19 sesiones en vez de 20).
- **Solución:** Guarda estricta condicional y verificación final `if (porSemana.Sum() != S) return Result.Failure<AgPatronSemanal>(...)`, garantizando cumplimiento exacto de la carga requerida.

#### 2. Fase 424 — Reparto Monótono Creciente por Redondeo Acumulado
- **Archivos:**
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorPatronSemanal.cs`
  - `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgPlanificadorPatronSemanalTests.cs`
- **Diagnóstico:** Los desplazamientos unitarios heurísticos generaban vectores discordantes con aserciones exactas rígidas.
- **Solución:** Implementación de reparto acumulado telescópico $\lceil S \times w / W \rceil - \lceil S \times (w-1) / W \rceil$, garantizando monotonicidad no decreciente y $\sum = S$. En pruebas unitarias, se relajó a comprobación de invariantes fundamentales (`Sum() == S`, `Max() <= nCap`, orden ascendente).

#### 3. Fase 425 — Preservación de Patrones Uniformes Canónicos
- **Archivo:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorPatronSemanal.cs`
- **Diagnóstico:** El escalado forzado alteraba innecesariamente las 197 filas (90%) del Excel que presentan distribuciones perfectamente uniformes (`[2,2,2,2,2,2]`, `[2,2,2,2]`).
- **Solución:** El escalado solo se activa si el patrón uniforme no cabe en la ventana disponible `W` (`(int)Math.Ceiling(S / (double)baseD) > W`).

#### 4. Fase 426 — Integración de Festivos en Cálculo de 'SemanasDisponibles'
- **Archivo:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorPatronSemanal.cs`
- **Diagnóstico:** El parámetro `esFestivo` no era consumido en `SemanasDisponibles`, calculando semanas sobre días inhábiles e inflando la oferta disponible.
- **Solución:** Se evalúan iterativamente bloques de 7 días verificando la existencia de al menos un día de `diasK` no festivo en la ventana.

#### 5. Fase 427 — Restricción de Días a 'diasK' Declarados
- **Archivo:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorPatronSemanal.cs`
- **Diagnóstico:** Se incorporaban días no configurados por la institución cuando `n > diasK.Count`.
- **Solución:** `nCap` se limita estrictamente a `Math.Min(diasK.Count, TopeDiasSemana)`, y la lista de días se acota exclusivamente a los elementos de `diasK`.

#### 6. Fase 428 — Penalización Dura y Vacante ante Inconsistencia I1 en Fábrica
- **Archivo:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`
- **Diagnóstico:** La discrepancia en $\sum \text{HorasGen} \neq \text{totalHoras}$ solo emitía un log informativo sin penalizar al cromosoma.
- **Solución:** Se incrementa `individuo.HardViolations++` y se marcan los genes del grupo con `EsVacante = 1` y motivo `SIN_DISPONIBILIDAD`, asegurando que soluciones incongruentes sean descartadas por la presión de selección.

#### 7. Fase 429 — Normalización Canónica de Horas Semanales en Día Único
- **Archivo:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`
- **Diagnóstico:** En la rama de día único se conservaba `asig.HorasSemanales` (2 a 6h), provocando que `SesionesDelGen()` calculara un número incorrecto de sesiones.
- **Solución:** Se fijó `HorasSemanales = 4.0;` e `HorasGen = totalHoras;`, manteniendo la invariante de sesiones atómicas de 4 horas.

### 157.7 Ronda 54 — Estabilización de Orden Monótono y Ajuste Constructivo I1 (Fases 430 - 433)

Resolución de los 4 hallazgos residuales detectados en la verificación del orden de vectores y el balance de horas:

#### 1. Fase 430 — Orden Ascendente Estable en Reparto Escalado
- **Archivo:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorPatronSemanal.cs`
- **Diagnóstico:** El cálculo telescópico de diferencias generaba secuencias en orden decreciente (`[2,2,2,1]`), exigiendo la mayor concentración de sesiones al inicio del semestre.
- **Solución:** Se aplicó `porSemana.Sort()`, garantizando orden ascendente monótono estable, preservando idénticamente la suma $\sum = S$ y validando que cada elemento pertenezca al rango $[1..nCap]$.

#### 2. Fase 431 — Reporte Estricto de Semanas Disponibles sin Enmascarar 0
- **Archivo:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorPatronSemanal.cs`
- **Diagnóstico:** `SemanasDisponibles` retornaba `Math.Max(1, semanas)`, ocultando ventanas sin ningún día hábil y permitiendo aprobaciones falsas en pre-vuelo.
- **Solución:** Retorno directo de `semanas` (permitiendo 0) e inserción de guarda en `Construir`: si $W = 0$, retorna `Result.Failure` inmediato informando que la ventana carece de días hábiles.

#### 3. Fase 432 — Ajuste Constructivo del Residuo en Gen 0 sin Falsas Penalizaciones Duras
- **Archivo:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`
- **Diagnóstico:** Incrementar `HardViolations++` por ligeras discrepancias de redondeo descalificaba el individuo completo (`EsFactible = false`), impidiendo su promoción en el Módulo R.
- **Solución:** Se eliminó el incremento de `HardViolations` y se ajustó el residuo directamente en `g0.HorasGen += (totalHoras - sumaGenes)`, satisfaciendo constructivamente el invariante I1 de forma exacta.

#### 4. Fase 433 — Diagnóstico Accionable ante Insuficiencia de Franjas por Festivos
- **Archivo:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorPatronSemanal.cs`
- **Diagnóstico:** El mensaje de fallo al no alcanzar la cuota semanal era genérico sin indicar la acción requerida.
- **Solución:** Mensaje enriquecido detallando causa y acción correctiva: `"Infactible: semana n sólo x/y sesiones (festivos o días K no disponibles). Ajuste ventana o agregue franjas."`.

### 157.8 Ronda 55 — Sincronización Pre-Vuelo, Múltiplo Canónico de 4 y Diversificación de Franjas (Fases 434 - 437)

Resolución de los 4 hallazgos detectados en la auditoría de paridad del validador de ventana, cuantización de sesiones y asignación de franjas:

#### 1. Fase 434 — Sincronización Pre-Vuelo vs Planificador en Semanas Disponibles
- **Archivo:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorVentana.cs`
- **Diagnóstico:** `AgValidadorVentana` calculaba semanas calendario planas `Math.Max(1, (int)Math.Ceiling((fin - ini).TotalDays / 7.0))` ignorando festivos y días hábiles del patrón. Esto causaba falsos positivos en pre-vuelo (aprobaba la ventana) y bloqueos no tipificados durante la corrida del algoritmo genético (`Failure("Infactible: semana n...")`).
- **Solución:** Se delegó el cómputo de semanas en pre-vuelo a `AgPlanificadorPatronSemanal.SemanasDisponibles(diasKAsig, ini, fin, contexto.EsFestivo)`. Si retorna 0 semanas hábiles, se emite inmediatamente un bloqueo tipificado en el informe de pre-vuelo: `"Asignatura '{asig.Nombre}' (Id={asig.Id}): la ventana no contiene días hábiles para el patrón declarado."`. Además, se sincronizó el cómputo de horas con la jerarquía completa de la fábrica (`HorasObjetivo -> IntHoraria -> Horas -> Creditos*16 -> 64`).

#### 2. Fase 435 — Cuantización de Horas a Múltiplos Canónicos de 4 (Invariantes I1 e I2)
- **Archivo:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`
- **Diagnóstico:** Si una asignatura tenía horas no divisibles exactamente por 4 (p. ej. 45 o 65 h), el número de sesiones se calculaba con `Math.Ceiling(totalHoras / 4.0)` proyectando sesiones completas de 4 horas, lo que generaba un residuo negativo `totalHoras - (sesiones * 4)` y forzaba sesiones fraccionarias en el expansor, violando la regla del modelo de negocio de Horarios (todas las sesiones en Excel son exactamente de 4.0 horas).
- **Solución:** Antes de instanciar genes o construir el patrón, se normaliza `totalHoras` al múltiplo superior de 4 mediante `Math.Ceiling(totalHoras / 4.0) * 4.0` cuando `Math.Abs(totalHoras % 4.0) > 0.001`, emitiendo una advertencia trazable en `LoggerHelper.Warning` con el tag `[I2]`. Esto garantiza que `totalHoras == sesiones * 4.0` y previene residuos negativos.

#### 3. Fase 436 — Eliminación de Código Muerto 'Clamp'
- **Archivo:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorPatronSemanal.cs`
- **Diagnóstico:** El método auxiliar privado `Clamp` quedó sin referencias tras la refactorización de escalado de la Fase 427, produciendo warnings de compilación / análisis estático.
- **Solución:** Eliminación completa del método no utilizado para mantener el código conciso y libre de advertencias de análisis estático.

#### 4. Fase 437 — Rotación Determinista de Franjas Horarias por Asignatura y Gen
- **Archivo:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`
- **Diagnóstico:** La selección de bloque mediante `bloquesValidos.FirstOrDefault(b => b.DiaSemanaNumero == diaIso)` seleccionaba invariablemente la primera franja del día (en Sábado, siempre la franja MAÑANA 08:00–12:00), provocando colisiones sistemáticas de aula/horario (restricción dura R4) entre asignaturas que comparten el mismo día y reduciendo artificialmente la diversidad de la población inicial.
- **Solución:** Se implementó una selección por rotación determinista `franjasDia[Math.Abs(asig.Id + i) % franjasDia.Count]`, dispersando las asignaturas entre franjas disponibles (MAÑANA / TARDE) tanto en la rama multi-día como en la rama de día único.

### 157.9 Ronda 56 — Respeto de Patrón Declarado, Fuente Única de Días K y Limpieza de AulasAlcance (Fases 438 - 440)

Resolución de los 3 hallazgos residuales detectados en la auditoría de persistencia de patrones, unificación de días de la semana y optimización de consultas:

#### 1. Fase 438 — Respeto de Patrón Declarado en Catálogo frente a Heurística Fija
- **Archivo:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** Al sembrar asignaturas en `ag_aula_asignatura`, se sobrescribía siempre `dias_patron` con una heurística rígida basada en horas ($\ge 80 \rightarrow \text{"2,3,4,5,6"}$, $\ge 64 \rightarrow \text{"4,5,6"}$, $< 64 \rightarrow \text{"2,3"}$), ignorando si la asignatura ya poseía una combinación personalizada de días (en `Horarios.xlsx` coexisten 16 combinaciones distintas).
- **Solución:** Se modificó la consulta `sqlAsigs` para realizar `LEFT JOIN ag_aula_asignatura aa ON aa.id_aula = @IdAula AND aa.id_asignatura = a.id` recuperando `COALESCE(aa.dias_patron, '') AS DiasPatron`. Al asignar `diasPatron`, se evalúa prioritariamente `!string.IsNullOrWhiteSpace(aItem.DiasPatron) ? aItem.DiasPatron : (horas >= 80 ? "2,3,4,5,6" : horas >= 64 ? "4,5,6" : "2,3")`, preservando fielmente los patrones configurados por el usuario o importados desde Excel.

#### 2. Fase 439 — Fuente Única Canónica de Días K (`DiasKDe`)
- **Archivos:** 
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorPatronSemanal.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorVentana.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`
- **Diagnóstico:** `AgValidadorVentana` y `AgFabricaIndividuos` parseaban de forma independiente la cadena `DiasPatron` con código duplicado. Si surgía cualquier discrepancia o valor vacío, el pre-vuelo aprobaba con 2 días mientras la corrida intentaba con 1 día (o viceversa), produciendo divergencias en factibilidad.
- **Solución:** Se centralizó el parseo en el método compartido `AgPlanificadorPatronSemanal.DiasKDe(AgAulaAsignatura aa)`, el cual devuelve de forma inmutable la lista canónica de días ISO $[1..7]$ o el default `{ 2, 3 }`. Tanto el validador de ventana como la fábrica de cromosomas consumen este método único, eliminando 25 líneas de código redundante y garantizando paridad total.

#### 3. Fase 440 — Limpieza de Consulta Redundante de `AulasAlcance`
- **Archivo:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** Tras la unificación de alcances de la Fase 418, el validador opera directamente sobre `aulasDisponibles` (mismo alcance curricular que la demanda). La consulta asíncrona a base de datos `GetAulasPorPeriodoAsync` con semestre y grupo nulos era ejecutada redundantemente solo para poblar `AulasAlcance` y emitir una traza.
- **Solución:** Se eliminó la llamada duplicada a base de datos y se sincronizó `ctx.AulasAlcance = ctx.Aulas != null ? ctx.Aulas.ToList() : new List<AgAula>();`, manteniendo viva la métrica de telemetría sin latencia innecesaria de I/O.

### 157.10 Ronda 57 — Dapper POCO Mapping en .NET 4.8, Supresión de 64h Ficticias y SELECT DISTINCT (Fases 441 - 443)

Resolución del defecto bloqueante de runtime en Dapper y las dos inconsistencias de siembra de asignaturas:

#### 1. Fase 441 — Dapper POCO Mapping y `int_horaria` Anulable en .NET Framework 4.8
- **Archivo:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** `QueryAsync<(int Id, int IntHoraria, string DiasPatron)>` utilizaba `ValueTuple`, el cual no es soportado por el motor de mapeo de Dapper en .NET Framework 4.8 (requiere netstandard2.1+). Adicionalmente, valores NULL en la columna `int_horaria` de la base de datos SQLite provocaban fallos de casteo a tipos por valor no anulables.
- **Solución:** Se implementó la clase auxiliar anidada `private sealed class AsigSeedRow` con propiedades fuertemente tipadas (`int Id`, `int? IntHoraria`, `string DiasPatron`), permitiendo que Dapper realice la deserialización mediante reflexión estándar de CLR sin fallos de casteo ni dependencias de runtime no soportadas.

#### 2. Fase 442 — Supresión de Asignaturas sin Carga Horaria (Eliminación de 64h Ficticias)
- **Archivo:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** Asignar un default de 64 horas a materias con `int_horaria` nula o $\le 0$ creaba 16 sesiones obligatorias (4 semanas $\times$ 4 sesiones) para materias vacías, inflando artificialmente la demanda curricular y provocando cuellos de botella falsos en el validador pre-vuelo.
- **Solución:** Si `!aItem.IntHoraria.HasValue || aItem.IntHoraria.Value <= 0`, la asignatura se omite explícitamente (`omitidas++`), emitiendo una traza de telemetría informativa (`traza.Dato("AulaAsignaturaSinHoras", ...)`). Solo las asignaturas con carga real ingresan a `ag_aula_asignatura`.

#### 3. Fase 443 — `SELECT DISTINCT` para Prevenir Duplicación por LEFT JOIN
- **Archivo:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** Si la tabla `ag_aula_asignatura` contenía registros previos duplicados para el par aula-asignatura, el `LEFT JOIN` multiplicaba las filas de `asigsData`, produciendo sobreconteo en las métricas de inserción.
- **Solución:** Se incorporó `SELECT DISTINCT a.id AS Id, ...` en `sqlAsigs`, garantizando cardinalidad 1:1 por cada asignatura del catálogo académico.

### 157.11 Fase 444 — Desacoplamiento de Métrica 'sinHoras' del Contador 'omitidas'

- **Archivo:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** En la Fase 442, omitir una asignatura por falta de `int_horaria` incrementaba el contador `omitidas++`. Dado que dicho contador representaba históricamente *"Aulas preexistentes omitidas"*, el diálogo final y el registro del log presentaban una métrica sobrecargada y engañosa al usuario.
- **Solución:** Se introdujo la variable independiente `int sinHoras = 0;` para contabilizar exclusivamente las asignaturas omitidas por falta de intensidad horaria. Se desacopló de `omitidas`, registrando ambas magnitudes de forma transparente en `traza.Dato("SinHoras", sinHoras)`, `txtLogAG` y en el `MessageBox` resumen:
  - `- Aulas creadas: {insertadas}`
  - `- Aulas preexistentes omitidas: {omitidas}`
  - `- Asignaturas sin int_horaria omitidas: {sinHoras}`

### 157.12 Ronda 59 + 60 — Coherencia de Períodos, Desambiguación de Cohortes (A/B/C), Regla R8 y Staging Fiel (Fases 445 - 461)

Resolución integral de los defectos detectados en la auditoría de corridas reales (Ronda 59 y 60):

#### 1. Fase 445 / Fase 453 — Inyección de PeriodoObjetivo del Combo en Tupla S y Traza de Período
- **Archivos:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** En `ObtenerCombinacionesSeleccionadas()`, las tuplas S conservaban el `IdPeriodo` asociado a nodos del árbol o perfiles antiguos; si el usuario seleccionaba un período distinto en `cmbPeriodos` (p. ej. `2027-01` id=3), las combinaciones se generaban con el período anterior (id=2), produciendo discordancia temporal y ventanas obsoletas.
- **Solución:** Se capturó `idPeriodoObjetivo = (cmbPeriodos.SelectedItem as Periodo)?.Id ?? 0;` en `ObtenerCombinacionesSeleccionadas()`, asignándolo a cada DTO generado. En `ConstruirContextoAsync`, si `_ventanaResuelta` no coincide con el período del combo, se recalcula la ventana semestral automáticamente. Se añadió emisión de traza estructurada `traza.Dato("PeriodoObjetivo", ...)` y advertencia `MultiplesPeriodosActivos` si existen más de un período con `activo = 1`.

#### 2. Fase 446 — Resolución de Programa Homónimo con Pensum del Semestre Solicitado
- **Archivos:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** En la base de datos coexisten programas con nombres idénticos (p. ej. `ADMINISTRACION PÚBLICA TERRITORIAL`) con IDs distintos (p. ej. 1, 2, 3). Si la selección apuntaba a un ID sin asignaturas para el semestre solicitado (`I`), la optimización fallaba con catálogo vacío a pesar de que el homónimo ID 1 poseía el pensum completo.
- **Solución:** Se implementó detección y resolución de homónimos en `ConstruirContextoAsync` y `GenerarEncabezadosAulasAsync`: si el programa seleccionado carece de asignaturas para el semestre pero un homónimo activo posee registros en `asignaturas`, se conmuta automáticamente al homónimo con el pensum, registrando `traza.Dato("ProgramaHomonimoResuelto", ...)`.

#### 3. Fase 447 — Generación Asistida de Aulas sin Aulas Huérfanas
- **Archivos:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** `GenerarEncabezadosAulasAsync` insertaba la cabecera en `ag_aula` antes de comprobar si existían asignaturas para el programa y semestre. Al encontrar 0 asignaturas en el catálogo, dejaba un aula huérfana en la base de datos sin asignaturas en `ag_aula_asignatura`.
- **Solución:** Se verifica `asigsData.Count > 0` antes de persistir el aula en `ag_aula`. Si no existen asignaturas, se incrementa `sinAsignaturas`, se emite traza diagnóstica de bloqueo y se omite la inserción de la cabecera, evitando registros residuales vacíos.

#### 4. Fase 448 / Fase 449 / Fase 450 / Fase 452 — Propagación de Fallo de Contexto vía Result y Diagnóstico Accionable
- **Archivos:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** `ConstruirContextoAsync` retornaba una instancia no nula de `AgContexto` aun cuando `ctx.Asignaturas.Count == 0` o `ctx.Aulas.Count == 0`, registrando éxito engañoso y fallando en etapas posteriores sin mensaje accionable.
- **Solución:** Se refactorizó la firma a `Task<Result<AgContexto>>`. Si `ctx.Aulas.Count == 0`, se retorna `Result.Failure` instruyendo a ejecutar "Generar Aulas". Si `ctx.Asignaturas.Count == 0`, se retorna `Result.Failure` con el desglose completo del contexto (período, programa, semestre, subunidad y disponibilidad en programas homónimos). Se emite advertencia si el período carece de festivos/presenciales (`SinFestivos`).

#### 5. Fase 454 / Fase 455 — Multi-Aula por Combinación (Cohortes A, B, C) con Modalidad y Grupo de Aula
- **Archivos:** 
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs`
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgMaterializadorSesiones.cs`
- **Diagnóstico:** `AgFabricaIndividuos` iteraba únicamente sobre `Asignaturas` y asignaba todos los genes a `IdAulaCorrida` o al primer aula con `Grupo = "A"`. Al seleccionar grupos A, B y C para el mismo semestre, todos los cursos colapsaban en una única aula (p. ej. 266), generando 14 violaciones duras de superposición de aula (R2) y grupo (R3).
- **Solución:** En `ConstruirContextoAsync`, se cargan las aulas de la sede y período sin acotar a un solo grupo, fijando `IdAulaCorrida` únicamente si hay exactamente 1 combinación. En `AgFabricaIndividuos.CrearAleatorio`, se itera por cada aula/grupo representativo (`gruposAula`), asignando deterministamente `gen.AsignarAula(aula)` con su `Grupo` y `Modalidad` propios. En `AgMaterializadorSesiones`, se valida que todos los genes del grupo compartan el mismo grupo (`agVinc.Grupo == genes[0].Grupo`).

#### 6. Fase 456 — Mutación Segura de Aula/Bloque y Restricción Dura R8 (Días No Duplicados en Multi-Día)
- **Archivos:** 
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgOperadorMutacion.cs`
  - `ConversorPDF.XLL/Academic/Domain/Enums/AgReglaViolacion.cs`
  - `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorReglas.cs`
- **Diagnóstico:** La mutación de aula seleccionaba aulas de cualquier grupo mezclando cohortes; y la mutación de bloque elegía bloques aleatorios de cualquier día de la semana, haciendo que asignaturas multi-día terminaran con días duplicados (p. ej. Viernes y Viernes), violando el patrón curricular.
- **Solución:**
  1. La mutación de aula se restringe a aulas con el mismo `Grupo` del gen (`gen.Grupo`).
  2. En genes multi-día, la mutación de bloque solo conmuta franjas horarias pertenecientes al mismo día asignado al gen (`b.DiaSemanaNumero == gen.DiaSemana`), impidiendo días repetidos.
  3. Se incorporaron los códigos `R8_DiasDuplicadosGrupo` y `R8_DiaFueraDePatron` en `AgReglaViolacion` y su evaluación en `AgValidadorReglas`, penalizando como infracción dura cualquier cromosoma con días duplicados dentro de la misma `ClaveGrupo`.

#### 7. Fase 457 — Consulta Fiel de Staging en MODO AG
- **Archivos:**
  - `ConversorPDF.XLL/Academic/Application/DTOs/VinculacionFilterDto.cs`
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/VinculacionQueryRepository.cs`
  - `ConversorPDF.XLL/Academic/UI/Forms/FormVinculaciones.cs`
  - `ConversorPDF.XLL/Core/Helpers/SqliteHelper.cs`
- **Diagnóstico:** Al abrir `FormVinculaciones` en MODO AG desde Módulo R, el formulario aplicaba filtros predeterminados de producción (período de `config` id=1), devolviendo 0 sesiones. Además, `vw_AgSesionesDetalle` utilizaba `INNER JOIN` para programas y sedes y no proyectaba `IdCorrida`.
- **Solución:** Se añadió `IdCorrida` a `VinculacionFilterDto` y filtrado parametrizado en `VinculacionQueryRepository`. Se actualizó la vista `vw_AgSesionesDetalle` en `SqliteHelper.cs` para proyectar `v.id_corrida AS IdCorrida` y usar `LEFT JOIN` defensivos. En MODO AG, `FormVinculaciones` preselecciona la última corrida activa.

#### 8. Fase 458 — Coherencia Estricta de Periodo y Fechas en Staging
- **Archivos:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgMaterializadorSesiones.cs`
- **Diagnóstico:** `ag_vinculaciones` persistía fechas iniciales y períodos discordantes con la ventana semestral del período objetivo de la corrida.
- **Solución:** En `AgMaterializadorSesiones`, si `contexto.Ventana` está resuelta con `IdPeriodo > 0`, se valida que `agVinc.IdPeriodo == contexto.Ventana.IdPeriodo` (o se asigna), y se fuerzan `agVinc.FechaInicio` y `agVinc.FechaFin` desde la ventana semestral objetivo.

#### 9. Fase 459 — Migración V57: Reparación de Esquema en 'salvedades_conflictos'
- **Archivos:**
  - `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV57_RepairSalvedadesConflictos.cs` (Nuevo)
  - `ConversorPDF.XLL/Core/Helpers/SqliteHelper.cs`
  - `ConversorPDF.XLL/ConversorPDF.XLL.csproj`
- **Diagnóstico:** En vaults migrados desde versiones tempranas, la tabla `salvedades_conflictos` carecía de las columnas `idsesion_a` e `idsesion_b`, provocando la excepción `no such column: idsesion_a` al consultar o registrar salvedades en `SalvedadConflictoRepository`.
- **Solución:** Se implementó `MigrationV57_RepairSalvedadesConflictos` con `IMigrationTolerante`, agregando de forma idempotente las columnas ausentes (`idsesion_a`, `idsesion_b`, `idtercero`, `idperiodo`, `justificacion`) mediante `ALTER TABLE ... ADD COLUMN` si no existen.

#### 10. Fase 460 / Fase 461 — Fiel Desglose de Violaciones HARD/SOFT y Trazabilidad en Grid
- **Archivos:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **Diagnóstico:** El contador de violaciones en UI mostraba un resumen incompleto y el grid de horarios staging no proyectaba los identificadores de trazabilidad (`IdCorrida`, `IdVinculacion`, `IdGen`).
- **Solución:** En `CargarModuloResultadosAsync`, se actualizó `lblHardViolaciones` para reportar el conteo total: `Violaciones: {total} (HARD {hard}, SOFT {soft})`, ordenando `dgvViolaciones` con prioridad descendente para las restricciones duras. Se añadieron las columnas `IdCorrida`, `IdVinculacion`, `IdGen` e `IdSesionGlobal` a la proyección de la grilla de resultados.

