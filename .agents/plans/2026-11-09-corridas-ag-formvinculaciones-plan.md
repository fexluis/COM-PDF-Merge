# 📄 Plan R-007 — Combo "Corrida" en `FormVinculaciones` (MODO AG) · Ola A

> **Guardar como**: `.agents/plans/2026-11-09-corridas-ag-formvinculaciones-plan.md`
> **Reservar antes de ejecutar**: fila `R-007` en `.agents/plans/REGISTRY.md` (append-only) y contadores §2 → `ronda 282`, `fase 1488`.

```yaml
numbering:
  registry: .agents/plans/REGISTRY.md
  registry_row: R-007
  ronda_base: 282          # Wave 1 = Ronda 282
  fase_base: 1488          # T1 = Fase 1488
  waves: 3                 # Ronda 282–284
  tasks: 8                 # Fase 1488–1495
```

## Parallel execution groups (waves)

| Wave | Ronda | Tasks |
|---|---|---|
| 1 | 282 | T1, T2, T3 |
| 2 | 283 | T4, T5, T6, T7 |
| 3 | 284 | T8 |

## Execution envelope
- **branch**: `refactor/academic-modularization`
- **workspace**: `c:\Users\USUARIO\OneDrive - Escuela Superior de Administración Publica\Antigravity\ConversorPDFv1.6.3.72`
- **preflight**: `powershell -ExecutionPolicy Bypass -File scripts/preflight.ps1 -Workspace . -Plan .agents/plans/2026-11-09-corridas-ag-formvinculaciones-plan.md -AllowOneDrive`
- **baseline**: `tests/baseline/test_baseline.txt` (TOTAL 952) · **no modificar el baseline en Fase B** (lo actualiza Fase C)
- **build_gate**: **serializado** (Excel-DNA bloquea `bin/x64/Debug/*.xll` → `DNA-577853131`)
- **max_parallel_writers**: 2
- **hot_paths_L0** (solo `integrator`): `ConversorPDF.XLL/ConversorPDF.XLL.csproj`
- **no_push**: `true` · **stop_conditions**: acción destructiva sin respaldo · tocar el vault `.enc` · `DELETE/UPDATE` sobre `ag_*` · push/merge · editar `APP_VERSION` o `<Version>`

---

## 0. Hechos verificados (obligatorio leer antes de codificar)

Diagnóstico ejecutado sobre el vault real `AcademicData_VALLE.db.enc` (vía la copia descifrada temporal del propio add-in, `Read Only=True`), el **09/11/2026**:

| Dato | Valor |
|---|---|
| `config` | `id_periodo = 2` (2026-02), `subunidad = 'VALLE'` |
| `periodos` | 1=2026-01 · **2=2026-02** · **3=2027-01** |
| `ag_corrida` #46 | periodo **3**, programa **20** (MAESTRÍA, POSGRADO), **BOLIVAR**, estado **4 = Error**, **0 vinc / 0 ses** |
| `ag_corrida` #44 | periodo **3**, programa **1** (ADMÓN. PÚBLICA TERRITORIAL), VALLE, estado 2, **23 vinc / 244 ses** |
| `ag_corrida` #37 | periodo 3, programa 2, VALLE, estado 2, **0 / 0** |
| `ag_corrida` #1 | periodo **2**, programa 1, VALLE, estado 2, **18 / 145** |
| Huérfana | `ag_vinculaciones.id_corrida = 45` (BOLIVAR, periodo 3) con 3 sesiones **sin fila en `ag_corrida`** |

**Prueba del bug** (`vw_AgSesionesDetalle`, 43 columnas, incluye `IdCorrida`, `IdPeriodo`, `Subunidad`):

| Consulta | Filas |
|---|---|
| `IdCorrida = 44` | **244** |
| `IdCorrida = 44 AND IdPeriodo = 2` ← valor de `config` | **0** ❌ |
| `IdCorrida = 44 AND IdPeriodo = 3` | 244 |
| `IdCorrida = 44 AND Subunidad = 'VALLE'` | 244 |
| `IdCorrida = 46` ← default del combo (`MAX(id)`) | **0** ❌ |

Log `log_20261109.txt` (09:33:42 → 09:42:12, 6 veces): `[FormVinculaciones] Modo AG: Existen 244 sesiones en staging pero la consulta devolvió 0 filas.`

**Causas que la Ola A debe eliminar**
1. `FormVinculaciones` aplica `config.id_periodo` (2) y `config.subunidad` como filtros **invisibles** (`cmbPeriodo`/`cmbTerritoriales` en `Location(-100,-100)`, `Visible=false`).
2. El combo de corridas elige `MAX(id)` (#46: otra territorial, programa y estado Error) en vez de la corrida en uso (#44).
3. El combo se construye solo desde `ag_corrida` → ofrece corridas vacías (#46, #37) y **oculta** staging huérfano (#45).
4. Los contadores del status usan SQL crudo con otro predicado de estado y sin filtros.

---

## 1. Invariantes y reglas de la Ola A (no negociables)

1. **No** hay migraciones, ni cambios de esquema, ni DDL. Solo `SELECT`.
2. **Nada** de SQL en `Form*` (regla del propio archivo: *"Cero lógica de negocio en el code-behind"*). Todo SQL vive en `VinculacionQueryRepository`.
3. **MODO PRODUCCIÓN debe quedar intacto**: sin combo de corrida, con los filtros de `config` como hoy (el comportamiento cambia solo cuando `_modo.EsModoAg == true`).
4. En MODO AG, **la corrida es la fuente de verdad** de periodo/subunidad/programa.
5. No se reutiliza `LookupItem` para corridas (`Id` es `int`; `ag_corrida.id` es `long`).
6. No se modifica `_currentConnection` (conexión prestada, `Hotfix 1.2.51.25`): los nuevos repositorios usan `_dbFactory`.
7. `RibbonController.cs` **no se toca**: el nuevo parámetro del constructor es opcional (`= null`).
8. Cada test nuevo debe **fallar primero por la razón esperada** (contrato RED→GREEN de `.agents/rules/verification-commands.md` §2).

---

# Wave 1 · Ronda 282 — Contratos (integrator-first, todo debe compilar)

### T1 · Fase 1488 — Esqueletos + registro en csproj *(integrator, lease L0)*
- **fase**: 1488
- **files_lock**: `ConversorPDF.XLL/Academic/Domain/Models/Genetica/CorridaStagingItem.cs` **(nuevo)**, `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgStagingResumen.cs` **(nuevo)**, `ConversorPDF.XLL/Academic/Application/Services/AgCorridaSelector.cs` **(nuevo)**, `ConversorPDF.XLL/Academic/Application/Services/AgFiltrosPolicy.cs` **(nuevo)**, `ConversorPDF.XLL/ConversorPDF.XLL.csproj`
- **description**: crear los 4 tipos con firmas completas y cuerpos `throw new NotImplementedException();`; añadir los 4 `<Compile Include="..." />` al csproj (están ordenados alfabéticamente: respetar posición).
- **Contratos exactos**:

```csharp
// ConversorPDF.XLL.Academic.Domain.Models.Genetica.CorridaStagingItem
namespace ConversorPDF.XLL.Academic.Domain.Models.Genetica
{
    public sealed class CorridaStagingItem
    {
        public long   IdCorrida          { get; set; }
        public int    IdPeriodo          { get; set; }
        public string PeriodoNombre      { get; set; }
        public int    IdPrograma         { get; set; }
        public string ProgramaNombre     { get; set; }
        public string Subunidad          { get; set; }
        public string FechaInicioCorrida { get; set; }
        public int    EstadoCorrida      { get; set; }
        public int    TotalVinculaciones { get; set; }
        public int    TotalSesiones      { get; set; }
        public bool   TieneStaging       => TotalSesiones > 0;   // computada, Dapper la ignora
        public string Etiqueta           { get; set; }           // la construye AgCorridaSelector en origen
    }
}

// ConversorPDF.XLL.Academic.Domain.Models.Genetica.AgStagingResumen
namespace ConversorPDF.XLL.Academic.Domain.Models.Genetica
{
    public sealed class AgStagingResumen
    {
        public long   IdCorrida          { get; set; }
        public int    SesionesCorrida    { get; set; }   // solo filtro corrida
        public int    SesionesConFiltros { get; set; }   // corrida + filtros de UI
        public int    VincCorrida        { get; set; }   // COUNT(DISTINCT IdVinculacion), solo corrida
        public string PeriodosCorrida    { get; set; }   // "2027-01"
        public string SubunidadesCorrida { get; set; }   // "VALLE"
        public string ProgramasCorrida   { get; set; }   // "ADMÓN. PÚBLICA TERRITORIAL"
    }
}

// ConversorPDF.XLL.Academic.Application.Services.AgCorridaSelector  (estático, PURO, sin WinForms)
public static class AgCorridaSelector
{
    public static string FormatearEtiqueta(CorridaStagingItem c);
    public static CorridaStagingItem CrearTodasLasCorridas();                 // IdCorrida = 0, Etiqueta = "-- Todas las corridas --"
    public static CorridaStagingItem ResolverSeleccionInicial(
        IReadOnlyCollection<CorridaStagingItem> items, long? idCorridaEnUso);
}

// ConversorPDF.XLL.Academic.Application.Services.AgFiltrosPolicy  (estático, PURO)
public static class AgFiltrosPolicy
{
    public static bool AplicarPeriodoConfig(AgModoDatos modo);    // => !modo.EsModoAg
    public static bool AplicarSubunidadConfig(AgModoDatos modo);  // => !modo.EsModoAg
}
```
- **Reglas de `FormatearEtiqueta`** (usar `System.Globalization.CultureInfo.InvariantCulture`):
  - Con staging: `#44 · 2027-01 · ADMÓN. PÚBLICA TERRITORIAL · VALLE · 244 ses`
  - Sin staging: `#46 · 2027-01 · MAESTRÍA… · BOLIVAR · (sin sesiones)`
  - Estado 4: sufijo ` · Error`; estado 3: ` · Cancelada`. Campos vacíos se omiten (sin `·` duplicado).
- **Reglas de `ResolverSeleccionInicial`**: 1) si `idCorridaEnUso.HasValue` y existe en `items` → esa; 2) si no, la primera con `TieneStaging` (los `items` llegan ordenados `IdCorrida DESC`); 3) si ninguna tiene staging → `CrearTodasLasCorridas()`.
- **validation**: `scripts/verify.ps1 -Mode Build` exit 0 (los 4 tipos compilan con `NotImplementedException`).
- **status**: done

### T2 · Fase 1489 — Repositorio: corridas + extracción BuildWhere + resumen *(implementer, L1)*
- **fase**: 1489
- **files_lock**: `Academic/Domain/Interfaces/IVinculacionQueryRepository.cs`, `Academic/Infrastructure/Persistence/Repositories/VinculacionQueryRepository.cs`
- **description**:
  1. Añadir a la interfaz `IVinculacionQueryRepository` (con XML doc):
     ```csharp
     Task<Result<IReadOnlyCollection<CorridaStagingItem>>> GetCorridasStagingAsync(
         int? idPeriodo, int? idPrograma, string subunidad, CancellationToken ct);

     Task<Result<AgStagingResumen>> GetResumenStagingAsync(
         long idCorrida, VinculacionFilterDto filtros, CancellationToken ct);
     ```
  2. **Extraer primero `BuildWhere` estático**:
     Extraer del cuerpo actual de `GetSesionesAsync` (líneas 72‑151) un `private static string BuildWhere(...)` que reciba los 11 filtros + `idCorrida` y devuelva el fragmento `AND ...`, añadiendo los parámetros al `DynamicParameters` recibido. `GetSesionesAsync` debe quedar funcionalmente idéntico.
  3. **SQL de `GetCorridasStagingAsync`** (respetar alias PascalCase; `ORDER BY` al final del compuesto):
     ```sql
     SELECT v.id_corrida AS IdCorrida,
            MAX(v.idperiodo)  AS IdPeriodo,
            COALESCE(MAX(pe.periodo),'') AS PeriodoNombre,
            MAX(v.idprograma) AS IdPrograma,
            COALESCE(MAX(COALESCE(pr.abreviatura, pr.nombreprograma)),'') AS ProgramaNombre,
            COALESCE(MAX(NULLIF(v.subunidad,'')), '') AS Subunidad,
            COALESCE(MAX(c.fecha_inicio_corrida),'')  AS FechaInicioCorrida,
            COALESCE(MAX(c.estado),0) AS EstadoCorrida,
            COUNT(DISTINCT v.id) AS TotalVinculaciones,
            COUNT(s.id)          AS TotalSesiones
     FROM ag_vinculaciones v
     LEFT JOIN ag_sesiones s ON s.idvinculacion = v.id
     LEFT JOIN ag_corrida   c ON c.id = v.id_corrida
     LEFT JOIN periodos     pe ON pe.idperiodo = v.idperiodo
     LEFT JOIN programas    pr ON pr.id = v.idprograma
     WHERE IFNULL(v.estado,0) <> 3
       AND (@idPeriodo  IS NULL OR v.idperiodo  = @idPeriodo)
       AND (@idPrograma IS NULL OR v.idprograma = @idPrograma)
       AND (@subunidad  IS NULL OR UPPER(TRIM(v.subunidad)) = UPPER(TRIM(@subunidad)))
     GROUP BY v.id_corrida
     UNION ALL
     SELECT c2.id, c2.idperiodo, COALESCE(pe2.periodo,''), c2.idprograma,
            COALESCE(COALESCE(pr2.abreviatura, pr2.nombreprograma),''), c2.subunidad,
            c2.fecha_inicio_corrida, c2.estado, 0, 0
     FROM ag_corrida c2
     LEFT JOIN periodos  pe2 ON pe2.idperiodo = c2.idperiodo
     LEFT JOIN programas pr2 ON pr2.id = c2.idprograma
     WHERE NOT EXISTS (SELECT 1 FROM ag_vinculaciones v2 WHERE v2.id_corrida = c2.id)
       AND (@idPeriodo  IS NULL OR c2.idperiodo  = @idPeriodo)
       AND (@idPrograma IS NULL OR c2.idprograma = @idPrograma)
       AND (@subunidad  IS NULL OR UPPER(TRIM(c2.subunidad)) = UPPER(TRIM(@subunidad)))
     ORDER BY IdCorrida DESC;
     ```
  4. **`GetResumenStagingAsync` en dos consultas separadas**:
     - **Consulta A (SOLO corrida, descriptores reales sin filtros UI)**:
       ```sql
       SELECT COUNT(*) AS SesionesCorrida,
              COUNT(DISTINCT IdVinculacion) AS VincCorrida,
              (SELECT GROUP_CONCAT(p.periodo, ', ')
                 FROM (SELECT DISTINCT IdPeriodo FROM {Vista} WHERE IdCorrida=@idCorrida ORDER BY IdPeriodo) d
                 JOIN periodos p ON p.idperiodo = d.IdPeriodo) AS PeriodosCorrida,
              (SELECT GROUP_CONCAT(d.Subunidad, ', ')
                 FROM (SELECT DISTINCT Subunidad FROM {Vista} WHERE IdCorrida=@idCorrida AND Subunidad IS NOT NULL AND Subunidad <> '' ORDER BY Subunidad) d) AS SubunidadesCorrida,
              (SELECT GROUP_CONCAT(COALESCE(pr.nombreprograma, pr.abreviatura, 'PROG ' || d.IdPrograma), ', ')
                 FROM (SELECT DISTINCT IdPrograma FROM {Vista} WHERE IdCorrida=@idCorrida ORDER BY IdPrograma) d
                 LEFT JOIN programas pr ON pr.id = d.IdPrograma) AS ProgramasCorrida
       FROM {Vista} WHERE IdCorrida = @idCorrida;
       ```
     - **Consulta B (corrida + filtros activos de UI)**:
       Usa `BuildWhere(filtros, incluye idCorrida)` sobre `{Vista}` para calcular `SesionesConFiltros`.
     - **Determinismo en C#**: para evitar variaciones del agregador de SQLite con múltiples periodos/subunidades, normalizar en memoria con `string.Join(", ", valores.Distinct().OrderBy(s => s, StringComparer.Ordinal))`. Test 9 valida caso con periodo único determinista ("2027-01").
- **Patrón obligatorio**: copiar el estilo de resiliencia existente (`using (var db = _connectionFactory())`, `if (db.State != ConnectionState.Open) db.Open();`, `catch (OperationCanceledException)`, `LoggerHelper.Error/Warning`).
- **verification**: `scripts/verify.ps1 -Mode Build`
- **status**: done

### T3 · Fase 1490 — Delegación y población de Etiqueta en VinculacionListService *(implementer, L1)*
- **fase**: 1490
- **depends_on**: [T2]
- **lease_annotation**: LEASE-SERIALIZED T3 <- T2
- **files_lock**: `Academic/Application/Interfaces/IVinculacionListService.cs`, `Academic/Application/Services/VinculacionListService.cs`
- **description**:
  1. Añadir a `IVinculacionListService` las firmas:
     ```csharp
     Task<Result<IReadOnlyCollection<CorridaStagingItem>>> GetCorridasStagingAsync(
         int? idPeriodo, int? idPrograma, string subunidad, CancellationToken ct);

     Task<Result<AgStagingResumen>> GetResumenStagingAsync(
         long idCorrida, VinculacionFilterDto filtros, CancellationToken ct);
     ```
  2. En `VinculacionListService.GetCorridasStagingAsync` (capa Application): delegar en `_queryRepository`, y para cada ítem poblar `it.Etiqueta = AgCorridaSelector.FormatearEtiqueta(it);` antes de retornar (preserva la frontera de capa limpia `Application -> Domain`, sin que `Infrastructure` conozca `Application`):
     ```csharp
     var res = await _queryRepository.GetCorridasStagingAsync(idPeriodo, idPrograma, subunidad, ct);
     if (res.IsFailure) return res;
     foreach (var it in res.Value)
     {
         it.Etiqueta = AgCorridaSelector.FormatearEtiqueta(it); // Etiqueta poblada en Application
     }
     return res;
     ```
  3. `GetResumenStagingAsync`: delegación directa en `_queryRepository.GetResumenStagingAsync`.
- **validation**: las pruebas existentes de `VinculacionQueryRepositoryTests` (Unit + Integration) siguen verdes **sin modificarlas**.
- **verification**: `scripts/verify.ps1 -Mode Fast -Filter "FullyQualifiedName~VinculacionQueryRepository"`
- **status**: done

---

# Wave 2 · Ronda 283 — `FormVinculaciones` (lo que ve el usuario)

### T4 · Fase 1491 — Filtros: no aplicar `config` en MODO AG *(implementer, L1)*
- **fase**: 1491
- **files_lock**: `Academic/UI/Forms/FormVinculaciones.cs`
- **description**: en `CargarCatalogosAsync()` (líneas **326‑382**):
  - Envolver la lectura de `config.id_periodo` / `config.subunidad` y las dos preselecciones en `if (AgFiltrosPolicy.AplicarPeriodoConfig(_modo)) { … }` e `if (AgFiltrosPolicy.AplicarSubunidadConfig(_modo)) { … }` respectivamente (en AG ambas devuelven `false` ⇒ combos en índice 0 = "Todos").
  - `FiltrarSedesYTerceros()` (línea 377) debe ejecutarse **siempre**, en ambos modos.
  - No tocar el bloque de PRODUCCIÓN ni el manejo de `_isUpdatingFilters`.
- **validation**: en AG los combos ocultos quedan en índice 0; en Producción el comportamiento es idéntico al actual.
- **verification**: `scripts/verify.ps1 -Mode Fast -Filter "FullyQualifiedName~AgFiltrosPolicyTests"`
- **status**: done

### T5 · Fase 1492 — Combo de corridas correcto sin async void *(implementer, L1)*
- **fase**: 1492
- **files_lock**: `Academic/UI/Forms/FormVinculaciones.cs`
- **description**: reescribir el bloque **118‑157**:
  1. Nuevo campo `private long? _idCorridaInicial;` y parámetro opcional en el ctor:
     ```csharp
     public FormVinculaciones(IVinculacionListService listService, ISedeLookupRepository sedeLookup,
         ITerceroLookupRepository terceroLookup, Func<IDbConnection> dbFactory, dynamic excelApp,
         AgModoDatos modo = null, long? idCorridaInicial = null)
     ```
  2. Eliminar el SQL inline. Nuevo método `private async Task CargarCorridasAsync(bool preservarSeleccion = false)` que:
     - llama `_listService.GetCorridasStagingAsync(null, null, null, _cts.Token)`;
     - antepone `AgCorridaSelector.CrearTodasLasCorridas()` (los ítems ya traen `Etiqueta` poblada desde Application);
     - `cmbCorrida.DisplayMember = nameof(CorridaStagingItem.Etiqueta)`; `ValueMember = nameof(CorridaStagingItem.IdCorrida)`;
     - selección: `preservarSeleccion ? (id actual si sigue existiendo) : AgCorridaSelector.ResolverSeleccionInicial(items, _idCorridaInicial)`;
     - **todo bajo `_isUpdatingFilters = true/false`** para no disparar `EjecutarConsultaAsync` a media carga.
  3. **Orden de carga corregido** en el handler `Load`: crear controles → `await CargarCatalogosAsync()` → `await CargarCorridasAsync()` → **una sola** `await EjecutarConsultaAsync()`.
  4. Refresco seguro sin `async void`:
     ```csharp
     cmbCorrida.DropDownStyle = ComboBoxStyle.DropDownList;
     cmbCorrida.DropDown += (s, e) => _ = CargarCorridasSafeAsync();

     private async Task CargarCorridasSafeAsync()
     {
         try 
         { 
             await CargarCorridasAsync(preservarSeleccion: true); 
         }
         catch (OperationCanceledException) { }
         catch (Exception ex) 
         { 
             LoggerHelper.Warning($"[FormVinculaciones] Error al refrescar corridas: {ex.Message}"); 
         }
     }
     ```
  5. `EjecutarConsultaAsync`: `if (_modo.EsModoAg && cmbCorrida.SelectedItem is CorridaStagingItem c && c.IdCorrida > 0) idCorridaFiltro = c.IdCorrida;`
- **validation**: abriendo desde Módulo R con #44 cargada, el combo queda en #44 y la grilla muestra 244 filas sin pulsar nada.
- **verification**: `-Mode Fast -Filter "FullyQualifiedName~AgCorridaSelector"`
- **status**: done

### T6 · Fase 1493 — Mensaje accionable con guardia de centinela + botón "Ajustar filtros" *(implementer, L1)*
- **fase**: 1493
- **files_lock**: `Academic/UI/Forms/FormVinculaciones.cs`
- **description**:
  1. Sustituir el bloque **473‑495** (SQL crudo + aviso) por llamada a `_listService.GetResumenStagingAsync(idCorridaVal, filtros, _cts.Token)`.
  2. **Manejo de centinela**:
     ```csharp
     if (idCorridaVal <= 0)
     {
         lblStatus.Text = $"{_syncStatusText}Mostrando todas las corridas. Seleccione una para ver su diagnóstico.";
     }
     else
     {
         // 3 casos:
         // 1. SesionesConFiltros > 0 => "Se encontraron {vinc} vinculaciones / {n} sesiones (corrida #{id} · {PeriodosCorrida} · {SubunidadesCorrida})."
         // 2. SesionesCorrida > 0 && SesionesConFiltros == 0 => "La corrida #{id} tiene {SesionesCorrida} sesiones del periodo {PeriodosCorrida} ({SubunidadesCorrida}), pero los filtros activos dejan 0. Use «Ajustar filtros a la corrida»."
         // 3. SesionesCorrida == 0 => "La corrida #{id} no tiene sesiones en staging (estado {estado}). Ejecute el AG o elija otra corrida."
     }
     ```
  3. Nuevo `btnAjustarFiltrosCorrida` (visible solo en AG): fija periodo→territorial→programa→sede desde el ítem seleccionado y relanza la consulta:
     ```csharp
     _isUpdatingFilters = true;
     try {
         // periodo: buscar LookupItem con Id == item.IdPeriodo en cmbPeriodo
         // territorial: buscar LookupItem.Subunidad == item.Subunidad en cmbTerritoriales
         FiltrarSedesYTerceros();          // repuebla sedes/terceros según territorial
         // programa: LookupItem con Id == item.IdPrograma en cmbProgramas
     } finally { _isUpdatingFilters = false; }
     await EjecutarConsultaAsync();
     ```
  4. Prohibido cualquier `ExecuteScalarAsync` en el formulario tras este cambio.
- **validation**: con `config.id_periodo=2` y corrida #44 (2027-01) el aviso nombra el periodo real y un clic muestra las 244.
- **verification**: `-Mode Fast -Filter "FullyQualifiedName~VinculacionQueryRepository"`
- **status**: done

### T7 · Fase 1494 — Layout con MinimumSize, DPI y cableado del call site *(implementer, L1)*
- **fase**: 1494
- **files_lock**: `Academic/UI/Forms/FormVinculaciones.cs`, `Academic/UI/Forms/FormVinculaciones.Designer.cs`, `Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **description**:
  1. Mover `lblCorrida`/`cmbCorrida`/`btnAjustarFiltrosCorrida` a una **tercera fila** del `pnlFilters` (`Location.Y = 75/78`), `pnlFilters.Size` alto `80 → 105` en el Designer; fijar `pnlFilters.MinimumSize = new Size(0, 105);`.
  2. Eliminar los `pnlFilters.Width - 310/-365` (hoy tapan `chkSoloViaticas` en `(760,17)`, `btnAdicionar` en `(920,15)` y `btnVerVinculaciones` en `(920,45)`).
  3. Fijar anclajes explícitos para estabilidad en redimensionamiento:
     ```csharp
     lblCorrida.Anchor               = AnchorStyles.Top | AnchorStyles.Left;
     cmbCorrida.Anchor               = AnchorStyles.Top | AnchorStyles.Left;   // Width = 320 fijo
     btnAjustarFiltrosCorrida.Anchor = AnchorStyles.Top | AnchorStyles.Left;
     pnlFilters.Anchor               = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
     ```
     Verificar visualmente a 100 %, 125 % y 150 % de escala DPI (no solapar y no cortar "Adicionar/Vinculaciones").
  4. Call site Módulo R (`AgFormGeneradorHorarios.cs:2711`):
     ```csharp
     new FormVinculaciones(listService, sedeRepo, terceroRepo, _connectionFactory,
         ExcelDnaUtil.Application, modo,
         _idCorridaCargada > 0 ? _idCorridaCargada : (long?)null)
     ```
  5. **`RibbonController.cs` sin cambios** (el parámetro es opcional).
- **validation**: ninguna superposición visual; al abrir desde Módulo R la corrida preseleccionada es la cargada.
- **verification**: `scripts/verify.ps1 -Mode Build`
- **status**: done

---

# Wave 3 · Ronda 284 — Pruebas, documentación y cierre

### T8 · Fase 1495 — Suite RED→GREEN + docs + sello *(tdd-red-writer → implementer → build-verifier)*
- **fase**: 1495
- **files_lock**: `ConversorPDF.Tests/Academic/Application/Services/AgCorridaSelectorTests.cs` **(nuevo)**, `ConversorPDF.Tests/Academic/Application/Services/AgFiltrosPolicyTests.cs` **(nuevo)**, `ConversorPDF.Tests/Unit/Academic/VinculacionListServiceTests.cs`, `ConversorPDF.Tests/Integration/Academic/VinculacionQueryRepositoryTests.cs`, `docs/articles/changes/2026-11-09-corridas-ag-formvinculaciones.md` **(nuevo)**, `.agents/plans/REGISTRY.md`
- **Tests obligatorios** (nombres exactos sugeridos):

| # | Test | Archivo | Aserción |
|---|---|---|---|
| 1 | `AgFiltrosPolicy_ModoAg_NoAplicaPeriodoNiSubunidadDeConfig` | `AgFiltrosPolicyTests` | `AplicarPeriodoConfig(AgModoDatos.Ag).Should().BeFalse()` y `…Determinista…BeTrue()` |
| 2 | `ResolverSeleccionInicial_PrefiereCorridaEnUso_SobreMayorId` | `AgCorridaSelectorTests` | items [#46 sin staging, #44 con staging], `idCorridaEnUso=44` ⇒ 44 |
| 3 | `ResolverSeleccionInicial_SinCorridaEnUso_EligePrimeraConStaging` | idem | ⇒ #44 (no #46) |
| 4 | `ResolverSeleccionInicial_SinStaging_DevuelveTodas` | idem | `IdCorrida == 0` |
| 5 | `FormatearEtiqueta_ConStaging_ContienePeriodoProgramaSubunidadYSesiones` | idem | `"#44 · 2027-01 · … · VALLE · 244 ses"` |
| 6 | `FormatearEtiqueta_SinStaging_MarcaSinSesiones` | idem | contiene `"(sin sesiones)"`; no contiene `"· ·"` |
| 7 | `GetCorridasStagingAsync_IncluyeCorridaHuerfanaConStaging` | `Integration/…/VinculacionQueryRepositoryTests` | fixture con `ag_vinculaciones.id_corrida=45` **sin** `ag_corrida` ⇒ aparece con `TotalSesiones = 3` |
| 8 | `GetCorridasStagingAsync_MarcaCorridasSinStaging` | idem | #46 y #37 ⇒ `TieneStaging == false` |
| 9 | `GetResumenStagingAsync_PeriodoDistinto_ReportaCeroConFiltrosYPositivoSinFiltros` | idem | corrida periodo 3, filtro periodo 2 ⇒ `SesionesCorrida>0`, `SesionesConFiltros==0`, `PeriodosCorrida=="2027-01"` |
| 10 | `GetListaSesionesAsync_ConIdCorridaYPeriodoCoincidente_DevuelveFilas` | idem | control positivo (garantiza que el fix no "arregla" ocultando) |
| 11 | `VinculacionListService_GetCorridasStagingAsync_DelegaEnRepositorio` | `VinculacionListServiceTests` | `Moq.Verify … Times.Once` |
| 12 | `VinculacionListService_GetResumenStagingAsync_DelegaEnRepositorio` | idem | idem |

- **Fixture del test de integración** (SQLite `:memory:`, patrón ya usado en el archivo): replicar la realidad — `config.id_periodo=2`; `periodos` 1/2/3 = 2026-01/2026-02/2027-01; corrida 44 (periodo 3, 3 vinc, 3 ses), 46 (BOLIVAR, estado 4, sin staging), 37 (sin staging), 1 (periodo 2); y `ag_vinculaciones.id_corrida=45` sin cabecera. **Regla: no debilitar ni renombrar las 4 pruebas existentes de ese archivo.**
- **Documentación**: `docs/articles/changes/2026-11-09-corridas-ag-formvinculaciones.md` con el diagnóstico (tabla §0), la causa raíz y la tabla de aceptación de abajo.
- **Cierre**: `REGISTRY.md` → `R-007` a `closed`, `verify`, `HEAD`, contadores §2 → `282 / 1495`.
- **verification (secuencia obligatoria, serializada)**:
  ```
  1) scripts/verify.ps1 -Mode Build
  2) scripts/verify.ps1 -Mode Fast -Filter "FullyQualifiedName~AgCorridaSelector"
  3) scripts/verify.ps1 -Mode Fast -Filter "FullyQualifiedName~AgFiltrosPolicy"
  4) scripts/verify.ps1 -Mode Fast -Filter "FullyQualifiedName~VinculacionQueryRepository"
  5) scripts/verify.ps1 -Mode Fast -Filter "FullyQualifiedName~VinculacionListService"
  6) scripts/verify.ps1 -Mode Full          ← exit 0, TOTAL ≥ 952
  7) scripts/audit-tasks.ps1 -Plan .agents\plans\2026-11-09-corridas-ag-formvinculaciones-plan.md
  ```
- **evidence**: `_evidence/<stamp>/summary.md` + `test.trx` enlazados en el plan.
- **status**: done

---

## 2. Criterios de aceptación (verificables sobre el vault real)

| # | Escenario | Esperado |
|---|---|---|
| 1 | Abrir *Vinculaciones (AG)* desde Módulo R con **#44** cargada | Combo en **#44**; grilla con **244** filas; status `244 sesiones (corrida #44 · 2027-01 · VALLE)`; **cero** avisos en el log |
| 2 | Abrir y pulsar **Filtrar** con los combos visibles en "Todos" | Sigue mostrando **244** (no 0). Se elimina la dependencia de `config` |
| 3 | Combo: #46 y #37 | Visibles como `(sin sesiones)`, **no** preseleccionadas |
| 4 | Combo: corrida **#45** (huérfana BOLIVAR) | **Aparece** con `3 ses` (hoy es invisible) |
| 5 | Corrida **#1** (2026-02, 145 ses) | Seleccionable y carga sus 145 filas |
| 6 | Cambiar `config.id_periodo` a cualquier valor y reabrir | El resultado de AG **no cambia** |
| 7 | Ribbon → *Vinculaciones* (**Producción**) | Idéntico a hoy: sin combo de corrida, con filtros de `config` |
| 8 | Botón «Ajustar filtros a la corrida» con #44 | 244 filas sin tocar Configuración |
| 9 | Gate `-Mode Full` | exit `0`, `TOTAL ≥ 952`, 0 regresiones |
| 10 | `audit-tasks.ps1` | sin hallazgos `Compilation`/`Runtime`/`Logic` |

## 3. Riesgos y mitigaciones

| Riesgo | Mitigación |
|---|---|
| Romper Producción al tocar `CargarCatalogosAsync` | Toda la rama nueva se gatea con `_modo.EsModoAg` / `AgFiltrosPolicy`; criterio (7) lo verifica |
| `UNION ALL` con columnas mal alineadas | Test 7/8 con fixture real; el `UNION` exige **mismo orden y tipo** en ambas ramas |
| `ORDER BY` en consulta compuesta | `ORDER BY IdCorrida DESC` al final; el nombre sale del primer `SELECT` |
| Refactor de `BuildWhere` rompe filtros existentes | Pruebas existentes de `VinculacionQueryRepository` **no se tocan**; deben seguir verdes |
| Conexión compartida dispuesta | Usar siempre `_dbFactory`; nunca `Dispose()` sobre ella |
| `Dapper` + `long` vs `int` en `LookupItem` | DTO `CorridaStagingItem` dedicado con `long IdCorrida` |
| Layout en DPI alto | Tercera fila fija + `MinimumSize = (0, 105)` + `Anchor`; revisión manual 100/125/150 % |
| Baseline envejecido | No se edita en Fase B; Fase C actualiza `test_baseline.txt` |
