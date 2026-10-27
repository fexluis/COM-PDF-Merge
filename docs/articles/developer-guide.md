# Guía del Desarrollador

Esta sección proporciona una visión técnica profunda para los desarrolladores que deseen mantener o extender las capacidades de **ConversorPDF**.

## Arquitectura del Proyecto

El proyecto sigue una arquitectura de capas simplificada para facilitar la integración con Excel-DNA:

1.  **Capa UI (Ribbon & FormProgramador)**: Contiene los controladores del Ribbon y los formularios WinForms. Los controladores deben ser delgados y delegar la lógica a `App` o `Helpers`.
2.  **Capa Core (Core)**: 
    -   **Services**: Lógica de negocio pesada (PDF, Usuarios, Viáticos).
    -   **Processors**: Implementaciones específicas de flujos de trabajo (Individual, Masivo, Gestión).
    -   **Repositories**: Acceso a datos (principalmente Excel y objetos de negocio).
    -   **Helpers**: Funcionalidades transversales y orquestación de UI (Logging, Diálogos, Optimización de Excel).

## Estándares de Codificación

### Logging Centralizado
Siempre utiliza `LoggerHelper` para registrar eventos. Los logs se guardan automáticamente en la carpeta de Documentos del usuario.
```csharp
LoggerHelper.Info("Mensaje informativo");
LoggerHelper.Error("Error detectado", exception);
```

### Gestión de Excel
Para operaciones intensivas en Excel, utiliza siempre `ExcelOptimizeHelper` para desactivar el refresco de pantalla y cálculos automáticos:
```csharp
using (ExcelOptimizeHelper.Optimize(excelApp))
{
    // Lógica intensiva aquí
}
```

## Extensión del Ribbon

Para añadir un nuevo botón:
1.  Modifica `Ribbon.xml` añadiendo el elemento `<button>`.
2.  Define el `onAction` (nombre del método).
3.  Implementa el método en `RibbonController.cs`.
4.  Si el botón requiere selección de archivos o lógica compleja de orquestación, crea un nuevo **Helper** en `Core.Helpers`.

## Sistema de Metadatos

ConversorPDF inyecta metadatos personalizados en los archivos PDF utilizando un diccionario propio (`/MetadataCustom`). Esto permite que otros sistemas identifiquen el documento sin necesidad de realizar procesos costosos de OCR.

### Actualización de Metadatos
Puedes actualizar campos individuales sin reconstruir todo el objeto `Metadata`:
```csharp
_metadataService.ActualizarMetadato(rutaPdf, "estado", "VIA_VERIFICADO");
```
El sistema mapea automáticamente alias comunes como `"estado"` a la propiedad interna correspondiente.

Para más detalles sobre la implementación técnica, consulta la sección de [Gestión de Metadatos](pdf-metadata.md).

## Refactorización de Helpers a Servicios

El proyecto se encuentra en una fase de transición para mover la lógica de negocio desde **Helpers Estáticos** (ej: `ContadorHelper`) hacia **Servicios Basados en Interfaces** (ej: `ContadorService`).

### Caso de Estudio: ContadorService
Al refactorizar servicios complejos que interactúan con Excel, se deben seguir estas pautas:

1.  **Inyección de Dependencias**: El servicio debe recibir sus dependencias (`IUserService`, `IPdfService`, `IPdfMetadataService`) en el constructor.
2.  **Modelos de Extracción**: No pases variables sueltas. Utiliza modelos como `ObligacionExtract` para agrupar datos extraídos de Excel.
3.  **Encapsulamiento de Variantes**: Si existen múltiples formatos de hoja, encapsula la lógica de detección en métodos como `ExtraerDatos(sheet)`, manteniendo los métodos públicos limpios y enfocados en el flujo de negocio.

### Ejemplo de Uso (ContadorService)
A continuación, se muestra cómo invocar el servicio desde un controlador o helper:

```csharp
// Obtener el servicio desde el contenedor de dependencias (App)
var contadorSvc = app.Contador; 

// 1. Convertir hojas de Excel a PDF con firma y metadatos
int procesados = contadorSvc.ConvertirObligacionContador(
    activeWorkbook, 
    userFirmaBase64, 
    (cur, tot) => ProgressBar.Update(cur, tot)
);

// 2. Unificar soportes PDFs (requiere IPdfService configurado)
int unificados = contadorSvc.UnificarObligacionContador(
    activeWorkbook, 
    "OB"
);
```

## Construcción y Release

El proyecto utiliza scripts de PowerShell para automatizar la construcción:
-   `build_project.ps1`: Compila la solución y empaqueta el archivo XLL.
-   `build_release.ps1`: Genera el instalador MSI utilizando WiX Toolset.

---

## Base de Datos Local Encriptada (Vault Pattern)

El módulo académico utiliza una base de datos SQLite encriptada con AES-256-CBC. Todo acceso a datos debe seguir el patrón de dos niveles (Vault Pattern) para garantizar estabilidad bajo Excel-DNA.

> [!IMPORTANT]
> El retry residual en el primer `Open()` de cada sesión de Excel es **comportamiento esperado**, no un bug.
> Ver `.agents/rules/sqlite-excel-dna-gotchas.md` para el diagnóstico completo del Shadow Copy.

### Nivel 1 — Bootstrap del Motor (Una vez por proceso)

Siempre llama a `EnsureInitialized()` antes de crear cualquier contexto. Es idempotente y retorna en <1ms si el motor ya fue inicializado.

```csharp
// En el constructor de cualquier formulario que use la BD académica:
SQLiteEngineBootstrap.EnsureInitialized();
```

Este método usa `Process.GetCurrentProcess().Modules` para detectar si `SQLite.Interop.dll`
ya está en memoria, sobreviviendo a las múltiples recargas de AppDomain de Excel-DNA.

### Nivel 2 — Unit of Work por Formulario

Cada formulario crea **su propio** `EncryptedSQLiteContext`. El contexto descifra la BD al
abrirse y la re-encripta al cerrarse. No reutilices contextos entre formularios.

```csharp
public partial class MiFormularioAcademico : Form
{
    private readonly EncryptedSQLiteContext _dbContext;

    public MiFormularioAcademico()
    {
        SQLiteEngineBootstrap.EnsureInitialized(); // Nivel 1

        // Nivel 2 — SqliteHelper resuelve la ruta y la contraseña
        _dbContext = SqliteHelper.GetContext();

        InitializeComponent();
        CargarDatosIniciales();
    }

    private void CargarDatosIniciales()
    {
        var db = _dbContext.Open(); // Descifra y abre
        var items = db.Query<MiEntidad>("SELECT * FROM mi_tabla").ToList();
    }

    protected override void OnFormClosing(FormClosingEventArgs e)
    {
        _dbContext?.Dispose(); // Re-encripta si hubo cambios
        base.OnFormClosing(e);
    }
}
```

### Señalizar Cambios para Re-encriptación

Si modificas datos, notifica al contexto para que sepa que debe re-encriptar al cerrar:

```csharp
db.Execute("INSERT INTO mi_tabla VALUES (@Nombre)", new { Nombre = "Ejemplo" });
_dbContext.NotifyChanges(); // ← Obligatorio para persistir cambios
```

### Workaround de Configuración (Testable)

Si necesitas probar el workaround de shadow-copy en un test unitario, usa la sobrecarga con directorio:

```csharp
// Producción (resuelve ruta shadow-copy automáticamente):
EncryptedSQLiteContext.EnsureSQLiteConfigurationFile();

// Tests (directorio simulado):
EncryptedSQLiteContext.EnsureSQLiteConfigurationFile("/tmp/shadow-copy-sim/");
```

## Gotchas de WinForms: DataGridViewCheckBoxColumn

Cuando se trabaja con `DataGridView` sin un `DataSource` vinculado (modo un-bound), la columna de tipo `DataGridViewCheckBoxColumn` tiene un comportamiento atípico conocido en WinForms:

1. **Bug del `DataPropertyName`**: Si la columna tiene asignado un `DataPropertyName` pero no hay `DataSource`, el motor interno de WinForms intentará resolver la propiedad contra un contexto nulo durante repaints o llamadas a `Rows.Add()`, provocando que el valor booleano se restablezca silenciosamente a `false`. **Solución**: Asegúrate de que `DataPropertyName` esté vacío en grids un-bound.
2. **Renderizado de Asignaciones Programáticas**: Si se asigna el valor de la celda de forma programática (ej. `row.Cells["ColViatica"].Value = true;`), el backing store se actualiza, pero la interfaz gráfica (el checkmark) podría no renderizarse de inmediato. **Solución**: Llama explícitamente a `dgv.Refresh()` después de poblar el grid para forzar la actualización visual del estado.

## Pruebas Unitarias y Desacoplamiento (v1.2.51+)

Para garantizar la integridad de los datos en módulos complejos (como el Académico), se debe seguir el patrón de desacoplamiento de lógica de UI.

### 1. Extracción de Lógica (Mappers)
No implementes lógica de negocio pesada dentro de controladores de Form (WinForms). Utiliza Mappers que reciban DTOs abstractos:

```csharp
// Mal
private void GuardarGrid() { foreach(DataGridViewRow r in dgv...) { ... } }

// Bien (Testeo Unitario Posible)
public static class MiModuloMapper {
    public static List<Model> Extract(IEnumerable<GridRowDto> rows) { ... }
}
```

> [!TIP]
> Consulta la regla formal en `.agents/rules/grid-mapper-pattern.md` para ver el patrón de implementación completo y las guardas de integridad obligatorias.

### 3. Patrón de Resultados y Errores (CSharpFunctionalExtensions)

Evita lanzar excepciones para reglas de negocio. Utiliza `Result` o `Result<T>` para canalizar errores de forma segura hacia la UI.

```csharp
// Ejemplo: Generador de Sesiones
public Result<IReadOnlyList<Sesion>> Generar(...)
{
    if (!Validar(dias))
        return Result.Failure<IReadOnlyList<Sesion>>("Días no consecutivos.");

    var sesiones = new List<Sesion>();
    // ... lógica ...
    return Result.Success<IReadOnlyList<Sesion>>(sesiones);
}

// Consumo en UI (WinForms)
var result = _service.Generar(...);
if (result.IsFailure)
{
    MessageBox.Show(result.Error, "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
    return;
}
_grid.DataSource = result.Value;
```

### 4. Inmutabilidad y Builder Pattern

Para modelos de estado complejo que se construyen paso a paso, utiliza objetos inmutables con un Builder interno o dedicado. Esto previene cambios accidentales de estado en capas superiores.

```csharp
// Modelo Inmutable
public sealed class AcademicDateAdjustment
{
    public bool NeedsAdjustment { get; }
    public IReadOnlyList<string> Messages { get; }

    internal AcademicDateAdjustment(bool needsAdjustment, List<string> messages)
    {
        NeedsAdjustment = needsAdjustment;
        Messages = messages.AsReadOnly();
    }
}

// Uso del Builder
var builder = new AcademicDateAdjustmentBuilder();
builder.AddMessage("Fecha ajustada");
AcademicDateAdjustment result = builder.Build();
```

### 5. Caché de Propiedades (Lazy Calculation)

Para propiedades pesadas que dependen de colecciones (ej: cálculos LINQ sobre miles de filas), utiliza un patrón de **Dirty Flag** para cachear el resultado y mejorar el rendimiento de la UI (DataGridView).

```csharp
private double? _cachedTotal;
private bool _dirty = true;

public double TotalHoras
{
    get
    {
        if (_dirty || !_cachedTotal.HasValue)
        {
            _cachedTotal = Sesiones.Sum(s => s.Horas);
            _dirty = false;
        }
        return _cachedTotal.Value;
    }
}

public void Agregar(Sesion s)
{
    Sesiones.Add(s);
    _dirty = true; // Invalida la caché
}
```

### 6. Análisis de Conflictos (ViaticoConflictAnalyzer)

Para detectar solapamientos de agenda entre sedes distintas, utilice el servicio `ViaticoConflictAnalyzer`. Este servicio está desacoplado de la UI y es ideal para validaciones masivas.

```csharp
// 1. Preparar lista plana de (Sede, Bloque)
var bloquesDocente = new List<(string Sede, ViaticoBloque Bloque)>
{
    ("Sede A", bloque1),
    ("Sede B", bloque2)
};

// 2. Detectar bloques cruzados
HashSet<ViaticoBloque> conflictos = ViaticoConflictAnalyzer.DetectarBloquesCruzados(bloquesDocente);

if (conflictos.Contains(bloque1))
{
    // 3. Obtener detalle para Tooltip o logs
    string detalle = ViaticoConflictAnalyzer.ObtenerDetalleCruce(bloque1, bloquesDocente, "Sede A");
    Console.WriteLine(detalle);
}
```

### 7. Análisis de Conflictos Académicos y Salvedades (SesionConflictAnalyzer)

Para detectar conflictos de horarios en la programación académica de los docentes, utilice `SesionConflictAnalyzer` en conjunto con `HorarioSesionParser`. Esta lógica maneja tres estados de conflicto y permite cargar salvedades previamente autorizadas por el usuario.

```csharp
// 1. Cargar las salvedades confirmadas previamente desde la Base de Datos
var salvedades = await _salvedadRepository.GetByPeriodoAsync(idPeriodo);

// 2. Ejecutar el análisis para todas las sesiones de un docente
var analisis = SesionConflictAnalyzer.Analizar(sesionesDocente, salvedades);

// 3. Evaluar el estado de cada sesión (Ninguno, Salvado, RevisarViabilidad, ConflictoReal)
foreach(var sesion in sesionesDocente)
{
    var estado = analisis.ContainsKey(sesion.IdSesion) ? analisis[sesion.IdSesion] : EstadoConflicto.Ninguno;
    
    if (estado == EstadoConflicto.ConflictoReal)
    {
        Console.WriteLine($"La sesión {sesion.IdSesion} tiene un conflicto real de horario.");
    }
    else if (estado == EstadoConflicto.RevisarViabilidad)
    {
        Console.WriteLine($"La sesión {sesion.IdSesion} ocurre el mismo día en otra sede. Requiere confirmación de traslado.");
    }
}
```

### 8. Lógica de Distribución en Exportadores (AsignaturasExcelExporter)

Cuando un reporte requiere cruzar una entidad padre (Asignatura) con una entidad global agrupada (Viáticos Multi-Sede del Docente), se emplea un algoritmo de distribución con punteros y duplicación dinámica de filas (Phantom Rows) para acomodar el exceso de datos secundarios.

```csharp
// 1. Extraer grupos globales (Viáticos de todas sus asignaturas)
var bloquesGlobales = ViaticoBloque.AgruparConsecutivos(todasSesiones);
var asignaturas = ObtenerAsignaturas(todasSesiones);

int asigIndex = 0;
int bloqueIndex = 0;

// 2. Iterar hasta agotar tanto Asignaturas como Viáticos
while (asigIndex < asignaturas.Count || bloqueIndex < bloquesGlobales.Count)
{
    var grupoAsignatura = asigIndex < asignaturas.Count 
        ? asignaturas[asigIndex] 
        : asignaturas[0]; // Tomar la primera asignatura como base para Fila Duplicada

    bool esFilaDuplicada = asigIndex >= asignaturas.Count;

    if (esFilaDuplicada)
    {
        // 3. Imprimir solo campos de identificación (PKs visuales)
        EscribirCelda(ws, fila, CodMateria, grupoAsignatura.CodMateria);
        EscribirCelda(ws, fila, DocenteCedula, grupoAsignatura.Cedula);
    }
    else
    {
        // 4. Imprimir toda la información salarial y horas
        EscribirFilaCompleta(ws, fila, grupoAsignatura);
    }

    // 5. Tomar Chunk de hasta 4 viáticos y escribirlos en la fila actual
    var bloquesEnFila = bloquesGlobales.Skip(bloqueIndex).Take(4).ToList();
    EscribirComisiones(ws, fila, bloquesEnFila);

    asigIndex++;
    bloqueIndex += 4;
    fila++;
}
```

### 9. Suite de Pruebas (.NET 4.8)
El proyecto utiliza **xUnit** para lógica pura y **MSTest (STAThread)** para integración con WinForms.

- **Ubicación**: `ConversorPDF.Tests/`
- **Capas**:
  - **Capa 0 (Unit)**: Pruebas de lógica pura en Mappers y Helpers.
  - **Capa 1 (Integration)**: Pruebas de Repositorio con SQLite in-memory.
  - **Capa 3 (UI)**: Pruebas de estado interno de Form usando Reflexión y STAThread.

```powershell
# Ejecutar tests desde la terminal
dotnet test ConversorPDF.Tests\ConversorPDF.Tests.csproj
```
