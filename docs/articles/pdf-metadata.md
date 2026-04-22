# Gestión de Metadatos PDF

El servicio de metadatos permite adjuntar información estructurada a los archivos PDF generados por ConversorPDF. Esta información es crucial para la integración con sistemas de gestión documental y para la trazabilidad de los documentos (CDP, Autor, Fechas, etc.).

## IPdfMetadataService

La interfaz `IPdfMetadataService` define tres operaciones principales:

- `GuardarMetadatos`: Persiste un objeto `Metadata` en el archivo PDF.
- `LeerMetadatos`: Recupera la información estructurada de un PDF existente.
- `ObtenerResumenMetadatos`: Genera un reporte legible en texto plano.
- `ActualizarMetadato`: Permite modificar un campo específico (ej: "estado") sin cargar todo el modelo.

### Ejemplo de Uso

Para guardar metadatos en un procesador, primero inyecta la interfaz en el constructor:

```csharp
private readonly IPdfMetadataService _metadataService;

public MyProcessor(IPdfMetadataService metadataService)
{
    _metadataService = metadataService;
}
```

Luego, construye el objeto y guárdalo:

```csharp
var info = new Metadata
{
    Titulo = "Solicitud de Viático",
    CDP = "12345",
    MetaDocumento = TipoMetadata.VIA_SOLICITUD
};

_metadataService.GuardarMetadatos(rutaArchivo, info);
```

## Modelo de Datos (Metadata)

El modelo `Metadata` contiene tanto campos estándar del PDF como diccionarios personalizados:

- **Estandar**: Título, Autor, Asunto, Palabras Clave.
- **Personalizados**: CDP, RP, OB, OP, Tipo, Comisión, Fecha y MetaDocumento.

> [!IMPORTANT]
> Los campos personalizados se guardan internamente en el diccionario `/MetadataCustom` dentro del catálogo del PDF para evitar colisiones con metadatos de otras herramientas.

## Rendimiento

El servicio utiliza `PdfSharp` para un acceso directo a los objetos internos del PDF. La operación de inyección de metadatos añade aproximadamente **10-20ms** de procesamiento por archivo, lo cual es despreciable en flujos individuales y masivos.

## Herramienta de Usuario (Ribbon)

Desde la versión **1.2.46.01**, se incluye una herramienta visual para que los usuarios finales puedan auditar los metadatos de sus PDFs.

### Cómo usarla:
1. Navega a la pestaña **ESAP**.
2. Despliega el menú del botón principal **PDF**.
3. Selecciona la opción **Metadatos**.
4. Selecciona uno o varios archivos PDF.
5. El sistema procesará los archivos y generará un reporte detallado en el log de la aplicación.

> [!TIP]
> Esta herramienta es ideal para verificar que los campos `metaFecha`, `metaCDP` y `metaTipo` se hayan grabado correctamente antes de radicar documentos en sistemas externos.
