# Gestión Documental de Viáticos

El `ViaticosGestionProcessor` es el componente encargado de la organización y consolidación de los expedientes de viáticos. Su objetivo es asegurar que los documentos PDF (solicitudes, resoluciones, registros presupuestales) estén correctamente nombrados y unificados para su archivo definitivo.

## Renombramiento de Solicitudes

Este proceso resuelve el problema de los archivos PDF descargados con nombres genéricos. Utiliza una combinación de datos de Excel y Reconocimiento de Texto (OCR) para renombrar los archivos siguiendo el patrón `[Comisión]-[Cédula]-[Nombre].pdf`.

### Funcionamiento y Optimización:
1.  **Lectura de Excel**: Obtiene la lista de comisiones vigentes de la hoja `CONSULTAS`.
2.  **Capa 1 - Metadatos (Optimización)**: El sistema primero intenta leer los metadatos internos del PDF. Si el `MetaDocumento` es `VIA_SOLICITUD`, extrae la cédula directamente desde las `PalabrasClave`. Esta operación es instantánea y evita el uso de OCR.
3.  **Capa 2 - Identificación por Nombre**: Intenta machear el archivo por el número de comisión que suele estar al inicio del nombre.
4.  **Capa 3 - Validación via OCR (Fallback)**: Solo si hay múltiples coincidencias por número y no se encontraron metadatos válidos, utiliza `IPdfToTexto` para leer el contenido del PDF y extraer la **Cédula** mediante expresiones regulares.
5.  **Estandarización**: Renombra el archivo final siguiendo el patrón `[Comisión]-[Cédula]-[Nombre].pdf`.

## Unificación de Expedientes

Este proceso genera un único archivo PDF consolidado por cada funcionario/comisión, agrupando todos los documentos de soporte en un orden lógico predefinido.

### Estructura del Expediente Unificado:
El servicio busca archivos en subcarpetas específicas dentro del directorio de trabajo y los combina en el siguiente orden:

1.  **Solicitud/Soporte**: El documento inicial de trámite.
2.  **Autorización (Auto)**: Si existe (Resolución o autorización firmada).
3.  **Registro Presupuestal (RP)**: Documento de compromiso presupuestal.
4.  **Certificado de Disponibilidad (CDP)**: Documento de reserva presupuestal.
5.  **Verificación**: Documento de cumplimiento o validación final.

### Lógica de Búsqueda:
*   Se basa en el **Número de Comisión** para encontrar la Solicitud y los Autos.
*   Se basa en el campo **RP** para encontrar el Registro Presupuestal.
*   Se basa en el campo **CDP** para encontrar el Certificado de Disponibilidad.

## Dependencias Técnicas

*   **IPdfToTexto**: Utiliza la librería `PdfPig` para extraer texto de los PDFs sin necesidad de licencias comerciales.
*   **CombinarArchivos**: Utiliza `PdfSharp` para la concatenación de páginas de diferentes documentos.

---

> [!NOTE]
> Para que la unificación sea exitosa, los archivos deben estar previamente organizados en las subcarpetas estandarizadas del sistema (`Solicitudes`, `RPs`, `CDPs`, etc.).
