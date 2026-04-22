# Arquitectura y Flujo de Datos — ConversorPDF

ConversorPDF es un complemento de Excel (XLL) desarrollado en **.NET Framework 4.8** utilizando **Excel-DNA**. Su propósito es la gestión académica, generación de documentos y procesamiento de viáticos con un alto enfoque en la seguridad y el rendimiento.

## 1. Estructura del Proyecto

El proyecto está organizado siguiendo principios de arquitectura multicapa:

*   **`ConversorPDF.XLL`**: Proyecto principal (Add-In).
    *   **`Core/Models/`**: Definición de entidades de dominio (Sesión, Vinculación, Tercero, etc.).
    *   **`Core/Services/`**: Lógica de negocio (ej. `GeneradorSesionesRecurrentes`).
    *   **`Core/Repositories/`**: Capa de acceso a datos. Abstrae el origen (Dapper para SQLite y `dynamic` para Excel).
    *   **`Core/Data/`**: El motor de persistencia segura (Encrypted Context) y esquemas SQL.
    *   **`Core/Helpers/`**: Utilidades transversales (Logging, Caché de catálogos, OneDrive Helper).
    *   **`FormAcademic/`**: Capa de presentación (WinForms) para el módulo de Sesiones Académicas.

## 2. Componentes Principales

### A. Persistencia Segura (The Vault Pattern)
Dado que la aplicación maneja datos sensibles (IDs de terceros, contratos), no se usa una base de datos abierta.
- **Archivo Maestro**: `AcademicData.db.enc` (AES-256).
- **EncryptedSQLiteContext**: Al abrirse, desencripta el archivo maestro en la memoria volátil del sistema (`%TEMP%`) y lo carga como una base de datos SQLite estándar.
- **Shared Context (Optimización 1.2.49.2)**: El formulario mantiene una única conexión abierta durante su ciclo de vida para evitar el costo de desencriptación (haciendo que la UI sea instantánea).

### B. Integración con Excel
- **Excel-DNA**: Permite exponer métodos de C# como funciones de Excel y callbacks de Ribbon.
- **ExcelRepository**: Interactúa con las celdas de Excel para leer configuraciones (como el Departamento Base) o escribir resultados de procesamiento.

### C. Generación de PDF
- **PdfPig / IText7**: Utilizados para leer metadatos de PDFs existentes (como el valor de los viáticos) y realizar mezclas o inserción de firmas.

## 3. Flujo de Datos (Data Flow)

### Escenario: Gestión de Sesiones Académicas

1.  **Inicialización**:
    - El usuario hace clic en el Ribbon.
    - Se instancia `FormSesionesAcademicas`.
    - Se dispara el **Shared Context**: Desencripta `AcademicData.db.enc` -> Crea DB temporal en RAM/Temp.
    - Se carga la **CatalogCacheService**: Pobla combos (Sedes, Programas) con un único viaje a la DB (Dapper `QueryMultiple`).

2.  **Generación de Datos**:
    - El usuario ingresa parámetros (Horas, Fechas).
    - `GeneradorSesionesRecurrentes` calcula el calendario en memoria.
    - El resultado se pre-visualiza en un `DataGridView`.

3.  **Persistencia y Sincronización**:
    - El usuario pulsa "Guardar".
    - Se ejecuta la **Validación Cruzada**: Verifica que los IDs de Tercero coincidan entre UI y modelo.
    - **AcademicRepository** guarda los cambios usando **Dapper Batch Inserts** (Transaccional).
    - Al finalizar, el sistema lee los bytes de la DB temporal, los cifra y sobreescribe el `AcademicData.db.enc`.
    - El archivo temporal es eliminado de forma segura (`SecureWipe`).

## 4. Tecnologías Clave

*   **Dapper**: Micro-ORM para mapeo de objetos SQL ultra-rápido.
*   **System.Data.SQLite**: Motor de base de datos local.
*   **AES-256 (PBKDF2)**: Derivación de claves para el Vault.
*   **WiX Toolset v4**: Generación del instalador MSI.

## 5. Mantenibilidad e Integridad (.agents)

El proyecto incluye una carpeta `.agents/workflows` que automatiza tareas críticas:
- `build-release`: Garantiza que el código compile, pase los 23 tests unitarios y genere el MSI.
- `contexto-previo`: Crea trazas de auditoría de cada cambio en los archivos `.cs` antes de cada despliegue.
