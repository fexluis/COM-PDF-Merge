# Botón "Buscar Informes" (v1.6.3.72)

## 1. Qué hace

El botón **Buscar Informes** del menú PDF del Ribbon abre `FormBuscarInformes`, que ejecuta el
mismo flujo del módulo VBA histórico (`Módulo2`), pero **sin binarios externos**:

- Lectura y búsqueda de texto en PDF: **PdfPig** (`IPdfToTexto` / `PdfToTextoHelper`).
- Unificación de PDFs: **PDFsharp** (`IPdfService` / `PdfService`).
- Ya no se usa `pdftotext.exe`, `PDFConversor.exe`, `WScript.Shell`, `cmd.exe` ni archivos `.txt` temporales.

## 2. Flujo

### Fase 1 — Renombrado y unificación (`BusquedaInformesProcessor.RenombrarYUnificar`)

1. Recorre `<carpeta del libro>\Informes` de forma recursiva, **excluyendo** `VERIFICADOS`, `_Duplicados`
   y cualquier carpeta que empiece por `_`.
2. Para cada PDF sin nombre estándar:
   - Extrae el número de solicitud del nombre y busca la cédula/nombre del comisionado en **CONSULTAS**.
   - Lee la página 1 (y la 2 como respaldo) con PdfPig y detecta el tipo:
     `CERTIFICADO ... CUMPLIMIENTO` → `CUM`, `INFORME ... CUMPLIMIENTO` → `INF`, si no → `ERR`.
   - Verifica la cédula (ver sección 4) y aplica la tabla de decisión del VBA.
   - Renombra a `{solicitud}-{cedula}-{tipo}-{nombre}` (12 caracteres, sin tildes ni separadores).
3. En las carpetas especiales (`Academica`, `Capacitacion`, `Direccion`, `Fortalecimiento`) forma
   parejas `INF + CUM` por (solicitud, nombre) y genera `{solicitud}-{cedula}-UNI-{nombre}.pdf`
   **en la raíz de `Informes`**, con el orden CUM → INF.
   - El PDF se escribe primero en `*.tmp.pdf`, se valida con PdfPig (`páginas salida == CUM + INF`)
     y solo entonces se mueve a su destino y se borran los originales.

### Fase 2 — Auditoría de la columna INFORME (`AuditarConsultas`)

1. Los `UNI` de la raíz se contrastan con la cédula de CONSULTAS:
   - coincide → `INFORME = VERIFICADO` y el archivo se mueve a `Informes\VERIFICADOS`;
   - no coincide → `INFORME = ERROR` y el archivo se renombra a `ERR`.
2. Por cada fila de CONSULTAS dentro del rango:
   - `ERR` presente en carpeta especial → `ERROR`
   - solo `INF` → `FALTA CUM`
   - solo `CUM` → `FALTA INF`
   - ambos → `PAR OK` (no se reescribe)
   - sin archivos → `SIN INFORME`
3. Se respetan los estados protegidos: `VERIFICADO`, `REINTEGRO`, `REVOCADA`, `RECHAZADA`, `ANULADA`.
4. Si hubo escrituras, se guarda el libro.

## 3. Reglas de nombre

- `{solicitud}-{cedula}-{tipo}-{nombre}` con separador `-`.
- Tipos válidos: `INF`, `CUM`, `UNI`, `ERR`.
- `nombre`: máx. 12 caracteres, mayúsculas, sin tildes, sin espacios ni símbolos.
- Colisiones: contenido idéntico → `_Duplicados`; contenido distinto → sufijo `_1`, `_2`…
  Nunca se sobrescribe ni se borra un archivo sin validar antes el resultado.

## 4. Verificación de la cédula (información contenida en el PDF)

Se compara contra las cédulas de CONSULTAS para la solicitud, con cuatro niveles:

| Nivel | Estrategia | Confianza |
|---|---|---|
| E1 | Cédula seguida de correo (formato GF-FO-023) | Alta |
| E2 | Etiqueta (`No. Identificación Persona`, `C.C.`, `Cédula`, `Documento`…) y número en la misma línea | Alta |
| E3 | Coincidencia exacta y única con una cédula de CONSULTAS | Media |
| E4 | Heurística libre (6-10 dígitos, descarta años/celulares) | Baja (no valida por sí sola) |

Reglas: ambigüedad (dos cédulas candidatas distintas) ⇒ **no verificado**; PDF sin capa de texto ⇒
**no verificable**. Cualquier fallo de verificación produce `ERR` (fail-closed, nunca `XXX`).

## 5. Permisos y UI

- Botón gobernado por el permiso `BuscarInformes` (perfiles `Admin` y `Auditor` por defecto).
- El handler del Ribbon valida `EnsureInitialized`, `HasPermission(RibbonButtonId.BuscarInformes)`
  y que la hoja activa sea `CONSULTAS`.
- Formulario: rango de solicitudes (prellenado con `ObtenerRangoSolicitudes`), carpeta editable,
  casillas "Unificar parejas", "Actualizar INFORME" y "Simulación" (dry-run), log por archivo y resumen final.

## 6. Archivos

**Nuevos (XLL):** `Core/Models/{InformeTipo,InformeSolicitud,InformePdf,MotivoInforme,ConfianzaCedula,ResultadoCedulaPdf,ResultadoBusquedaInformes}.cs`,
`Core/Helpers/{InformeNombreHelper,InformeTextoHelper}.cs`, `Core/Interfaces/IBusquedaInformesProcessor.cs`,
`Core/Processors/BusquedaInformesProcessor.cs`, `FormProgramador/FormBuscarInformes{,.Designer}.cs`.

**Modificados:** `Ribbon/RibbonController.cs` (handler real), `Core/Services/ViaticosService.cs` y
`Core/Interfaces/IViaticosService.cs` (`ObtenerCarpetaInformes`, `ProcesarInformes`),
`Core/Helpers/App.cs` (carpetas), `Core/Helpers/ConstantsHelper.cs` (carpetas especiales),
`Core/Interfaces/IPdfToTexto.cs` y `Core/Helpers/PdfToTextoHelper.cs` (`ContarPaginas`),
`Core/Models/EstadoComision.cs` (estados nuevos), `ConversorPDF.XLL.csproj`.

**Tests:** `ConversorPDF.Tests/Core/Helpers/Informe*Tests.cs`, `Core/Processors/BusquedaInformes*Tests.cs`,
`Core/Services/PdfServiceTests.cs`, `Integration/Informes/BusquedaInformesFlujoCompletoTests.cs`
y `TestData/{PdfFixtureFactory,ExcelRepositoryFalsa}.cs`.

## 7. Limitaciones conocidas

- Los PDF escaneados sin capa de texto no pueden clasificarse: quedan como `ERR` (no hay OCR).
- PDFsharp 1.50 no abre documentos cifrados o dañados: la unificación se marca como fallida y
  **se conservan los originales**.
- La página de búsqueda de texto es la 1 con respaldo a la 2; documentos con la cédula solo en
  páginas posteriores no se validan.
