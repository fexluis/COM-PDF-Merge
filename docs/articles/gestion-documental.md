# Gestión Documental de Viáticos

El `ViaticosGestionProcessor` es el componente encargado de la organización y consolidación de los expedientes de viáticos. Su objetivo es asegurar que los documentos PDF (solicitudes, resoluciones, registros presupuestales) estén correctamente nombrados y unificados para su archivo definitivo.

## Convención de Números de Comisión `#YY` y Filtrado por Rango

Los números de comisión leídos desde la hoja `CONSULTAS` y desde los formatos individuales (`P26`) siguen la convención SIIF Nación `[Consecutivo][Año2Digitos]`:

* **`126`**: Consecutivo 1, vigencia 2026.
* **`14526`**: Consecutivo 145, vigencia 2026.
* **`12625`**: Consecutivo 126, vigencia 2025.

El filtrado por rango compara la tupla `(año, consecutivo)` de forma inclusiva:
* **Filtro activo**: Si se indica viático inicial y final (ej. `126` a `14526` o consecutivos `1` a `145`), solo se procesan las comisiones dentro del rango y del año correspondiente.
* **Sin filtro**: Valores `(0, 0)` procesan la totalidad de registros.

## Renombramiento de Solicitudes

Este proceso resuelve el problema de los archivos PDF descargados con nombres genéricos. Escanea la **carpeta raíz del libro de Excel y la subcarpeta `Solicitudes`** (sin recursión más allá de estas dos), renombrando los archivos bajo el patrón `[Comisión]-[Cédula]-[Nombre18].pdf`.

### Política de Desambiguación (N0 → N1 → N2 → N3)

Para evitar falsos positivos y ejecuciones innecesarias de OCR, el procesamiento opera en 4 niveles estrictos:

| Nivel | Regla implementada | Cuándo actúa | Resultado |
|---|---|---|---|
| **N0** | `sols.Where(Num == prefijo).GroupBy(NormalizarDocumento(CC))`; si queda 1 candidato único | Siempre, en primer lugar | Renombra sin OCR (colapsa duplicados idénticos en `CONSULTAS`). |
| **N1** | Metadatos internos (`MetaDocumento == VIA_SOLICITUD` y cédula en `PalabrasClave`) | PDF generados o marcados por la app | Renombra sin OCR de manera instantánea. |
| **N2** | `FindCC` validado exclusivamente contra el conjunto de cédulas candidatas | Comisión con más de una persona/cédula | Renombra si coincide una única cédula candidata (ignora rubros como `0505023`). |
| **N3** | Sin coincidencia unívoca concluyente | N0–N2 no concluyen | No se toca el archivo; se reporta motivo `AMBIGUA`. |

### Catálogo de Motivos de Renombrado

Cada archivo examinado se clasifica bajo uno de los siguientes motivos registrados en `ResultadoRenombrado.Detalle` y en el log:

* **`RENOMBRADO`**: Archivo renombrado exitosamente al nuevo patrón.
* **`YA_CORRECTO`**: El archivo ya cumple con el nombre estandarizado `[Num]-[CC]-[Nom18].pdf`.
* **`COLISION`**: El nombre de destino ya existe con contenido **distinto** (no se sobrescribe).
* **`DUPLICADO`**: El nombre de destino ya existe con contenido **idéntico**; el archivo se mueve a `Solicitudes\_Duplicados`.
* **`AMBIGUA`**: Múltiples comisionados posibles sin desambiguación concluyente.
* **`SIN_COMISION_EN_CONSULTAS`**: El número extraído del prefijo no existe en la hoja `CONSULTAS` dentro del rango.
* **`PREFIJO_INVALIDO`**: El nombre del archivo no inicia con una secuencia numérica de comisión.
* **`ERROR_IO`**: Excepción de entrada/salida al intentar mover o renombrar el archivo.

### Requisitos en la Hoja `CONSULTAS`

* Debe contener al menos las columnas de número de comisión, cédula de ciudadanía y nombre completo.
* Filas repetidas con la misma comisión y cédula son consolidadas automáticamente por el nivel N0.
* En comisiones compartidas por distintos funcionarios, el PDF debe incluir la cédula en el texto de la primera página para su resolución vía N2.

### Trazabilidad y Logs

El procesador emite trazas individuales a nivel `Debug` por cada archivo y una línea de resumen estructurado en el log:
`[RENOMBRAR] carpetas=... | comisionesEnRango=... | pdfs=... | renombrados=... | yaCorrectos=... | colisiones=... | ambiguos=... | sinComision=... | prefijoInvalido=... | erroresIo=...`

## Duplicados: subcarpeta `Solicitudes\_Duplicados`

Cuando el nombre destino `[Comisión]-[Cédula]-[Nombre18].pdf` ya existe:

1. Se comparan tamaños y contenido byte a byte.
2. **Idéntico** → el archivo de origen se mueve a `Solicitudes\_Duplicados` conservando su nombre (si colisiona allí, se añade sufijo `_<6hex>`), y se contabiliza como `DUPLICADO`.
3. **Distinto** → se reporta `COLISION` y no se mueve nada.

`_Duplicados` no se indexa: la verificación de presencia física sólo considera la carpeta raíz del libro y `Solicitudes` (sin recursión).

## Verificación de Presencia Física (columna INFORME)

Tras el renombrado, el proceso recorre la columna B (SOLICITUD) dentro del rango del formulario y evalúa la columna H (INFORME):

| Orden | Condición | Acción |
|---|---|---|
| 1 | `H ∈ {"VERIFICADO","REINTEGRO"}` | No consulta disco, no escribe (`OmitidasProtegidas`) |
| 2 | `C ∈ {REVOCADA, RECHAZADA, ANULADA}` | No consulta disco; copia el estado en H |
| 3 | Existe PDF con `Split('-')[0] = B` y `Split('-')[1] = J` | Si `H` contiene una marca de ausencia se **limpia** (`CeldasLimpiadas`); si no, no se toca |
| 4 | No existe el PDF | Escribe `SIN SOLICITUD` |
| — | `H` con otra marca no vacía (`ERROR`, `SIN INFORME`, `DUPLICADA`, …) | Se preserva (`OmitidasOtrasMarcas`) |

Comparaciones normalizadas (`Trim` + mayúsculas); cédulas sin puntos ni espacios. Al finalizar, si hubo escrituras o limpiezas, el libro se **guarda** (`Workbook.Save`) y el resultado reporta `LibroGuardado`.

Resumen en log: `[VERIFICAR_SOLICITUDES] filas=… | protegidas=… | otrasMarcas=… | copiadasEstado=… | conArchivo=… | sinArchivo=… | yaMarcadas=… | celdasEscritas=… | celdasLimpiadas=… | libroGuardado=…`

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
*   Se basa en el **Número de Comisión** para encontrar la Solicitud y los Autos, aplicando el filtro por rango sobre la celda `P26`.
*   Se basa en el campo **RP** para encontrar el Registro Presupuestal.
*   Se basa en el campo **CDP** para encontrar el Certificado de Disponibilidad.

## Dependencias Técnicas

*   **IPdfToTexto**: Utiliza la librería `PdfPig` para extraer texto de los PDFs sin necesidad de licencias comerciales.
*   **CombinarArchivos**: Utiliza `PdfSharp` para la concatenación de páginas de diferentes documentos.

---

> [!NOTE]
> Para que la unificación sea exitosa, los archivos deben estar previamente organizados en las subcarpetas estandarizadas del sistema (`Solicitudes`, `RPs`, `CDPs`, etc.).
