# Servicio de Obligaciones (Contador)

El `ContadorService` es el componente encargado de procesar los libros de Excel relacionados con el flujo de **Obligaciones Presupuestales**. Su función principal es automatizar la firma digital de las hojas, su exportación a PDF y la inyección de metadatos de trazabilidad.

## Flujo de Procesamiento

El servicio opera sobre el libro de Excel activo y realiza las siguientes acciones por cada hoja:

1.  **Validación de Formato**: Verifica que la celda `B14` contenga el texto `"REGISTRO PRESUPUESTAL DE OBLIGACION."`.
2.  **Extracción de Datos**: Utiliza el método `ExtraerDatos` para poblar un modelo `ObligacionExtract`. Identifica el formato de la hoja (existen dos variantes principales) para extraer:
    *   Número de Obligación (`ObNumero`).
    *   Fecha de la Obligación (`ObFechaStr`).
    *   Tercero e Identificación (`ObTercero`, `ObIdentificacion`).
    *   Registro Presupuestal (`ObRP`).
    *   Objeto del gasto (`ObObjeto`).
3.  **Firma Digital**: Inserta la imagen de la firma (proporcionada en Base64) en una posición calculada dinámicamente basándose en la ubicación del texto `"FIRMA(S) RESPONSABLE(S)"`.
4.  **Generación de PDF**: Exporta la hoja individualmente aplicando configuraciones de página óptimas (ajuste a una página, márgenes mínimos).
5.  **Inyección de Metadatos**: Utiliza el `IPdfMetadataService` para persistir la información de `ObligacionExtract` dentro del PDF.

## Metadatos PDF

A continuación se detalla cómo se mapean los campos de Excel a los metadatos del PDF:

| Campo Metadato | Origen en Excel / Sistema | Notas |
| :--- | :--- | :--- |
| **Título** | Nombre del archivo generado | Ej: `OB-123-VIA-JUAN_PEREZ` |
| **Autor** | Usuario de Windows | En mayúsculas y sin dominio. |
| **Asunto** | Celda de "Objeto" | `C38` o búsqueda dinámica de la fila. |
| **Palabras Clave** | Identificación, Tercero, Dependencia | Facilita la indexación en sistemas de gestión. |
| **Fecha** | Celda de Fecha | Se intenta parsear a `DateTime`. |
| **CDP / RP / OB** | Celdas correspondientes | Valores exactos del presupuesto. |
| **Tipo** | Categoría calculada | `VIA`, `HON`, `SSO`, etc. |

## Clasificación de Categorías

El servicio utiliza el método `GetCategoria` para asignar un código corto basado en el texto del Objeto:

*   **VIA**: Viáticos.
*   **HON**: Honorarios.
*   **SSO**: Seguridad Social.
*   **NOM**: Nómina.
*   **ARL**: Riesgos Laborales.
*   *Otros...*

## Unificación de Soportes

El método `UnificarObligacionContador` permite buscar archivos PDF en una carpeta de soportes (`\S\`) que coincidan con el número de obligación y combinarlos con el PDF principal generado, guardando el resultado en la carpeta `\Unificados\`.

---

> [!TIP]
> Para depurar la extracción de datos, revisa la lógica de selección de rangos en `ConvertirObligacionContador`, ya que las coordenadas de las celdas cambian si la hoja tiene el campo "Estado:" en `F16`.
