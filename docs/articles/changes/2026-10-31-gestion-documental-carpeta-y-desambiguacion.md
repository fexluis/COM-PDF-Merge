# Bitácora de Cambios: Renombrado de Solicitudes en Carpeta Solicitudes y Desambiguación Robusta

- **Fecha**: 2026-10-31
- **Plan de Referencia**: `.agents/plans/2026-10-31-gestion-documental-carpeta-solicitudes-fix-plan.md`
- **Rondas / Fases**: Rondas 266–268 / Fases 1437–1442 (Registro `R-003`)

---

## 1. Diagnóstico y Causa Raíz

1. **D1 - Carpeta errónea**: `RenombrarSolicitudes` escaneaba únicamente la raíz del libro (`GetOneDriveFolderPath()`), ignorando la subcarpeta `Solicitudes` donde residen los 419 PDFs reales de comisiones.
2. **D2 - Mensaje engañoso**: Cuando no se renombraba ningún PDF (`count == 0`), el formulario mostraba "No se encontraron comisiones para el rango indicado" aun cuando la hoja `CONSULTAS` contenía cientos de registros en rango.
3. **D3 - Deduplicación faltante**: Filas duplicadas idénticas en `CONSULTAS` (misma comisión y cédula) producían `matches.Count == 2`, forzando OCR innecesario.
4. **D4 - Falso positivo en OCR (`0505023`)**: La expresión regular general capturaba números de rubro presupuestal (`0505023`) en vez de la cédula del funcionario comisionado.
5. **D6 - Estandarización indiscriminada**: `Estandarizar` renombraba todos los archivos del directorio en vez de limitarse exclusivamente a `*.pdf`.

---

## 2. Cambios Implementados

### Fase 1437 (T1) — Procesador base y modelos de resultado
- Creación de `ResultadoRenombrado` y `MotivoRenombrado` en `ConversorPDF.XLL.Core.Models`.
- Escaneo dual de carpeta raíz y subcarpeta estándar `Solicitudes` mediante helper `CarpetasDeSolicitudes`.
- Restricción de `Estandarizar` a `*.pdf`.
- Registro detallado en log estructurado `[RENOMBRAR]`.

### Fase 1438 (T2) — Interfaces, servicio y mensajes UI
- Adición de `RenombrarSolicitudesConDetalle` en `IViaticosGestionProcessor` y `IViaticosService`.
- Reemplazo del manejador `BtnRenombrar_Click` en `FormGestionDocumental` con desglose informativo de comisiones en rango, PDFs examinados, renombrados, colisiones, ambiguos y errores.

### Fase 1439 (T3) — Política N0–N3 y normalización
- Normalización con `Trim()` y sanitización de separadores en `GetSolicitudes` y `NormalizarDocumento`.
- Implementación de pipeline N0 (dedupe `(Num, CC)`), N1 (metadatos), N2 (OCR validado) y N3 (ambigua sin tocar).

### Fase 1440 (T4) — `FindCC` endurecido contra cédulas candidatas
- `FindCC` restringido para validar únicamente secuencias numéricas que pertenezcan al conjunto de cédulas candidatas de la comisión, descartando rubros presupuestales u otros códigos ajenos.

### Fase 1441 (T5) — Suite de pruebas de regresión
- 7 pruebas unitarias nuevas en `ViaticosGestionProcessorTests`:
  1. Renombrado exitoso en subcarpeta `Solicitudes`.
  2. Dedupe de comisión duplicada con misma cédula sin invocar OCR.
  3. Desambiguación por OCR ante comisiones compartidas con cédulas distintas.
  4. Protección ante OCR con número ajeno (clasificado como `AMBIGUA`).
  5. Detección de colisión sin sobrescritura de archivo existente.
  6. Detección y clasificación de archivos con prefijo no numérico.
  7. Manejo defensivo ante carpeta inexistente.

### Fase 1442 (T6) — Documentación y cierre
- Actualización de `gestion-documental.md` con política N0–N3 y catálogo de motivos.
- Sincronización de esquema de API DocFX y cierre de fila `R-003` en `REGISTRY.md`.

