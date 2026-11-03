# Bitácora: Duplicados en `_Duplicados` y auditoría de INFORME (columna H)

- **Fecha**: 2026-10-31
- **Plan**: `R-006` · Ronda 279–281 · Fases 1483–1487
- **Sustituye**: `R-005` (verificación base) → `abandoned`

## 1. Diagnóstico

| Hecho | Evidencia |
|---|---|
| Renombrado operativo | corrida 14:03: `renombrados=0 · yaCorrectos=381 · colisiones=14 · sinComision=28 · prefijoInvalido=11` |
| 14 "colisiones" reales | duplicados exactos: el destino ya existía con el mismo contenido |
| 2 prefijos inválidos | `202426` → `20426` (ILDELBRANDO AREVALO) y `202626` → `20626` (GERMAN LUIS ROLDAN) |
| 23 sin comisión | `41426`–`43626`: no existen en `CONSULTAS` |
| H mal marcada | 2 filas con `SIN SOLICITUD` **con** archivo presente; 1 variante `FALTA SOLICITUD` |
| Auditoría | inexistente antes de esta ronda |

## 2. Cambios

### Fase 1483 — Duplicados
- `App.FolderNames.DUPLICADOS = "_Duplicados"`.
- `MotivoRenombrado.Duplicado`, `ResultadoRenombrado.Duplicados`.
- `RenombrarFila`: si el destino existe y el contenido es idéntico → mueve el origen a `Solicitudes\_Duplicados` (sufijo `_<6hex>` si ya existe allí). Si el contenido difiere → `COLISION`.
- `MismoContenido`: comparación por longitud + bytes.

### Fase 1484 — Auditoría de presencia física
- `EstadoComision` y `ResultadoVerificacionSolicitudes`.
- `IExcelRepository.SaveActiveWorkbook()` + implementación.
- `VerificarSolicitudesPresentes(folder, inicio, fin)`:
  - índice `comisión|cédula` desde raíz + `Solicitudes` (sin `_Duplicados`),
  - protección `VERIFICADO`/`REINTEGRO`,
  - copia de estados terminales (`REVOCADA`, `RECHAZADA`, `ANULADA`),
  - limpieza de marcas erróneas cuando el archivo existe,
  - `SIN SOLICITUD` cuando no existe,
  - preservación de otras marcas (`ERROR`, `SIN INFORME`, `DUPLICADA`),
  - `ScreenUpdating(false)` + `finally`,
  - guardado del libro cuando `CeldasEscritas + CeldasLimpiadas > 0`.

### Fase 1485 — Normalización y remapeo
- Remapeo de prefijos de 6 dígitos probando la eliminación de una posición (1–3) con candidato único y coincidencia tolerante de nombre (`CoincideNombre`, prefijos de 5+ letras sin acentos).
- Comparaciones `Norm()` (trim + mayúsculas) y variante `FALTA SOLICITUD`.

### Fase 1486 — Pruebas
- `ViaticosGestionProcessorTests`: 12 casos nuevos (duplicado, colisión, remapeo, prefijo sin coincidencia, rango, marcado, limpieza, protegidos, estados terminales, otras marcas, idempotencia, guardado).

### Fase 1487 — Documentación y cierre
- `docs/articles/gestion-documental.md`: secciones de duplicados y auditoría.
- Registro `R-006` y contadores globales.
