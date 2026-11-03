---
numbering:
  registry: .agents/plans/REGISTRY.md
  registry_row: R-002
  ronda_base: 264
  fase_base: 1426
  waves: 2
  tasks: 11
---

# Plan: Corregir RenombrarSolicitudes y hacer efectivo el rango #YY

## Resumen Ejecutivo
Implementar soporte para rango de comisiones `#YY` (formato SIIF Nación `[Consecutivo][Año2Dígitos]`) en `RenombrarSolicitudes` y `UnificarSolicitudes`, desacoplar la interfaz gráfica eliminando `MessageBox` de la capa Core y manejar códigos de retorno estructurados.

## Tasks & Fases

### T1: Helper de dominio puro para convención #YY
- **fase**: 1426
- Archivo: `ConversorPDF.XLL/Core/Helpers/ViaticoNumeroComision.cs`
- Métodos: `TryParse`, `TryParseLimite`, `EnRango`, `HayFiltro`.

### T2: Corrección Clean() y filtro en GetSolicitudes
- **fase**: 1427
- Archivo: `ConversorPDF.XLL/Core/Processors/ViaticosGestionProcessor.cs`
- Corrección de `ArgumentOutOfRangeException` y filtrado por tupla `(año, consecutivo)`.

### T3: Filtro en UnificarSolicitudes
- **fase**: 1428
- Archivo: `ConversorPDF.XLL/Core/Processors/ViaticosGestionProcessor.cs`
- Filtrado contra `P26` (`order`) y retorno `-1` en caso de fallo.

### T4: Contratos e interfaces
- **fase**: 1429
- Archivos: `IViaticosGestionProcessor.cs`, `IViaticosService.cs`, `ViaticosService.cs`
- Firmas `int RenombrarSolicitudes(string folder, int inicio, int fin)` y `int UnificarSolicitudes(string folder, int inicio, int fin)`.

### T5: Validación defensiva y feedback en UI
- **fase**: 1430
- Archivos: `FormGestionDocumental.cs`, `RibbonController.cs`
- `TryLeerRango` con `int.TryParse`, validación `inicio <= fin` y presentación de resultados en vista.

### T6: Pruebas unitarias para ViaticosGestionProcessor
- **fase**: 1431
- Archivo: `ConversorPDF.Tests/Core/Processors/ViaticosGestionProcessorTests.cs`
- Pruebas xUnit de regresión para nombre largo (219 chars) y filtrado de matriz 1-based.

### T7: Retiro de MessageBox en capa Core y desacople WinForms
- **fase**: 1432
- Archivo: `ConversorPDF.XLL/Core/Processors/ViaticosGestionProcessor.cs`
- Supresión de dependencias de `System.Windows.Forms` en capa Core.

### T8: Retorno de error -1 al unificar y propagación en Ribbon
- **fase**: 1433
- Archivos: `ViaticosGestionProcessor.cs`, `FormGestionDocumental.cs`, `RibbonController.cs`
- Distinción entre expedientes vacíos (`0`) y fallo de infraestructura (`-1`).

### T9: Desacople dinámico de vigencia en tests
- **fase**: 1434
- Archivo: `ConversorPDF.Tests/Core/Helpers/ViaticoNumeroComisionTests.cs`
- Pruebas con año dinámico derivado de `App.APP_VIGENCIA`.

### T10: Actualización de metadatos API docfx
- **fase**: 1435
- Archivo: `docs/api/ConversorPDF.XLL.Core.Interfaces.IViaticosGestionProcessor.yml`
- Sincronización de esquemas con retorno `int`.

### T11: Registro y trazabilidad en REGISTRY.md
- **fase**: 1436
- Archivo: `.agents/plans/REGISTRY.md`
- Asignación formal de contadores para Ronda 265 / Fase 1436.

