# Habilitar "Usar Autofiltro" en Hoja Protegida de Exportar Asignaturas

**Fecha:** 2026-10-17  
**Módulo:** Académico (`ConversorPDF.XLL.Academic`)  
**Rondas:** 134 (Fase 728), 135 (Fases 729 a 732), 136 (Fases 733 a 735), 137 (Fases 736 a 738), 138 (Fases 739 y 740) y 139 (Fases 741 a 743)

## Purpose
Garantizar que tras la exportación de vinculaciones docentes por asignaturas (`FV-FO-033`), los usuarios puedan utilizar las capacidades de autofiltro en la tabla `T_Principal8` (fila 13, rango `A13:DF234`) de la hoja `PROGRAMACIÓN` sin que Excel arroje la alerta de celda/gráfico protegido, manteniendo protección contra escritura, blindaje con fallback trazable ante excepciones de invocación COM, y control estricto para solo reproteger hojas que fueron efectivamente desprotegidas por el sistema.

## Changes
1. **`ConversorPDF.XLL\Academic\Infrastructure\Excel\AsignaturasExcelExporter.cs`:**
   - Incorporación del método auxiliar privado estático `ProtegerHojaPermitiendoAutofiltro(dynamic ws, string password)`.
   - Invocación de `Worksheet.Protect` mediante reflexión COM tardía (`InvokeMember`) pasando un arreglo de 16 argumentos posicionales con `Missing.Value`.
   - Activación explícita del permiso posicional `AllowFiltering` (`args[14] = true`, Fase 728), preservando la contraseña interna (`ConstantsHelper.FORMATO_PASSWORD`).
   - Reversión de `AllowSorting` (Fase 733): se conserva en `Missing.Value` (deshabilitado) dado que Excel exige que todas las celdas del rango estén desbloqueadas para permitir ordenamiento en hojas protegidas.
   - Bandera `hojaDesprotegida` (Fases 734/736): `true` inmediatamente tras `Unprotect` exitoso; se invalida a `false` si `ProtectContents` persiste activo (contraseña distinta). El bloque `finally` únicamente ejecuta reprotección si la bandera es `true`.
   - Trazabilidad y observabilidad en fallback (Fases 730 / 735): ante fallo en la invocación reflexiva, se registra advertencia mediante `LoggerHelper.Warning` y se aplica fallback con `ws.Protect(...)`. Si el fallback también falla, se captura y emite error detallado con `LoggerHelper.Error`.
2. **`installer\InstaladorConversorPDF.wxs` y Workflow de Release (Fases 739, 741, 742):**
   - Actualización de versión a `1.6.3.57` en manifiesto WiX, `App.cs` y `.csproj`.
   - Enlace automático dinámico de versión de plantillas mediante `DefaultVersion="!(bind.FileVersion.filXllDll)"` en `filTemplate1` y `filTemplate2`.
   - Sincronización de tamaños de plantillas en checklist de `build-release.md` (`FV-FO-033 = 777082` B, `GF-FO-023 = 2938592` B).

## Invariante de Rango de Datos (Fase 731)
- La tabla de datos en la plantilla `FV-FO-033` (`T_Principal8`) posee un rango nativo de 221 filas de datos (`A13:DF234`).
- Las exportaciones que se encuentren dentro de dicho rango operan con soporte íntegro de fórmulas, formatos de tabla y autofiltro nativo.
- Por diseño no destructivo, no se fuerza redimensionamiento dinámico de la tabla (`ListObject.Resize`) en tiempo de exportación para proteger la integridad estructural de las fórmulas preexistentes en columnas calculadas.

## How to use
1. En Excel, abrir el complemento **Conversor Académico**.
2. Navegar a **Ver Vinculaciones** y seleccionar la vista **Asignaturas**.
3. Hacer clic en **Exportar**.
4. Al abrirse el archivo generado (`FV-FO-033-<timestamp>.xlsx`), la hoja `PROGRAMACIÓN` permanece protegida contra edición de celdas bloqueadas.
5. Hacer clic en las flechas de desplegable de filtro en la fila 13: los filtros se abren y aplican criterios sin advertencia de bloqueo. El ordenamiento desde el desplegable requiere celdas desbloqueadas y no está habilitado en esta plantilla (Fase 733).

## Test Coverage (Fase 732, 733, 735)
- **Análisis estático:**
  - `Protect` vía `InvokeMember` con 16 argumentos posicionales COM.
  - `args[0] = password`, `args[14] = true` (`AllowFiltering`), `args[13] = Missing.Value` (`AllowSorting` desactivado).
  - Trazabilidad con `LoggerHelper.Warning` y `LoggerHelper.Error` en bloque de protección.
  - Condición de reprotección vinculada a `hojaDesprotegida`.
- **Smoke test manual (Excel 365 x64 - 2026-10-17 / build 1.6.3.57):**
  - Desplegable de filtro en fila 13 sobre `T_Principal8` (`A13:DF234`) permite seleccionar y filtrar valores sin advertencia de celda protegida.
  - Celdas bloqueadas de la hoja permanecen protegidas contra escritura directa.
  - Menú "Desproteger hoja" valida la contraseña interna definida en `ConstantsHelper.FORMATO_PASSWORD`.
  - Caso borde documentado: exportaciones > 221 filas (filas 235+) quedan fuera del rango de `T_Principal8` (ver Fase 731).
