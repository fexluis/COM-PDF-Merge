# Corrección: Renombrar y Unificar Solicitudes con Rango de Comisiones #YY

**Fecha**: 2026-10-30 · **Componente**: `ConversorPDF.XLL/Core/Processors/ViaticosGestionProcessor.cs`

## Diagnóstico y Propósito

1. **D1 - Substring de nombre en `Clean`**: En `ViaticosGestionProcessor.Clean()`, `Substring(0, Math.Min(s.Length, 18))` utilizaba la longitud de la cadena de entrada `s` (con espacios) en lugar de la longitud de la cadena filtrada `limpio`, lo que provocaba un `ArgumentOutOfRangeException` en nombres extensos de funcionarios e impedía la carga de solicitudes.
2. **D2 - Rango descartado en `ViaticosService`**: Los métodos `RenombrarSolicitudes(inicio, fin)` y `UnificarSolicitudes(inicio, fin)` descartaban los parámetros de rango y delegaban llamadas sin argumentos al procesador.
3. **D3/D4 - Flujo silencioso en errores o lista vacía**: Al filtrar solicitudes y no obtener coincidencias, el procesador finalizaba en silencio o capturaba excepciones sin avisar al usuario.
4. **D5 - Validación de entrada en formulario**: `FormGestionDocumental` realizaba `int.Parse` sin control de excepciones ni validación de orden (`inicio > fin`).
5. **D6 - Unificación sin filtro**: `UnificarSolicitudes` procesaba todas las hojas sin considerar el rango solicitado en la interfaz.

## Convención de Rango `#YY`

Las comisiones de viáticos en SIIF Nación siguen la convención `[Consecutivo][Año2Digitos]` (ej. `126` = No. 1 de 2026, `14526` = No. 145 de 2026, `12625` = No. 126 de 2025). La evaluación de rango se realiza mediante la tupla `(año, consecutivo)` para evitar que años anteriores queden incluidos incorrectamente.

## Cambios Implementados

1. **`ConversorPDF.XLL/Core/Helpers/ViaticoNumeroComision.cs`** (nuevo)
   - Parseo y evaluación de comisiones con la regla `#YY`.
   - Interpretación de límites de < 3 dígitos o sin sufijo de año como consecutivos de la vigencia actual (`App.APP_VIGENCIA`).
   - Normalización de rangos invertidos.
2. **`ConversorPDF.XLL/Core/Processors/ViaticosGestionProcessor.cs`**
   - Corrección en `Clean()`: `limpio.Substring(0, Math.Min(limpio.Length, 18))`.
   - Soporte de filtro `(inicio, fin)` en `RenombrarSolicitudes` y `UnificarSolicitudes`.
   - Mensajes de advertencia al usuario vía `MessageBox` ante colecciones vacías o errores.
   - `GetSolicitudes` expuesto como `internal` para pruebas unitarias sin dependencia de interfaz gráfica.
3. **`ConversorPDF.XLL/Core/Interfaces/IViaticosGestionProcessor.cs`**
   - Actualización de firmas con parámetros de rango `(string folder, int inicio, int fin)`.
4. **`ConversorPDF.XLL/Core/Services/ViaticosService.cs`**
   - Propagación de los parámetros `inicio` y `fin` al procesador.
5. **`ConversorPDF.XLL/FormProgramador/FormGestionDocumental.cs`**
   - Implementación de `TryLeerRango` con validación `int.TryParse` y regla `inicio <= fin`.
6. **`ConversorPDF.XLL/ConversorPDF.XLL.csproj`**
   - Registro de `ViaticoNumeroComision.cs` en `<Compile Include>`.

## Pruebas

- **`ConversorPDF.Tests/Core/Helpers/ViaticoNumeroComisionTests.cs`**: Casos de `TryParse` (formatos `#YY`, vigencia actual, no numéricos) y `EnRango` (inclusivo, sin filtro, consecutivos puros, años distintos, invertido).
- **`ConversorPDF.Tests/Core/Processors/ViaticosGestionProcessorTests.cs`**: Prueba de regresión de nombre largo con espacios (219 caracteres) y verificación del filtrado por rango sobre datos 1-based de `CONSULTAS`.

