# Colisión de Nombre de Función Excel 'De' en Excel-DNA 1.8.0

**Fecha:** 2026-10-22  
**Módulo:** Académico (`ConversorPDF.XLL.Academic`)  
**Ronda:** 203 (Fases 1033 a 1038)

## Propósito
Resolver la colisión de nombres de función registrada automáticamente en Excel-DNA 1.8.0 generada por dos métodos públicos estáticos con el nombre `De`:
1. `AgSemanaIso.De(DateTime)` (`ConversorPDF.XLL.Academic.Application.Services.Genetica`)
2. `AgClaveCohorte.De(int, int, int, string, string)` (`ConversorPDF.XLL.Academic.Domain.Models.Genetica`)

En Excel-DNA 1.8.0, el registro por defecto procesa todas las clases públicas y exporta métodos estáticos públicos compatibles sin usar el prefijo de la clase en el nombre de la macro/UDF de Excel. Cuando dos métodos en clases distintas coinciden en nombre, Excel-DNA emite un mensaje de error en el log de registro:
`Logger.Registration.Error("Repeated function name: '{0}' - previous registration will be overwritten. ")`
donde `name = mi.Name` (sin forzar mayúsculas: `'De'`), sin lanzar excepción y sin deshabilitar el registro general: la última función registrada sobrescribe a la anterior en `XlRegistration.cs`.

## Cambios
1. **Eliminación del alias `De(DateTime)` (`AgSemanaIso.cs`, Fase 1033):**
   - Se eliminó el método `AgSemanaIso.De(DateTime)` que actuaba como alias de `DiaIso(DateTime)`.
   - Se mantiene como API canónica y unificada `AgSemanaIso.DiaIso(DateTime)`.

2. **Renombrado semántico y ajuste de visibilidad en `AgClaveCohorte` (`AgClaveCohorte.cs`, Fases 1033 y 1036):**
   - `De(int idSede, int idPrograma, int semestre, string grupo, string aulaRef = null)` pasa a ser `internal static string DesdeCohorte(...)`. Al ser un helper de dominio con parámetro por defecto (`aulaRef = null`), se protege con `internal` para evitar que Excel-DNA 1.8.0 lo registre como UDF exigiendo obligatoriamente 5/6 argumentos en Excel.
   - `De(AgAula aula)` pasa a ser `DesdeAula(AgAula aula)`.
   - `De(AgGen gen)` pasa a ser `DesdeGen(AgGen gen)`.

3. **Actualización de invocación en fábrica genética (`AgFabricaIndividuos.cs`, Fase 1033):**
   - Actualizada la función local `ClaveCohorteAula(AgAula a)` para invocar `AgClaveCohorte.DesdeAula(a)`.

4. **Test guardián de registro en Excel-DNA (`ExcelDnaRegistrationGuardTests.cs`, Fases 1034 y 1037):**
   - Suite de pruebas automatizadas en `ConversorPDF.Tests\Unit` que inspecciona por reflexión el ensamblado `ConversorPDF.XLL`.
   - Traslada fielmente las reglas de `AssemblyLoader.GetExcelMethods` y `IsMethodSupported` de Excel-DNA 1.8.0 (incluyendo accesores de propiedad `get_`/`set_` estáticos públicos y aplicando los filtros de clases, tipos genéricos abiertos y arrays).
   - Verifica que ningún método registrable repita nombre (`OrdinalIgnoreCase`) en todo el ensamblado.
   - Verifica que ningún método registrable se llame `De`.

## Cómo usar y Compatibilidad
- Para obtener el día de la semana ISO (Lunes=1 .. Domingo=7), invocar `AgSemanaIso.DiaIso(fecha)`.
- Para construir la clave de cohorte a partir de un aula o gen, invocar `AgClaveCohorte.DesdeAula(aula)` o `AgClaveCohorte.DesdeGen(gen)`.
- Para construir la clave de cohorte a partir de sus componentes primarios en código C#, invocar `AgClaveCohorte.DesdeCohorte(idSede, idPrograma, semestre, grupo, aulaRef)`.
- En Excel, la función canónica `=DIAISO(...)` continúa operativa sin colisiones ni advertencias en el registro.
- **Advertencia de Compatibilidad:** Cualquier fórmula preexistente en hojas o libros de usuario que haya intentado invocar `=De(...)` pasará a evaluar como `#NAME?` al haberse eliminado la función duplicada de la tabla global de macros/UDFs.

## Cobertura de pruebas (Fases 1034 y 1037)
- `ExcelDnaRegistrationGuardTests.cs`:
  - `MetodosRegistrables_WhenInspectedUnderExcelDna180Rules_ShouldNotHaveDuplicateNames`
  - `MetodosRegistrables_WhenInspectedUnderExcelDna180Rules_ShouldNotContainMethodNamedDe`
