# Sistema de Carriles Disjuntos y Oleadas Secuenciales en el Algoritmo Genético

**Fecha:** 2026-10-22  
**Módulo:** Académico (`ConversorPDF.XLL.Academic`)  
**Ronda:** 206 (Fases 1052 a 1065)

## Propósito
Resolver de forma estructural e integral la viabilidad, calidad pedagógica y consistencia normativa de las mallas horarias generadas por el motor genético académico:
1. **Persistencia de Decisiones de Selección (I1):** Garantizar que las dimensiones elegidas en el árbol jerárquico (`Jornada`, `HorasDia`, `DiasVentana`) se persistan fielmente en `ag_aula` y `ag_aula_asignatura` (mediante operaciones `UPDATE` y `ON CONFLICT DO UPDATE`), eliminando inferencias tardías o discrepancias de configuración.
2. **Carriles Disjuntos de 2 Días (I2, I3):** Descomponer la ventana semanal $\beta$ en pares disjuntos de 2 días (ej. `MAR-VIE` $\rightarrow$ Carril 0: `[2,3]`, Carril 1: `[4,5]`), asignando cada asignatura mediante rotación modular `carril = idxAsignatura mod carriles`. En ventanas impares (ej. `VIE-DOM` `[5,6,7]`), el día residual no entra en rotación principal.
3. **Oleadas Temporales Secuenciales (I4):** Dividir el semestre en sub-ventanas temporales `[FechaInicio, FechaFin]` para cada carril mediante `oleada = idxAsignatura div carriles`. Las asignaturas de la misma cohorte que comparten un carril se encadenan en oleadas sucesivas sin solapamiento temporal.
4. **Erradicación de Degradación Silenciosa a 1 Día:** Eliminar el fallback que reducía una asignatura a un solo día cuando el planificador de patrón semanal fallaba, lo que multiplicaba sesiones en el mismo día o generaba sobrecargas. Ahora, ante un plan infactible, la asignatura se marca explícitamente como vacante con motivo `SIN_DISPONIBILIDAD`.
5. **Erradicación de Clases de 8 Horas en Pregrado (D4/I6):** Blindar la duración de sesiones a 4.0 horas en todos los programas de pregrado. Las clases de 8.0 horas quedan estrictamente reservadas para posgrados con `HorasDia == 8`. Se sembraron bloques de 4h faltantes mediante la migración V74 y se eliminaron todos los fallbacks no acotados de selección de bloques en la fábrica y en la mutación.
6. **Adaptación de Restricciones R5 y R6:**
   - **R5:** La concurrencia de asignaturas se evalúa ahora por semana ISO contra la cantidad de carriles concurrentes del aula, permitiendo que una cohorte curse hasta $N$ asignaturas simultáneas si dispone de $N$ carriles paralelos.
   - **R6:** La secuencia temporal de asignaturas se evalúa agrupando por carril (`DiasPatron`), permitiendo que asignaturas en carriles disjuntos se dicten en paralelo en las mismas fechas sin marcar violaciones duras.
7. **Compuertas de Seguridad (I5, I8):** Verificación matemática temprana de capacidad de ranuras en pre-vuelo (`I5`), y advertencia/bloqueo de seguridad al exportar o promover programaciones que presenten restricciones duras no resueltas (`HardViolations > 0`).

---

## Cambios por Componente

### 1. Dominio (`ConversorPDF.XLL.Academic.Domain.Models.Genetica`)
- **`AgCarriles.cs`:**
  - Clase estática interna (`internal static class`) para neutralizar colisiones de exportación de UDFs en Excel-DNA 1.8.0.
  - `Ventana(string diasVentana)`: Normalización y parseo seguro tanto de códigos canónicos (`"MAR-VIE"`) como de listas numéricas ISO (`"2,3,4,5"`).
  - `Cantidad(IReadOnlyList<int> ventana)`: Cálculo de carriles como $\max(1, |\beta| / 2)$.
  - `Indice(int idxAsignatura, int carriles)`: Rotación modular `idxAsignatura % carriles`.
  - `DelCarril(IReadOnlyList<int> ventana, int idxAsignatura)`: Extracción de la sub-lista de 2 días correspondiente al carril rotado.
  - `Residual(IReadOnlyList<int> ventana)`: Identificación del día sobrante en ventanas impares.
- **`AgPlanOleada.cs`:**
  - POCO inmutable (`internal sealed class`) con propiedades `IdAsignatura`, `Carril`, `Oleada`, `DiasPatron`, `FechaInicio`, `FechaFin`, `SemanasOleada`, `SesionesRequeridas` y `CapacidadRanuras`.

### 2. Aplicación (`ConversorPDF.XLL.Academic.Application`)
- **`AgPlanificadorOleadas.cs`:**
  - Servicio determinista para calcular rebanadas de semanas útiles en la ventana semestral, distribuir las semanas entre oleadas sucesivas y acotar las fechas `[FechaInicio, FechaFin]` de cada asignatura.
  - Validación de capacidad de ranuras (Invariante I5) con retorno `Result.Failure` si el calendario no ofrece ranuras suficientes para las sesiones requeridas.
- **`AgSeleccionMapper.cs`:**
  - Mapeador seguro entre `AgNodoArbolKey` y `AgContextoSeleccionDto`.
  - Valida completitud de `Jornada` (no vacía), `HorasDia` (estrictamente 4 u 8) y `DiasVentana` (al menos un día ISO válido), conservando el código canónico (ej. `"MAR-VIE"`).
- **`AgPlanificadorPatronSemanal.cs`:**
  - `DiasKDe(...)`: Delegación a `AgCarriles.DelCarril` para asignación de días de carril.
  - Normalización de patrones multi-día legados mayores a 2 días para preservar el invariante de 2 días por carril.
- **`AgValidadorVentana.cs`:**
  - Validación obligatoria de dimensiones de decisión (`diasVentana`, `jornada`, `horasDia`) en cada aula activa.
  - Verificación de factibilidad de capacidad total (I5) sumando ranuras requeridas vs disponibles en la ventana.
- **`AgValidadorReglas.cs`:**
  - **R5:** Verificación de concurrencia por semana ISO acotada a la cantidad de carriles del aula.
  - **R6:** Verificación de orden y solape agrupando asignaturas por carril (`DiasPatron`).
  - **R11:** Regla dura que penaliza genes cuya duración de sesión no corresponde a la requerida por el nivel académico (4h pregrado, 8h posgrado).
  - **R13:** Regla dura que penaliza solapamientos temporales entre asignaturas asignadas al mismo carril.
- **`AgOperadorMutacion.cs`:**
  - Mutación acotada para respetar la duración esperada según nivel académico y mantener los genes dentro de los días de su carril asignado.
- **`AgExpansorSesiones.cs`:**
  - Defensa en profundidad: verificación estricta de que la duración del bloque coincide con `AgUnidadesSemanales.HorasSesionAula`, retornando fallo temprano si difieren.
- **`AgFabricaIndividuos.cs`:**
  - Ordenamiento canónico por `(OrdenCurricular, IdAsignatura)` para cada cohorte/aula.
  - Integración pre-vuelo de `AgPlanificadorOleadas`: construcción de sub-ventanas y carriles antes de generar genes.
  - Erradicación de la degradación silenciosa a 1 día: si el plan es infactible, el gen se crea con `EsVacante = 1` y motivo `SIN_DISPONIBILIDAD`.
  - Filtro estricto de bloques por duración requerida `horasEfectivas`, eliminando cualquier fallback a bloques de 8h en pregrado.

### 3. Infraestructura (`ConversorPDF.XLL.Academic.Infrastructure`)
- **`MigrationV74_SembrarBloques4hFaltantes.cs`:**
  - Migración versión 74 que garantiza la existencia de bloques activos de 4.0h para los días 1 a 6 (Lunes a Sábado) en la tabla `ag_bloque_horario`.
  - Registrada en `SqliteHelper.cs`.

### 4. Interfaz de Usuario (`ConversorPDF.XLL.Academic.UI.Forms`)
- **`AgFormGeneradorHorarios.cs`:**
  - `ObtenerCombinacionesSeleccionadas`: consumo de `AgSeleccionMapper.DesdeNodo` para validar las 3 dimensiones de selección del árbol.
  - `GenerarEncabezadosAulasAsync`:
    - Operación `UPDATE` sobre `ag_aula` para sincronizar `jornada`, `horas_dia` y `dias_ventana` cuando el aula ya existe en la base de datos.
    - Consulta `sqlAsigs` con `COALESCE(aa.orden_curricular, 0) AS OrdenCurricular` y ordenamiento determinista `ORDER BY COALESCE(aa.orden_curricular, 0), a.id`.
    - Inserción en `ag_aula_asignatura` mediante `ON CONFLICT(id_aula, id_asignatura) DO UPDATE SET dias_patron = ..., horas_objetivo = ...`.
    - Asignación de `diasPatron` mediante `AgCarriles.DelCarril(ventanaIso, idxAulaAsig)`.
  - `AsigSeedRow`: adición de propiedad `OrdenCurricular`.
  - `ExportarResultadosExcel` y `ExportarProyeccionEspejo`: compuerta de seguridad I8 que advierte al usuario si la solución activa contiene violaciones duras (`HardViolations > 0`).

### 5. Suite de Pruebas Unitarias (`ConversorPDF.Tests`)
- **`AgCarrilesTests.cs` (T1 a T5):** Cobertura exhaustiva de partición de carriles par/impar, días residuales y parseo de ventanas.
- **`AgPlanificadorOleadasTests.cs` (T6 a T9):** Cobertura de distribución de 4 asignaturas en 2 oleadas y 2 carriles en 16 semanas, fallo por capacidad insuficiente (I5), ventana vacía y lista vacía.
- **`AgSeleccionMapperTests.cs` (T10 a T12):** Cobertura de mapeo fiel de dimensiones, validación de horas/día (4 u 8), validación de ventana no vacía y jornada no vacía.
- **`MigrationV74Tests.cs` (T13 a T14):** Metadatos de migración V74 y verificación de siembra idempotente de bloques de 4h para días 1..6.

---

## Verificación de Calidad y Criterios de Aceptación

| Criterio | Estado | Observación |
|---|:---:|---|
| **I1: Persistencia de ventana $\beta$** | ✅ CUMPLIDO | Persistida en `ag_aula.dias_ventana` vía `UPDATE` / `AgAulaFactory`. |
| **I2: Carril = par disjunto de 2 días** | ✅ CUMPLIDO | Calculado por `AgCarriles.DelCarril` ($\max(1, |\beta|/2)$). |
| **I3: Rotación modular por orden curricular** | ✅ CUMPLIDO | Orden `(OrdenCurricular, Id)` en `sqlAsigs`, fábrica y planificador. |
| **I4: Oleadas temporales secuenciales** | ✅ CUMPLIDO | Sub-ventanas no solapadas por carril mediante `AgPlanificadorOleadas`. |
| **I5: Pre-vuelo de capacidad** | ✅ CUMPLIDO | Validador comprueba ranuras suficientes contra demanda horaria. |
| **I6: Duración de sesión según nivel** | ✅ CUMPLIDO | 4h obligatorias en pregrado (R11, fábrica, mutación, expansor). |
| **I7: Sin degradación silenciosa** | ✅ CUMPLIDO | Fallo de patrón se marca como vacante `SIN_DISPONIBILIDAD`. |
| **I8: Compuerta de exportación y promoción** | ✅ CUMPLIDO | Advertencia/bloqueo ante `HardViolations > 0`. |
| **Arquitectura Limpia & .NET 4.8** | ✅ CUMPLIDO | Tipos internos/sellados, C# 7.3, Single Type Per File, sin fuga de dependencias. |
