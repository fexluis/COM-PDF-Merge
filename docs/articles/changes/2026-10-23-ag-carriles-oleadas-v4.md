# Carriles disjuntos y oleadas secuenciales (v4) — Fase 1279

## Purpose
Reversión del modelo de bandas paralelas (Fase 1226) a un modelo determinista de carriles disjuntos con oleadas secuenciales temporales por carril (Fase 1279). En F1226, la partición simultánea forzaba a todas las asignaturas a compartir el mismo inicio de ventana (`iniBanda = fechaInicio`, `Oleada = 0`), provocando apilamiento masivo de asignaturas por semana e incompatibilidad estructural con las restricciones duras R5 (máximo de asignaturas simultáneas por semana) y R13 (dispersión/solapamiento indebido).

## Changes
1. **`AgPlanificadorOleadas.cs`**:
   - Sustitución de `AgBandasDias.Resolver` por el algoritmo en 3 pasadas:
     - **Pasada 1**: Asignación obligatoria por patrones sembrados (`dias_patron`). Selección del carril cuyo conjunto de días sea el superconjunto mínimo (ej. resuelve ambigüedad de SÁB frente a {JUE,VIE,SÁB}).
     - **Pasada 2**: Reparto greedy por utilización relativa `(carga[c] + S) / cap[c]`, equilibrando el makespan y preservando el orden curricular.
     - **Pasada 3**: Cursor temporal secuencial dentro de cada carril, avanzando `k = ceil(S / slots[c])` semanas contiguas ajustadas por festivos mediante `AgPlanificadorPatronSemanal.Construir`.
   - Eliminación del helper obsoleto `RanurasEfectivas`.
2. **`AgBandasDias.cs`**:
   - `AgBandasDias.Resolver` documentado como `[Obsolete documental: legado F1226; el camino canónico es carriles + oleadas (F1279)]`.
3. **Versión**:
   - Incremento a `1.6.3.70` en `ConversorPDF.XLL.csproj` y `Core/Helpers/App.cs`.

## How to use
El uso de `AgPlanificadorOleadas.Construir(...)` mantiene su firma intacta:
- `AgPlanOleada.Carril` transporta el índice de carril (`0..carriles-1`).
- `AgPlanOleada.Oleada` transporta el orden secuencial (`0, 1, ...`) dentro del carril.
- Cada asignatura de un mismo carril queda en una sub-ventana contigua y estrictamente disjunta respecto a sus compañeras de carril (`secuencia[i].FechaInicio > secuencia[i-1].FechaFin`).

### Caso de referencia: prog 4 (aula JUE-SAB, MANANA, 4h, 2027-1)
| Carril | Días | Ranuras/sem | Asignaturas | Semanas |
|---|---|---|---|---|
| 0 | JUE, VIE | 2 | 1 (48h) → 3 (48h) → 5 (32h) | 1-6, 7-12, 13-16 |
| 1 | SÁB (M+T) | 2 | 2 (48h) → 4 (32h) → 6 (32h) | 1-6, 7-10, 11-14 |

## Test coverage
- `Construir_When4AsignaturasAnd16SemanasMarVie_ShouldGenerateDisjointLanesAndSequentialWaves`: valida 2 carriles, oleadas 0 y 1 disjuntas en el tiempo.
- `Construir_CuandoFestivosEstrechanLaOleadaFinal_NoMarcaInfactible`: valida secuenciación y márgenes temporales correctos ante festivos.
- `Construir_When6AsignaturasJueSab_ShouldSecuenciarOleadasPorCarrilYNoApilar`: regresión de prog 4 sin apilamiento semanal (R5 respetada).
- `Construir_WhenCarrilesDesbalanceados_ShouldRepartirProporcionalALaUtilizacion`: balance de carga proporcional a la utilización de los carriles.
- `Construir_WhenDemandaExcedeCapacidadDeUnCarril_ShouldMarcarInfactibleConDetalle`: detección y reporte detallado de sobrecarga por patrones.
- `Construir_ConFestivoDentroDeLaOleada_ShouldExtenderUnaSemanaSinPerderSesiones`: extensión de semanas ante festivos intermedios.
- `Construir_WhenVentanaDe2Dias_ShouldUnSoloCarril`: verificación de carril único para ventanas de 2 ranuras.
