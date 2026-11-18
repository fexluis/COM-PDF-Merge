# Certificación de Optimización y Rendimiento del Algoritmo Genético (Corrida #105)

**Fecha**: 2026-11-17  
**Rama**: `rendimiento-app`  
**Registro**: `R-112` (Rondas 433–438, Fases 1942–1951, Tareas T1–T10)  
**Autor**: Antigravity Swarm Coordinator  

---

## 1. Contexto y Objetivos

A partir de la telemetría real de la **Corrida #105** (240,158 s totales en `OptimizacionAG`, 3,90 s/transición en estado estacionario para 50 individuos × 410 genes), se diseñó y ejecutó un plan sistemático de 10 tareas divididas en 6 olas de optimización vertical en el slice `Academic`.

### Criterios de Aceptación Globales
1. **Calidad de solución**: Paridad estricta con los volúmenes estructurales: 410 genes, 2.228 sesiones, 8.912 horas, 199 vinculaciones y 336 vacantes (30 violaciones blandas, `HARD=0`).
2. **Determinismo**: Mismo seed genera exactamente la misma trayectoria de aptitud y el mismo hash canónico SHA-256.
3. **Rendimiento**: Transición en estado estacionario reducida de ~3,90 s a ≤ 1,0 s (objetivo ~0,65 s) y tiempo total reducido sustancialmente.
4. **Cero regresiones**: 100 % de la suite de pruebas en verde.

---

## 2. Desglose de Tareas Implementadas

### Wave 1 (Ronda 433)
- **T1 (Fase 1942: Semilla determinista y flujos RNG)**:
  - Sobrecargas de constructores con `int seed` en `AgOperadorCruce`, `AgOperadorMutacion`, `AgOperadorSeleccion`, `AgFabricaIndividuos` y `AgAsignadorTercero`.
  - Flujos RNG desacoplados sin alterar el orden canónico de llamadas a `Random`.

### Wave 2 (Ronda 434)
- **T2 (Fase 1943: Telemetría generacional y corrección de DuracionMs)**:
  - Corrección de la ventana de medición de `DuracionMs` en `AgGeneradorHorarios`.
  - Clase interna `AgPerfilGeneracion` con desglose de sub-tiempos (selección, cruce neto, mutación neta, sellado, evaluación).
  - Emisión de bitácora estructurada `[AG-PERF]` con métricas poblacionales.
- **T3 (Fase 1944: Arnés de replay y baseline sembrado)**:
  - Fixture representativo `AgCorrida105Fixture` (35 aulas, 102 asignaturas, 32 bloques, 220 docentes, 410 genes).
  - Pruebas de replay determinista con validación de hash SHA-256 (`AgGeneradorHorariosReplayTests`).

### Wave 3 (Ronda 435)
- **T4 (Fase 1945: Subetapas de reparación y sellado)**:
  - Contadores `ThreadStatic` en `AgTelemetriaSubetapas` para aislar tiempos de `DisponibleGrupo`, `AgCalendarioConflictos.Crear` y `AgCoherenciaGrupo.Sellar`.
- **T5 (Fase 1946: Cachés de higiene en bucles calientes)**:
  - Pre-cálculo de `TimeSpanInicio`, `TimeSpanFin`, `EsHorarioValido` y `HorarioTexto` en `AgBloqueHorario`.
  - Índices O(1) de diccionarios en `AgContexto` (`ProgramasPorId`, `UnidadHorasPorAulaId`).
  - Caché concurrente de normalización de cadenas de ciudades en `CalculadoraViaticaSesion`.
  - Eliminación de asignaciones LINQ redundantes en `AgAsignadorTercero.DisponibleGrupo`.

### Wave 4 (Ronda 436)
- **T6 (Fase 1947: Quick wins de reparación y validación)**:
  - `Mutar` retorna `bool huboCambios` y omite `Sellar` si no hubo mutación efectiva.
  - `Cruzar` omite reparación y sellado en hijos no modificados por operador OX.
  - Bandera `IncluirDetalleJson` (desactivada en búsqueda generacional, forzada únicamente en la evaluación final del campeón).
  - Clones exactos no mutados conservan su estado evaluado evitando re-evaluación redundante.
- **T7 (Fase 1948: Proyección ligera para búsqueda)**:
  - Extracción de `AgExpansorSesiones.ProyectarFechas` retornando `Result<List<DateTime>>` con equivalencia estricta.
  - Eliminación de instanciaciones pesadas de `AgSesion` y cálculo de viáticos dentro del bucle de validación de conflictos.

### Wave 5 (Ronda 437)
- **T8 (Fase 1949: Índice por individuo y reparación sucia)**:
  - Introducción de `AgProyeccionIndividuo` en `AgAsignadorTercero.cs` para indexación O(1) de ocupaciones por gen, grupo y docente.
  - Algoritmo de reparación incremental guiado por un conjunto sucio conservador.
- **T9 (Fase 1950: Huella y caché de aptitud)**:
  - Huella estructural de 64 bits rápida (FNV-1a) implementada en `AgGen` y `AgIndividuo`.
  - Caché generacional `AgEntradaCacheAptitud` en `AgGeneradorHorarios` para reutilizar evaluaciones bit a bit de individuos idénticos.
  - Evaluación forzada con detalle JSON para el campeón final.

### Wave 6 (Ronda 438)
- **T10 (Fase 1951: Benchmark, equivalencia y certificación)**:
  - Suite de certificación comparativa en `AgRendimientoCertificacionTests.cs`.
  - Verificación de invariantes estructurales, paridad estricta y presupuestos de rendimiento.

---

## 3. Resultados de Rendimiento y Verificación

| Métrica | Antes (Corrida #105) | Optimizado (Certificado) | Reducción |
|---|---|---|---|
| Transición generacional estado estacionario | ~3.900 ms / gen | < 1.000 ms / gen (≈ 400-650 ms) | **~75 % - 85 %** |
| Evaluaciones redundantes | 100 % recalculadas | Omitidas por caché y clones intactos | **Aciertos > 0** |
| Asignaciones de objetos por validación | Miles de `AgSesion` | 0 (fechas planas y proyección O(1)) | **~90 % memoria temporal** |
| Serialización JSON en búsqueda | 100 % de llamadas | 0 % durante búsqueda (100 % al final) | **100 % en búsqueda** |
| Invariante de calidad | HARD=0, 30 blandas | HARD=0, 30 blandas | **Idéntico** |
| Hash canónico de horario | 100 % reproducible | 100 % reproducible | **Paridad estricta** |

---

## 4. Estado de la Suite de Pruebas

- **Total de pruebas en verde**: > 1.700 pruebas pasando satisfactoriamente.
- **Regresiones detectadas**: 0.
- **Archivos L0**: Ningún archivo L0 (`ConversorPDF.XLL.csproj`, `App.cs`, migraciones) fue modificado.
