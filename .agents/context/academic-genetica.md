# Context Pack: Academic / Genetica

- Generado: 2026-10-27 11:36
- Cabecera obligatoria: ver `references/plan-contract.md` (Execution envelope).
- Presupuesto: <= 400 lineas.

## 1. Inventario del slice
- Archivos de produccion: 63
- Archivos de test: 27
- Ultima migracion detectada: V78

## 2. Contratos
- Sin declaraciones `public interface` dentro del slice.

Interfaces de `Academic/Domain/Interfaces` referenciadas por el slice:
- IAgEvaluadorAptitud (Academic/Domain/Interfaces/IAgEvaluadorAptitud.cs)
- IAgExpansorSesiones (Academic/Domain/Interfaces/IAgExpansorSesiones.cs)
- IAgGeneradorHorarios (Academic/Domain/Interfaces/IAgGeneradorHorarios.cs)
- IAgMaterializadorSesiones (Academic/Domain/Interfaces/IAgMaterializadorSesiones.cs)
- IAgPromotorVinculaciones (Academic/Domain/Interfaces/IAgPromotorVinculaciones.cs)
- IAgResolutorVentanaSemestre (Academic/Domain/Interfaces/IAgResolutorVentanaSemestre.cs)
- IAgSalvedadRepository (Academic/Domain/Interfaces/IAgSalvedadRepository.cs)
- IAgScheduleRepository (Academic/Domain/Interfaces/IAgScheduleRepository.cs)
- IAgValidadorReglas (Academic/Domain/Interfaces/IAgValidadorReglas.cs)
- IAgValidadorVentana (Academic/Domain/Interfaces/IAgValidadorVentana.cs)

## 3. Persistencia
- Archivos con SQLite/Dapper/EncryptedSQLiteContext: 3
  - AgEsquemaDiagnostico.cs
  - AgPromotorVinculaciones.cs
  - AgResolutorVentanaSemestre.cs
- Version de migracion vigente (maxima en disco): V78

## 4. UI y entry points
- Formularios/UI en el slice: 0
- Referenciados desde RibbonController: 0

## 5. Gotchas detectadas (conteo de patrones prohibidos)
- Interop Office en el slice: 0 (debe ser 0 fuera de XLL/Infrastructure)
- WinForms/DataGridView en el slice: 0 (0 esperado en Application/Domain)
- Sintaxis C# 12 (`[]`): 22 (debe ser 0 en net48)
- DateTime.Now: 8 (usar IDateTimeProvider)
- .Result/.Wait(): 0 (prohibido salvo COM STA justificado)

## 6. Archivos
### Produccion (63)
- ConversorPDF.XLL\Academic\Application\Services\Genetica\AgAsignadorTercero.cs
- ConversorPDF.XLL\Academic\Application\Services\Genetica\AgAulaFactory.cs
- ConversorPDF.XLL\Academic\Application\Services\Genetica\AgBloqueSolape.cs
- ConversorPDF.XLL\Academic\Application\Services\Genetica\AgCoherenciaGrupo.cs
- ConversorPDF.XLL\Academic\Application\Services\Genetica\AgDiagnosticoTraza.cs
- ConversorPDF.XLL\Academic\Application\Services\Genetica\AgEncabezadoVinculacionFactory.cs
- ConversorPDF.XLL\Academic\Application\Services\Genetica\AgEsquemaDiagnostico.cs
- ConversorPDF.XLL\Academic\Application\Services\Genetica\AgEvaluadorAptitud.cs
- ConversorPDF.XLL\Academic\Application\Services\Genetica\AgExpansorSesiones.cs
- ConversorPDF.XLL\Academic\Application\Services\Genetica\AgFabricaIndividuos.cs
- ConversorPDF.XLL\Academic\Application\Services\Genetica\AgFormatoAsignatura.cs
- ConversorPDF.XLL\Academic\Application\Services\Genetica\AgGeneradorHorarios.cs
- ConversorPDF.XLL\Academic\Application\Services\Genetica\AgMaterializadorSesiones.cs
- ConversorPDF.XLL\Academic\Application\Services\Genetica\AgMensajeDimensionesAula.cs
- ConversorPDF.XLL\Academic\Application\Services\Genetica\AgMetricas.cs
- ConversorPDF.XLL\Academic\Application\Services\Genetica\AgOperadorCruce.cs
- ConversorPDF.XLL\Academic\Application\Services\Genetica\AgOperadorMutacion.cs
- ConversorPDF.XLL\Academic\Application\Services\Genetica\AgOperadorSeleccion.cs
- ConversorPDF.XLL\Academic\Application\Services\Genetica\AgPatronSemanal.cs
- ConversorPDF.XLL\Academic\Application\Services\Genetica\AgPlanificadorOleadas.cs
- ConversorPDF.XLL\Academic\Application\Services\Genetica\AgPlanificadorPatronSemanal.cs
- ConversorPDF.XLL\Academic\Application\Services\Genetica\AgPlanificadorTemporal.cs
- ConversorPDF.XLL\Academic\Application\Services\Genetica\AgPromotorVinculaciones.cs
- ConversorPDF.XLL\Academic\Application\Services\Genetica\AgResolutorVentanaSemestre.cs
- ConversorPDF.XLL\Academic\Application\Services\Genetica\AgSeleccionAulaExpandidor.cs
- ConversorPDF.XLL\Academic\Application\Services\Genetica\AgSeleccionBloque.cs
- ConversorPDF.XLL\Academic\Application\Services\Genetica\AgSemanaIso.cs
- ConversorPDF.XLL\Academic\Application\Services\Genetica\AgSincronizadorCorrecciones.cs
- ConversorPDF.XLL\Academic\Application\Services\Genetica\AgTreeViewCheckService.cs
- ConversorPDF.XLL\Academic\Application\Services\Genetica\AgUnidadesSemanales.cs
- ConversorPDF.XLL\Academic\Application\Services\Genetica\AgValidadorReglas.cs
- ConversorPDF.XLL\Academic\Application\Services\Genetica\AgValidadorVentana.cs
- ConversorPDF.XLL\Academic\Domain\Models\Genetica\AgArbol.cs
- ConversorPDF.XLL\Academic\Domain\Models\Genetica\AgArbolNodo.cs
- ConversorPDF.XLL\Academic\Domain\Models\Genetica\AgAula.cs
- ConversorPDF.XLL\Academic\Domain\Models\Genetica\AgAulaAsignatura.cs
- ConversorPDF.XLL\Academic\Domain\Models\Genetica\AgBandasDias.cs
- ConversorPDF.XLL\Academic\Domain\Models\Genetica\AgBitacora.cs
- ConversorPDF.XLL\Academic\Domain\Models\Genetica\AgBloqueHorario.cs
- ConversorPDF.XLL\Academic\Domain\Models\Genetica\AgCarriles.cs
- ConversorPDF.XLL\Academic\Domain\Models\Genetica\AgClaveCohorte.cs
- ConversorPDF.XLL\Academic\Domain\Models\Genetica\AgContexto.cs
- ConversorPDF.XLL\Academic\Domain\Models\Genetica\AgCorrida.cs
- ConversorPDF.XLL\Academic\Domain\Models\Genetica\AgDiasVentana.cs
- ConversorPDF.XLL\Academic\Domain\Models\Genetica\AgEscaleraVentanas.cs
- ConversorPDF.XLL\Academic\Domain\Models\Genetica\AgFranjaJornada.cs
- ConversorPDF.XLL\Academic\Domain\Models\Genetica\AgGen.cs
- ConversorPDF.XLL\Academic\Domain\Models\Genetica\AgIndividuo.cs
- ConversorPDF.XLL\Academic\Domain\Models\Genetica\AgPesosFitness.cs
- ConversorPDF.XLL\Academic\Domain\Models\Genetica\AgPlanOleada.cs
- ConversorPDF.XLL\Academic\Domain\Models\Genetica\AgPlanOleadasResultado.cs
- ConversorPDF.XLL\Academic\Domain\Models\Genetica\AgPreferenciaDocente.cs
- ConversorPDF.XLL\Academic\Domain\Models\Genetica\AgPromocion.cs
- ConversorPDF.XLL\Academic\Domain\Models\Genetica\AgRanura.cs
- ConversorPDF.XLL\Academic\Domain\Models\Genetica\AgRanuras.cs
- ConversorPDF.XLL\Academic\Domain\Models\Genetica\AgResultadoCorrida.cs
- ConversorPDF.XLL\Academic\Domain\Models\Genetica\AgSalvedad.cs
- ConversorPDF.XLL\Academic\Domain\Models\Genetica\AgSesion.cs
- ConversorPDF.XLL\Academic\Domain\Models\Genetica\AgVinculacion.cs
- ConversorPDF.XLL\Academic\Domain\Models\Genetica\AgViolacion.cs
- ConversorPDF.XLL\Academic\Domain\Models\Genetica\FranjaJornada.cs
- ConversorPDF.XLL\Academic\Domain\Models\Genetica\InformeFactibilidad.cs
- ConversorPDF.XLL\Academic\Domain\Models\Genetica\VentanaSemestre.cs

### Tests (27)
- ConversorPDF.Tests\Academic\Application\Services\Genetica\AgAsignadorTerceroTests.cs
- ConversorPDF.Tests\Academic\Application\Services\Genetica\AgAulaFactoryTests.cs
- ConversorPDF.Tests\Academic\Application\Services\Genetica\AgCoherenciaGrupoTests.cs
- ConversorPDF.Tests\Academic\Application\Services\Genetica\AgDiagnosticoTrazaTests.cs
- ConversorPDF.Tests\Academic\Application\Services\Genetica\AgEsquemaDiagnosticoTests.cs
- ConversorPDF.Tests\Academic\Application\Services\Genetica\AgEvaluadorAptitudTests.cs
- ConversorPDF.Tests\Academic\Application\Services\Genetica\AgExpansorSesionesTests.cs
- ConversorPDF.Tests\Academic\Application\Services\Genetica\AgFabricaIndividuosTests.cs
- ConversorPDF.Tests\Academic\Application\Services\Genetica\AgFormatoAsignaturaTests.cs
- ConversorPDF.Tests\Academic\Application\Services\Genetica\AgGeneradorHorariosTests.cs
- ConversorPDF.Tests\Academic\Application\Services\Genetica\AgMaterializadorSesionesTests.cs
- ConversorPDF.Tests\Academic\Application\Services\Genetica\AgMensajeDimensionesAulaTests.cs
- ConversorPDF.Tests\Academic\Application\Services\Genetica\AgMetricasTests.cs
- ConversorPDF.Tests\Academic\Application\Services\Genetica\AgOperadorCruceTests.cs
- ConversorPDF.Tests\Academic\Application\Services\Genetica\AgOperadorMutacionCruceFrozenTests.cs
- ConversorPDF.Tests\Academic\Application\Services\Genetica\AgOperadorMutacionTests.cs
- ConversorPDF.Tests\Academic\Application\Services\Genetica\AgOperadorSeleccionTests.cs
- ConversorPDF.Tests\Academic\Application\Services\Genetica\AgPlanificadorOleadasTests.cs
- ConversorPDF.Tests\Academic\Application\Services\Genetica\AgPlanificadorPatronSemanalTests.cs
- ConversorPDF.Tests\Academic\Application\Services\Genetica\AgPlanificadorTemporalTests.cs
- ConversorPDF.Tests\Academic\Application\Services\Genetica\AgPromotorVinculacionesTests.cs
- ConversorPDF.Tests\Academic\Application\Services\Genetica\AgResolutorVentanaSemestreTests.cs
- ConversorPDF.Tests\Academic\Application\Services\Genetica\AgSeleccionAulaExpandidorTests.cs
- ConversorPDF.Tests\Academic\Application\Services\Genetica\AgTreeViewCheckServiceTests.cs
- ConversorPDF.Tests\Academic\Application\Services\Genetica\AgValidadorReglasTests.cs
- ConversorPDF.Tests\Academic\Application\Services\Genetica\AgValidadorVentanaTests.cs
- ConversorPDF.Tests\Academic\Application\Services\Genetica\PoolDocenteUnionTests.cs
