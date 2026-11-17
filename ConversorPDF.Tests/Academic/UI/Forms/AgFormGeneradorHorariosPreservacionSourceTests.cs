using System;
using System.IO;
using FluentAssertions;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ConversorPDF.Tests.Academic.UI.Forms
{
    // Contratos estructurales complementarios: no sustituyen la integración UI/SQLite pendiente.
    [TestClass]
    public class AgFormGeneradorHorariosPreservacionSourceTests
    {
        [TestMethod]
        public void Inicio_NoPurgaHistorico_YAdquiereIdentidadAntesDelPrimerAwait()
        {
            var inicio = Inicio();
            inicio.Should().NotContain("LimpiarStagingAsync(");
            inicio.Should().Contain("_diagnosticoEnCurso || _intentoOptimizacion.HasValue");
            Antes(inicio, "_intentoOptimizacion = intento;", "await _semaforoSyncDimensiones.WaitAsync(");
            inicio.Should().Contain("if (semaforoAdquirido) _semaforoSyncDimensiones.Release();");
            inicio.Should().Contain("treeContextoAcademico.Enabled = arbolHabilitado;");
            inicio.Should().Contain("btnEjecutarAG.Enabled = btnEjecutarHabilitado;");
            inicio.Should().Contain("btnCancelar.Enabled = false;");
            inicio.Should().Contain("_ctsEjecucion.Dispose();");
            Antes(inicio, "await RefrescarComboCorridasAsync();", "_intentoOptimizacion = null;");
        }

        [TestMethod]
        public void Failure_CierraVisualYPersistenciaAntesDelModalYRetorno()
        {
            var inicio = Inicio();
            var failure = Entre(inicio, "if (optResult.IsFailure)", "_ultimoResultado = optResult.Value;");
            Antes(failure, "CerrarEstadoVisualOptimizacion(cancelada);", "MessageBox.Show(");
            Antes(failure, "await CerrarCorridaFallidaAsync(corrida, cancelada);", "MessageBox.Show(");
            Antes(failure, "MessageBox.Show(", "return;");
            inicio.Should().Contain("if (!PuedeRecibirProgresoOptimizacion(intento)) return;");
            Antes(inicio, "_aceptaProgresoOptimizacion = false;", "if (optResult.IsFailure)");
        }

        [TestMethod]
        public void Intento_EnviaUnInicioYUnFinAlLogger_YSnapshotNoTieneFallbackExitoso()
        {
            var inicio = Inicio();
            Antes(inicio, "_intentoOptimizacion = intento;", "RegistrarEventoIntento(\"INICIO\", intento, 0);");
            Antes(inicio, "RegistrarEventoIntento(\"INICIO\", intento, 0);", "await _semaforoSyncDimensiones.WaitAsync(");
            inicio.Split(new[] { "RegistrarEventoIntento(\"INICIO\"" }, StringSplitOptions.None).Length.Should().Be(2);
            inicio.Split(new[] { "RegistrarEventoIntento(\"FIN\"" }, StringSplitOptions.None).Length.Should().Be(2);
            var cierre = Entre(inicio, "finally\n            {", "_intentoOptimizacion = null;");
            cierre.Should().Contain("RegistrarEventoIntento(\"FIN\", intento, idCorridaCreada);");
            Antes(inicio, "{ \"intentoId\", intento.ToString(\"D\") }", "InsertCorridaAsync(corrida,");
            inicio.Should().NotContain("parametrosJson = \"{\\\"intentoId");
            inicio.Should().Contain("throw new InvalidOperationException(\"SNAPSHOT_SERIALIZACION_ERROR\", ex)");
            inicio.Should().Contain("{ \"manifiestoCobertura\", _contextoActual.ManifiestoCobertura }");
            inicio.Should().Contain("{ \"planCanonico\", _contextoActual.PlanAprobadoSnapshot }");
            inicio.Should().Contain("{ \"validacionPlan\", _contextoActual.ValidacionPlanAprobado }");
            inicio.Should().Contain("ParametrosJson = parametrosJson");
            var logger = Entre(Fuente(), "private void RegistrarEventoIntento(", "private async Task CerrarCorridaFallidaAsync(");
            logger.Should().Contain("LoggerHelper.Info(");
            logger.Should().Contain("intentoId={intento:D} corridaId={idCorrida}");
            logger.Should().Contain("fase={_faseFinOptimizacion");
            logger.Should().Contain("motivo={_motivoFinOptimizacion");
            logger.Should().NotContain("ParametrosJson").And.NotContain("ex.Message").And.NotContain("Environment.UserName");
        }

        [TestMethod]
        public void Snapshot_CatchYFinConservanMotivo_YNuevoIntentoDescartaTerminalAnterior()
        {
            var inicio = Inicio();
            Antes(inicio, "_terminalVisualOptimizacion = false;", "await _semaforoSyncDimensiones.WaitAsync(");
            Antes(inicio, "_faseFinOptimizacion = null;", "await _semaforoSyncDimensiones.WaitAsync(");
            Antes(inicio, "_motivoFinOptimizacion = null;", "await _semaforoSyncDimensiones.WaitAsync(");
            var serializacion = Entre(inicio, "MarcarErrorPreparacion(\"SNAPSHOT_SERIALIZACION_ERROR\")", "corrida = new AgCorrida");
            serializacion.Should().Contain("throw new InvalidOperationException(\"SNAPSHOT_SERIALIZACION_ERROR\", ex)");
            var catches = Entre(inicio, "catch (OperationCanceledException)", "finally\n            {");
            catches.Should().Contain("CerrarEstadoVisualOptimizacion(true);");
            catches.Should().Contain("CerrarEstadoVisualOptimizacion(_ctsEjecucion.IsCancellationRequested);");
            catches.Should().NotContain("EstablecerEstadoOptimizacion(");
            var cierre = Entre(Fuente(), "private void CerrarEstadoVisualOptimizacion(bool cancelada)", "// Los helpers compartidos");
            Antes(cierre, "if (_terminalVisualOptimizacion) return;", "EstablecerEstadoOptimizacion(");
            inicio.Should().Contain("if (!_terminalVisualOptimizacion) CerrarEstadoVisualOptimizacion(cancelada);");
            Antes(inicio, "CerrarEstadoVisualOptimizacion(_ctsEjecucion.IsCancellationRequested);", "RegistrarEventoIntento(\"FIN\", intento, idCorridaCreada);");
        }

        [DataTestMethod]
        [DataRow("P08_SincronizarDimensiones", "MessageBox.Show(this, resSync.Error")]
        [DataRow("P08_ValidarContexto", "MessageBox.Show(this, ctxRes.Error")]
        [DataRow("P08_ValidarAulas", "MessageBox.Show(msg,")]
        [DataRow("P08_ValidarAnclajeAula", "MessageBox.Show(this, err,")]
        [DataRow("P08_ValidarAsignaturas", "MessageBox.Show(msg,")]
        [DataRow("P08_ValidarAulaAsignaturas", "MessageBox.Show(this, msg,")]
        [DataRow("P08_HorasDiaNivel", "MessageBox.Show(this, errD4,")]
        [DataRow("P08_ValidarVentanaBloques", "MessageBox.Show(this, msgModal,")]
        public void Preflight_ErrorEsTerminalAntesDelModalYRetorno(string codigo, string modal)
        {
            var rama = Inicio();
            int desde = rama.IndexOf("MarcarErrorPreparacion(\"" + codigo + "\")", StringComparison.Ordinal);
            desde.Should().BeGreaterOrEqualTo(0);
            rama = rama.Substring(desde);
            Antes(rama, "MarcarErrorPreparacion(", modal);
            Antes(rama, modal, "return;");
        }

        [DataTestMethod]
        [DataRow("ESCALADO_EVALUACION_")]
        [DataRow("ESCALADO_SELECCION_")]
        [DataRow("ESCALADO_SIMULACION_")]
        [DataRow("ESCALADO_PERFIL_OBSOLETO")]
        [DataRow("ESCALADO_APLICACION")]
        [DataRow("ESCALADO_RECARGA_CONTEXTO")]
        [DataRow("ESCALADO_RECARGA_DIMENSIONES")]
        [DataRow("ESCALADO_ARGUMENTO_INVALIDO")]
        public void Escalado_ErrorEsTerminalAntesDeSuModal(string codigo)
        {
            var escalado = Entre(Fuente(), "private async Task<AgResultadoEscaladoUI> EvaluarYAplicarEscaladoInternoAsync(",
                "private async Task IniciarOptimizacionAsync()");
            int desde = escalado.IndexOf("MarcarErrorPreparacion(\"" + codigo, StringComparison.Ordinal);
            desde.Should().BeGreaterOrEqualTo(0);
            var rama = escalado.Substring(desde);
            Antes(rama, "MarcarErrorPreparacion(", "MessageBox.Show(");
            Antes(rama, "MessageBox.Show(", "return AgResultadoEscaladoUI.FalloAlAplicar;");
        }

        [TestMethod]
        public void Preparacion_PreguntasNoSonTerminales_YFinallyRespetaTerminalSignificativo()
        {
            var inicio = Inicio();
            Antes(inicio, "EstadoVisualOptimizacion.Preparando", "await _semaforoSyncDimensiones.WaitAsync(");
            Antes(inicio, "EstadoVisualOptimizacion.Validando", "await ConstruirContextoAsync(");
            Antes(inicio, "EstadoVisualOptimizacion.Optimizando", "await _generadorHorarios.EjecutarOptimizacionAsync(");
            Antes(inicio, "MarcarDecisionPreparacion();", "var cont = MessageBox.Show(");
            Antes(inicio, "EstablecerEstadoOptimizacion(EstadoVisualOptimizacion.Cancelado, \"ESCALADO_CONTINUACION_RECHAZADA\")", "// Fase 295, 296, 303 & 321");
            inicio.Should().Contain("if (!_terminalVisualOptimizacion) CerrarEstadoVisualOptimizacion(cancelada);");
            var escalado = Entre(Fuente(), "private async Task<AgResultadoEscaladoUI> EvaluarYAplicarEscaladoInternoAsync(",
                "private async Task IniciarOptimizacionAsync()");
            Antes(escalado, "MarcarDecisionPreparacion();", "dlg.ShowDialog(this)");
            Antes(escalado, "MarcarDecisionPreparacion();", "var respuesta = MessageBox.Show(");
            var filtro = Entre(Fuente(), "traza.Advertencia(\"P06_FiltroTercerosVacio\"", "// P06: Construir índices");
            Antes(filtro, "MarcarErrorPreparacion(\"P06_FiltroTercerosVacio\");", "MessageBox.Show(");
        }

        [TestMethod]
        public void AmbasRutas_PreparanOrigenAntesDeSincronizarOEscalar_YDisponibleConservaSuRuta()
        {
            var fuente = Fuente();
            var diagnostico = Entre(fuente, "private async Task EjecutarDiagnosticoPreVueloAsync()", "private enum AgResultadoEscaladoUI");
            foreach (var ruta in new[] { diagnostico, Inicio() })
            {
                Antes(ruta, "await PrepararCoberturaAntesDeCambiosAsync(", "SincronizarConfirmadoAsync(");
                Antes(ruta, "await PrepararCoberturaAntesDeCambiosAsync(", "await EvaluarYAplicarEscaladoAsync(");
                ruta.Should().Contain("if (!await PrepararCoberturaAntesDeCambiosAsync(");
                ruta.Should().Contain("await ConstruirContextoAsync(");
                ruta.Should().Contain("await EvaluarYAplicarEscaladoAsync("); // Disponible: se mantiene escalado heredado.
            }
            var preparacion = Entre(fuente, "private async Task<bool> PrepararCoberturaAntesDeCambiosAsync(", "private enum AgResultadoEscaladoUI");
            preparacion.Should().Contain("ObtenerCombinacionesSeleccionadas()");
            preparacion.Should().Contain("StringComparison.OrdinalIgnoreCase");
            preparacion.Should().NotContain("ExtraerCohorte(").And.NotContain("FirstOrDefault(");
            Antes(preparacion, "if (!manifiesto.Value.Relaciones.Any(", "await _planAprobadoRepository.ObtenerAsync(");
            preparacion.Should().Contain("TipoPreparacionPlan.Cancelada");
            preparacion.Should().Contain("TipoPreparacionPlan.ErrorTecnico");
            preparacion.Should().Contain("TipoPreparacionPlan.NoVerificable");
            preparacion.Should().Contain("ConfirmarAsync(contexto, propuesta, aula.Id,");
            preparacion.Should().Contain("PUBLICACION_NO_DISPONIBLE");
            fuente.Should().Contain("_planAprobadoRepository = planAprobadoRepository ?? new AgPlanAprobadoRepository(_connectionFactory)");
        }

        [TestMethod]
        public void Aprobadas_RatificacionEsExplicita_CandidatoNoAbreGeneracion_YDisponibleSigue()
        {
            var preparacion = Entre(Fuente(), "private async Task<bool> PrepararCoberturaAntesDeCambiosAsync(",
                "private bool MostrarFalloPlan(");
            Antes(preparacion, "if (!manifiesto.Value.Relaciones.Any(", "await _planAprobadoRepository.ObtenerAsync(");
            Antes(preparacion, "new AgSimuladorEscaladoAprobadas().Evaluar(contexto, ct)", "new AgFormConfirmacionPlanAprobado(");
            Antes(preparacion, "CapturarConfiguracionAsync(contexto, aula.Id, propuesta, ct)", "formulario.ShowDialog(this)");
            Antes(preparacion, "formulario.ShowDialog(this)", "ConfirmarAsync(contexto, propuesta, aula.Id,");
            preparacion.Should().Contain("if (confirmada.Codigo != \"CANDIDATO_PUBLICADO\"");
            preparacion.Should().Contain("_aulasConfirmadasEnIntento.Contains(aula.Id)");
            preparacion.Should().Contain("I5_COBERTURA_INSUFICIENTE");
            preparacion.Should().Contain("RegistrarCandidatoCobertura(confirmada.Candidato)");
            preparacion.Should().Contain("new AgProyectorSolucionCobertura().Conservar(contexto)");
            preparacion.Should().NotContain("AplicarDetalladoAsync(").And.NotContain("Autorizado = true");
            preparacion.Should().Contain("Días añadidos:").And.Contain("Sesiones canónicas anteriores:")
                .And.Contain("Simulación actual:").And.Contain("Patrón afectado");
        }

        [TestMethod]
        public void T4_NoConvierteAprobadaNiSaltaT5_SinPublicacionNoHayExito()
        {
            var fuente = Fuente();
            var preparacion = Entre(fuente, "private async Task<bool> PrepararCoberturaAntesDeCambiosAsync(",
                "private bool MostrarFalloPlan(");
            preparacion.Should().Contain("formulario.ShowDialog(this)");
            preparacion.Should().Contain("BuscarArchivadoAsync(");
            preparacion.Should().Contain("_aulasConfirmadasEnIntento.Contains(aula.Id)");
            // Requisito 2026-11-17: sólo las aprobadas CON origen exigen publicación; una aprobada sin
            // origen productivo es demanda pendiente y no abre ni cierra la compuerta de ratificación.
            preparacion.Should().Contain("contexto.CandidatosCobertura.Count == conservables.Count");
            preparacion.Should().Contain("var conservables = aprobadas.Where(a => contexto.TieneOrigenConservable(a.Id)).ToList();");
            preparacion.Should().NotContain("UPDATE ag_aula").And.NotContain("AgEstadoAula.Disponible;");
            preparacion.Should().Contain("TipoPreparacionPlan.FlushFallido : TipoPreparacionPlan.Rechazado;");
            preparacion.Should().Contain("contexto.SesionesConservadas = conservadas.Value;");
        }

        [TestMethod]
        public void T4_TodasLasAulas_ReevaluaAislado_YTerminalesTipados_SinGeneracion()
        {
            var preparacion = Entre(Fuente(), "private async Task<bool> PrepararCoberturaAntesDeCambiosAsync(",
                "internal static string ConstruirDetallePropuestaAprobada(");
            // Requisito 2026-11-17: se recorren las aprobadas con origen; las pendientes se planifican
            // como demanda nueva y se registran sin pasar por ratificación.
            preparacion.Should().Contain("foreach (var aula in conservables)");
            preparacion.Should().Contain("var pendientesProgramacion = aprobadas.Where(a => !contexto.TieneOrigenConservable(a.Id)).ToList();");
            preparacion.Should().Contain("ReevaluarCandidatoConfirmado(contexto, propuesta,");
            preparacion.Should().Contain("ClonarContextoParaSimulacion(origen)");
            preparacion.Should().Contain("resultado.Value.Planes.Where(");
            preparacion.Should().Contain("TipoPreparacionPlan.CandidatoListo");
            preparacion.Should().Contain("TipoPreparacionPlan.CoberturaInsuficiente");
            preparacion.Should().Contain("TipoPreparacionPlan.FlushFallido");
            preparacion.Should().Contain("TipoPreparacionPlan.Rechazado");
            preparacion.Should().Contain("return MostrarFalloPlan(terminal, string.Join(");
            preparacion.Should().NotContain("return MostrarFalloPlan(TipoPreparacionPlan.RatificacionPendiente,");
            preparacion.Should().NotContain("AplicarDetalladoAsync(");
        }

        [TestMethod]
        public void T4_PublicacionCondicionada_ReservaBackup_RestauraYCompruebaVault()
        {
            var contexto = Archivo("ConversorPDF.XLL", "Academic", "Infrastructure", "Persistence", "Context", "EncryptedSQLiteContext.cs");
            var publisher = Archivo("ConversorPDF.XLL", "Academic", "Application", "Services", "Genetica", "AgPersistenciaRatificacion.cs");
            contexto.Should().Contain("CaptureVaultIdentity(vaultPath)");
            contexto.Should().Contain("expected.SameAs(current)");
            contexto.Should().Contain("ct.ThrowIfCancellationRequested();");
            Antes(contexto, "ct.ThrowIfCancellationRequested();\n                if (verified", "File.Replace(newEncPath, vaultPath,");
            contexto.Should().Contain("File.Exists(bakPath) ? null : bakPath");
            contexto.Should().NotContain("if (File.Exists(bakPath)) File.Delete(bakPath)");
            contexto.Should().Contain("File.Move(bakPath, vaultPath)");
            contexto.Should().Contain("published.Hash != hash");
            publisher.Should().Contain("CaptureVaultIdentity(contexto.EncryptedFilePath)");
            publisher.Should().Contain("UpdateVaultIfUnchanged(bytes, contexto.EncryptedFilePath, identidad, ct)");
            publisher.Should().Contain("PublicadoContiene(contexto, c)");
        }

        [TestMethod]
        public void T4_RatificacionCompletaYReplay_ExigenOrigenConsentimientoYPruebaPublicada()
        {
            var confirmador = Archivo("ConversorPDF.XLL", "Academic", "Application", "Services", "Genetica", "AgConfirmadorPlanAprobado.cs");
            var formulario = Archivo("ConversorPDF.XLL", "Academic", "UI", "Forms", "AgFormConfirmacionPlanAprobado.cs");
            confirmador.Should().Contain("leido.Value.HuellaOrigen != c.PlanAprobadoSnapshot.HuellaOrigen");
            confirmador.Should().NotContain("leido.Value.HuellaOrigen != p.HuellaOrigen");
            confirmador.Should().Contain("PublicadoContiene(_vault, restaurado)");
            confirmador.Should().Contain("PublicadoContiene(_vault, candidato)");
            confirmador.Should().Contain("Coincide(previas[0], candidato)");
            confirmador.Should().Contain("PermisoVigente(db, tx, actor)");
            formulario.Should().Contain("archivada.Consentimiento.Split(',')");
        }

        [TestMethod]
        public void T5_ValidaManifiestoDespuesDeStaging_YConservacionNoCreaPoblacion()
        {
            var inicio = Inicio();
            Antes(inicio, "InsertCorridaAsync(corrida,", "bool soloConservacion =");
            Antes(inicio, "bool soloConservacion =", "_generadorHorarios.EjecutarOptimizacionAsync(");
            inicio.Should().Contain("Array.Empty<AgSesionCobertura>(), _contextoActual.IdTerceroVacante");
            inicio.Should().Contain("new AgProyectorSolucionCobertura().Nuevas(_contextoActual, mejor, sesiones.Value)");
            inicio.Should().Contain("new AgValidadorCoberturaSeleccion().Validar(_contextoActual.ManifiestoCobertura,");
            inicio.Should().Contain("AgCompuertaPromocion.EsPromovible(corrida, mejor, sesiones, _contextoActual)");
            inicio.Should().Contain("{ \"decisionesCobertura\", _contextoActual.CandidatosCobertura }");
            inicio.Should().Contain("{ \"sesionesConservadas\", _contextoActual.SesionesConservadas }");
            Fuente().Should().Contain("Promoción bloqueada: falta reemplazo transaccional T6");
        }

        private static string Archivo(params string[] partes)
        {
            var dir = new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory);
            while (dir != null)
            {
                var ruta = Path.Combine(dir.FullName, Path.Combine(partes));
                if (File.Exists(ruta)) return File.ReadAllText(ruta).Replace("\r\n", "\n");
                dir = dir.Parent;
            }
            throw new FileNotFoundException(string.Join("/", partes));
        }

        private static void Antes(string texto, string primero, string segundo)
        {
            int a = texto.IndexOf(primero, StringComparison.Ordinal);
            int b = texto.IndexOf(segundo, StringComparison.Ordinal);
            a.Should().BeGreaterOrEqualTo(0, primero);
            b.Should().BeGreaterThan(a, segundo);
        }

        private static string Entre(string texto, string desde, string hasta)
        {
            int a = texto.IndexOf(desde, StringComparison.Ordinal);
            int b = texto.IndexOf(hasta, a, StringComparison.Ordinal);
            return texto.Substring(a, b - a);
        }

        private static string Inicio()
        {
            return Entre(Fuente(), "private async Task IniciarOptimizacionAsync()",
                "private bool PuedeRecibirProgresoOptimizacion(");
        }

        private static string Fuente()
        {
            var directorio = new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory);
            while (directorio != null)
            {
                var archivo = Path.Combine(directorio.FullName, "ConversorPDF.XLL", "Academic", "UI", "Forms", "AgFormGeneradorHorarios.cs");
                if (File.Exists(archivo))
                    return File.ReadAllText(archivo).Replace("\r\n", "\n");
                directorio = directorio.Parent;
            }
            throw new FileNotFoundException("Ejecutar el fixture estructural desde el árbol del repositorio.");
        }
    }
}
