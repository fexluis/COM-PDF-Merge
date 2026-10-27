using System;
using System.Data;
using System.Linq;
using ConversorPDF.XLL.Core.Helpers;
using Dapper;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    /// <summary>
    /// Migración V76: Sanea dimensiones de decisión (jornada, horas_dia, dias_ventana) en ag_aula
    /// a partir de la configuración guardada en ag_arbol_nodo (Fase 1191 / Fase 1).
    /// </summary>
    public sealed class MigrationV76_SanearDimensionesAgAula : IMigration, IMigrationTolerante
    {
        public int Version => 76;
        public string Name => "SanearDimensionesAgAula";
        public bool Aditiva => true;

        public void Up(IDbConnection connection, IDbTransaction transaction)
        {
            try
            {
                int existeAula = connection.QueryFirstOrDefault<int>(
                    "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='ag_aula';",
                    transaction: transaction);
                int existeArbol = connection.QueryFirstOrDefault<int>(
                    "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='ag_arbol_nodo';",
                    transaction: transaction);

                if (existeAula == 0 || existeArbol == 0)
                {
                    return;
                }

                // Paso A: Saneamiento fiel a partir de ag_arbol_nodo (sin inventar, Fases 1196, 1197 y 1213)
                const string sqlCandidatos = @"
SELECT a.id AS Id, a.aula_ref AS AulaRef,
       a.jornada AS JornadaActual, a.horas_dia AS HorasDiaActual, a.dias_ventana AS DiasVentanaActual,
       n.jornada AS JornadaNodo, n.horas_dia AS HorasDiaNodo, n.dias_ventana AS DiasVentanaNodo
FROM ag_aula a
JOIN ag_arbol_nodo n
  ON n.idarbol = (
       SELECT MAX(x.idarbol) FROM ag_arbol_nodo x
       WHERE x.aula_ref = a.aula_ref
         AND (COALESCE(x.subunidad,'') = COALESCE(a.subunidad,'') OR TRIM(COALESCE(a.subunidad,'')) = '')
         AND TRIM(COALESCE(x.dias_ventana,'')) <> ''
         AND TRIM(COALESCE(x.jornada,'')) <> ''
         AND x.horas_dia > 0)
 AND n.aula_ref = a.aula_ref
WHERE (COALESCE(a.dias_ventana,'')='' OR COALESCE(a.jornada,'')='' OR a.horas_dia = 0);";

                var candidatos = connection.Query<AgAulaDimensionesCandidato>(
                    sqlCandidatos, transaction: transaction).ToList();

                int saneadas = 0;
                const string sqlUpdate = @"
UPDATE ag_aula 
SET jornada = @J, horas_dia = @H, dias_ventana = @D 
WHERE id = @Id;";

                foreach (var c in candidatos)
                {
                    string nuevaJornada = string.IsNullOrWhiteSpace(c.JornadaActual) ? c.JornadaNodo : c.JornadaActual;
                    int nuevasHoras = c.HorasDiaActual > 0 ? c.HorasDiaActual : c.HorasDiaNodo;
                    string nuevosDias = string.IsNullOrWhiteSpace(c.DiasVentanaActual) ? c.DiasVentanaNodo : c.DiasVentanaActual;

                    int afect = connection.Execute(sqlUpdate, new
                    {
                        J = nuevaJornada,
                        H = nuevasHoras,
                        D = nuevosDias,
                        Id = c.Id
                    }, transaction: transaction);

                    if (afect > 0)
                    {
                        saneadas += afect;
                    }
                }

                if (saneadas > 0)
                {
                    LoggerHelper.Info($"[V76] Aulas saneadas con configuración de árbol: {saneadas}");
                }

                // Paso B: Fallback canónico (desactivado por defecto: política segura de no inventar)
                const bool AplicarFallbackCanonico = false;
                if (AplicarFallbackCanonico)
                {
                    const string sqlFallback = @"
UPDATE ag_aula 
SET jornada = CASE WHEN COALESCE(jornada,'')='' THEN 'MANANA' ELSE jornada END,
    horas_dia = CASE WHEN horas_dia <= 0 THEN 4 ELSE horas_dia END,
    dias_ventana = CASE WHEN COALESCE(dias_ventana,'')='' THEN 'MAR-VIE' ELSE dias_ventana END
WHERE (COALESCE(dias_ventana,'')='' OR COALESCE(jornada,'')='' OR horas_dia = 0);";

                    int fb = connection.Execute(sqlFallback, transaction: transaction);
                    if (fb > 0)
                    {
                        LoggerHelper.Info($"[V76] Fallback canónico aplicado a aulas sin nodo: {fb}");
                    }
                }

                // Paso C: Trazabilidad diagnóstica de aulas que permanecen sin dimensiones (Fase 1216: acotado a 20)
                const string sqlFaltantes = @"
SELECT aula_ref 
FROM ag_aula 
WHERE (COALESCE(dias_ventana,'')='' OR COALESCE(jornada,'')='' OR horas_dia = 0);";

                var faltantes = connection.Query<string>(sqlFaltantes, transaction: transaction).ToList();
                const int MaxDetalleV76 = 20;
                foreach (var aulaRef in faltantes.Take(MaxDetalleV76))
                {
                    LoggerHelper.Warning($"[V76] AulaSinDimensiones aula_ref={aulaRef} (sin nodo con configuración; ejecute 'Generar Aulas')");
                }
                if (faltantes.Count > MaxDetalleV76)
                {
                    LoggerHelper.Warning($"[V76] {faltantes.Count - MaxDetalleV76} aula(s) adicionales sin dimensiones (detalle truncado a {MaxDetalleV76}).");
                }
                if (faltantes.Count > 0)
                {
                    LoggerHelper.Info($"[V76] Total aulas sin dimensiones: {faltantes.Count}");
                }
            }
            catch (Exception ex)
            {
                // Fase 1200: Si falla, registrar error y relanzar para que SchemaMigrator (Aditiva=true)
                // registre un hueco de versión y reintente el saneamiento en el próximo arranque.
                LoggerHelper.Error($"[V76] Saneamiento de dimensiones en ag_aula falló; se reintentará en el próximo arranque. {ex.Message}", ex);
                throw;
            }
        }
    }
}
