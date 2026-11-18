using System;

namespace ConversorPDF.XLL.Academic.Application.Services.Genetica
{
    /// <summary>
    /// Provee sugerencias diagnósticas legibles a partir del código de error
    /// retornado por las validaciones de cobertura y manifiesto.
    /// </summary>
    internal static class AgSugerenciaCobertura
    {
        internal static string Texto(string error)
        {
            if (string.IsNullOrWhiteSpace(error))
            {
                return "Revise el diagnóstico de cobertura.";
            }

            if (error.StartsWith("CRUCE_AULA", StringComparison.Ordinal))
            {
                return "Conflicto de aula: dos sesiones simultáneas en la misma aula.";
            }

            if (error.StartsWith("CRUCE_DOCENTE", StringComparison.Ordinal))
            {
                return "Conflicto de docente: asignado a dos sesiones simultáneas en aulas distintas.";
            }

            if (error.StartsWith("COBERTURA_INCOMPLETA", StringComparison.Ordinal))
            {
                return "Demanda sin cubrir: revise pool docente y relaciones aula-asignatura.";
            }

            if (error.StartsWith("COBERTURA_EXCESIVA", StringComparison.Ordinal))
            {
                return "Sobreasignación: horas planificadas por encima de las requeridas.";
            }

            if (error.StartsWith("COBERTURA_ENTRADA_INVALIDA", StringComparison.Ordinal))
            {
                return "Entrada de cobertura inválida: manifiesto, conservadas o nuevas nulos.";
            }

            if (error.StartsWith("RELACION_", StringComparison.Ordinal))
            {
                return "Relación de cobertura no reconciliada con el manifiesto.";
            }

            if (error.StartsWith("ORIGEN_SESION_INVALIDO", StringComparison.Ordinal))
            {
                return "Sesión con origen inconsistente (CONSERVADA/NUEVA).";
            }

            if (error.StartsWith("SESION_DUPLICADA", StringComparison.Ordinal))
            {
                return "Sesión duplicada: misma relación, fecha y franja.";
            }

            if (error.StartsWith("SESION_FUERA_DE_VENTANA", StringComparison.Ordinal))
            {
                return "Sesión fuera de la ventana semestral.";
            }

            if (error.StartsWith("SESION_EN_FESTIVO", StringComparison.Ordinal))
            {
                return "Sesión en día festivo.";
            }

            if (error.StartsWith("SESION_NO_VERIFICABLE", StringComparison.Ordinal))
            {
                return "Sesión sin fecha u horario verificable.";
            }

            if (error.StartsWith("SESION_SIN_DOCENTE", StringComparison.Ordinal))
            {
                return "Sesión sin docente ni marca de vacante.";
            }

            if (error.StartsWith("MANIFIESTO_AUSENTE", StringComparison.Ordinal))
            {
                return "Manifiesto de cobertura ausente.";
            }

            if (error.StartsWith("PLAN_NO_VERIFICABLE", StringComparison.Ordinal))
            {
                return "Plan aprobado no verificable: falta snapshot, validación o ratificación.";
            }

            if (error.StartsWith("RATIFICACION_NO_ACREDITADA", StringComparison.Ordinal))
            {
                return "Ratificación del aula aprobada no acreditada.";
            }

            if (error.StartsWith("PROYECCION_COBERTURA_ERROR", StringComparison.Ordinal))
            {
                return "Error al proyectar la cobertura del campeón.";
            }

            if (error.StartsWith("LECTURA_STAGING_ERROR", StringComparison.Ordinal))
            {
                return "Error al leer las sesiones de staging.";
            }

            return "Revise el diagnóstico de cobertura.";
        }
    }
}
