using System;

namespace ConversorPDF.XLL.Academic.Domain.Enums
{
    /// <summary>
    /// Identifica el código de causa raíz ante bloqueos, fallas de validación o inconsistencias
    /// durante la preparación y ejecución del Algoritmo Genético en Módulo S (Fase 280).
    /// </summary>
    public enum AgCausaFallo
    {
        /// <summary> Operación normal sin fallos detectados. </summary>
        Ninguna = 0,

        /// <summary> No se ha seleccionado o no existe un período académico activo. </summary>
        SinPeriodo = 1,

        /// <summary> No existe una ventana temporal válida configurada para el período. </summary>
        SinVentana = 2,

        /// <summary> No existen registros en ag_aula para los criterios seleccionados. </summary>
        SinAulas = 3,

        /// <summary> No se encontró correspondencia exacta de AulaRef para el anclaje de la corrida. </summary>
        SinAnclajeAula = 4,

        /// <summary> El programa o catálogo no tiene asignaturas activas registradas. </summary>
        SinAsignaturas = 5,

        /// <summary> El aula no cuenta con vinculaciones en ag_aula_asignatura. </summary>
        SinAulaAsignaturas = 6,

        /// <summary> No hay docentes disponibles en el pool de terceros_asignaturas. </summary>
        SinPoolDocente = 7,

        /// <summary> No se pudieron cargar bloques horarios para la subunidad actual. </summary>
        SinBloques = 8,

        /// <summary> La demanda horaria de asignaturas supera la capacidad de la ventana temporal. </summary>
        VentanaInfactible = 9,

        /// <summary> La estructura de tablas en base de datos no contiene las columnas requeridas. </summary>
        EsquemaInvalido = 10,

        /// <summary> Excepción de infraestructura o error no controlado durante la ejecución. </summary>
        ExcepcionNoControlada = 11,

        /// <summary> La solución calculada no generó vinculaciones o sesiones en staging (Fase 383). </summary>
        StagingVacio = 12,

        /// <summary> Horas/día incoherentes con el nivel académico del programa (D4', Fase 1120). </summary>
        HorasDiaInvalidoNivel = 13,

        /// <summary> El aula no tiene jornada/horas/días de ventana persistidos en ag_aula (I1, Fase 1192). </summary>
        VentanaAulaVacia = 14,

        /// <summary> La solución no satisface el certificado de cobertura del manifiesto (Fase 1873). </summary>
        CoberturaNoCertificada = 15
    }
}
