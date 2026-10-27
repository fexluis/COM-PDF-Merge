namespace ConversorPDF.XLL.Academic.UI.Enums
{
    using System;

    /// <summary>
    /// Índices de columna base-0 para la plantilla PROGRAMACIÓN ACADÉMICA PRESUPUESTAL.
    /// Al usar con la API COM de Excel (Cells[fila, col]) sumar siempre 1.
    /// Referencia: fila 12 de la hoja "PROGRAMACIÓN" del archivo de plantilla oficial.
    /// </summary>
    public enum ExcelColumnas
    {
        // ── DATOS DE LA ASIGNATURA (cols A–O) ──────────────────────────────
        Num = 0,   // A  — N°
        Territorial = 1,   // B  — TERRITORIAL
        Cetap = 2,   // C  — CETAP
        CodMateria = 3,   // D  — COD. MATERIA
        NombreMateria = 4,   // E  — NOMBRE MATERIA
        Creditos = 5,   // F  — CRÉDITOS
        Modalidad = 6,   // G  — MODALIDAD
        Programa = 7,   // H  — PROGRAMA
        Metodologia = 8,   // I  — METODOL.
        Nivel = 9,   // J  — NIVEL
        Grupo = 10,  // K  — GRUPO
        HorasAsignadasDocente1 = 11,  // L  — INTENSIDAD HORARIA
        NumTutorias = 12,  // M  — N° de tutorías
        NumTutoriasPresenciales = 13,  // N  — N° de tutorías presenciales
        NumTutoriasVirtuales = 14,  // O  — N° de tutorías virtuales

        // ── FÓRMULAS (no escribir) ──────────────────────────────────────────
        Revisar = 15,  // P  — REVISAR (fórmula automática)

        // ── DATOS DEL DOCENTE (cols Q–AE) ──────────────────────────────────
        Identificacion = 16,  // Q  — IDENTIFICACIÓN
        NombreDocente = 17,  // R  — NOMBRE (fórmula automática — no escribir)
        Vinculacion = 18,  // S  — VINCULACIÓN
        HorasAsignadasDocente2 = 19,  // T  — HORAS ASIGNADAS AL DOCENTE (fórmula — no escribir)
        FechaInicioVinculacion = 20,  // U  — FECHA DE INICIO VINCULACIÓN
        FechaFinVinculacion = 21,  // V  — FECHA FIN VINCULACIÓN
        Categoria = 22,  // W  — CATEGORÍA (fórmula — no escribir)
        Prestaciones = 23,  // X  — PRESTACIONES (PENSIONADO / NOPENSIONADO)
        Llave = 24,  // Y  — LLAVE (fórmula — no escribir)
        PuntajeOcasionales = 25,  // Z  — PUNTAJE OCASIONALES 

        // ── FÓRMULAS (cols AA–AF) (no escribir) ──────────────────────────────────────────
        ValorVinculacionOcasion = 26,  // AA — VALOR VINCULACIÓN OCASIONALES (fórmula)
        ValorHora = 27,  // AB — VALOR HORA (fórmula)
        ValorHoraPrestaciones = 28,  // AC — VALOR HORA + PRESTACIONES (fórmula)
        ValorPrestaciones = 29,  // AD — VALOR PRESTACIONES (fórmula)
        ValorTotalVinculacion = 30,  // AE — VALOR TOTAL VINCULACIÓN + PRESTACIONES (fórmula)
        TotalVinculacion4x1000 = 31,  // AF — TOTAL VINCULACIÓN + 4*1000 (fórmula)

        // ── VIÁTICOS (cols AG–AI) ───────────────────────────────────────────
        RequiereViaticos = 32,  // AG — ¿REQUIERE VIÁTICOS? ("SI"/"NO")
        SalarioParaCalculoVia = 33,  // AH — SALARIO PARA CALCULO DE VIÁTICOS
                                     //      Solo para CARRERA, OCASIONAL, ESPECIAL/VISITANTE.
                                     //      Dejar vacío para CATEDRA.
        TarifaPlanta = 34,  // AI — TARIFA PLANTA (fórmula — no escribir)
        TarifaCatedra = 35,  // AJ — TARIFA CATEDRA ("Tarifa A" si tiene viáticos/viajes, vacío si no)

        // ── COMISIÓN 1 (cols AK–AL) ─────────────────────────────────────────
        FechaInicio1 = 36,  // AK — FECHA INICIO 1
        FechaFin1 = 37,  // AL — FECHA FIN 1
        Dias1 = 38,  // AM — DÍAS (fórmula — no escribir)

        // ── COMISIÓN 2 (cols AN–AO) ─────────────────────────────────────────
        FechaInicio2 = 39,  // AN — FECHA INICIO 2
        FechaFin2 = 40,  // AO — FECHA FIN 2
        Dias2 = 41,  // AP — DÍAS2 (fórmula — no escribir)

        // ── COMISIÓN 3 (cols AQ–AR) ─────────────────────────────────────────
        FechaInicio3 = 42,  // AQ — FECHA INICIO 3
        FechaFin3 = 43,  // AR — FECHA FIN 3
        Dias3 = 44,  // AS — DÍAS3 (fórmula — no escribir)

        // ── COMISIÓN 4 (cols AT–AU) ─────────────────────────────────────────
        FechaInicio4 = 45,  // AT — FECHA INICIO 4
        FechaFin4 = 46,  // AU — FECHA FIN 4
        Dias4 = 47,  // AV — DÍAS4 (fórmula — no escribir)

        // ── COMISIÓN 5 (cols AW–AX) ─────────────────────────────────────────
        FechaInicio5 = 48,  // AW — FECHA INICIO 5
        FechaFin5 = 49,  // AX — FECHA FIN 5
        Dias5 = 50,  // AY — DÍAS5 (fórmula — no escribir)

        // ── COMISIÓN 6 (cols AZ–BA) ─────────────────────────────────────────
        FechaInicio6 = 51,  // AZ — FECHA INICIO 6
        FechaFin6 = 52,  // BA — FECHA FIN 6
        Dias6 = 53,  // BB — DÍAS6 (fórmula — no escribir)

        // ── COMISIÓN 7 (cols BC–BD) ─────────────────────────────────────────
        FechaInicio7 = 54,  // BC — FECHA INICIO 7
        FechaFin7 = 55,  // BD — FECHA FIN 7
        Dias7 = 56,  // BE — DÍAS7 (fórmula — no escribir)

        // ── COMISIÓN 8 (cols BF–BG) ─────────────────────────────────────────
        FechaInicio8 = 57,  // BF — FECHA INICIO 8
        FechaFin8 = 58,  // BG — FECHA FIN 8
        Dias8 = 59,  // BH — DÍAS8 (fórmula — no escribir)

        // ── DATOS DE RUTA (cols BI–BN) ──────────────────────────────────────
        Origen = 60,  // BI — ORIGEN (Tercero.Residencia)
        Escala = 61,  // BJ — ESCALA (ciudades intermedias de la ruta multi-sede)
        Destino = 62,  // BK — DESTINO (última ciudad visitada en el bloque de viáticos)
        CodRuta = 63,  // BL — COD_RUTA
        VerificacionRuta = 64,  // BM — VERIFICACIÓN RUTA (origen-escala-destino-origen)
        GastosTerrestres = 65,  // BN — GASTOS TERRESTRES (SesionListItem.TransporteSede)

        // ── TIQUETES AÉREOS (cols BO–BQ) ────────────────────────────────────
        RequiereTiquetesAereos = 66,  // BO — "¿REQUIERE TIQUETES AÉREOS?"  isTerminalAereo = 1 -> "SI" | isTerminalAereo = 0 -> "NO"
        TerminalAerea = 67,  // BP — "TERMINAL AÉREA"
        GastoTerminalAerea = 68,  // BQ — GASTO TERMINAL AÉREA (fórmula — no escribir)
        TotalViatico = 69,  // BR — TOTAL VIÁTICO (fórmula — no escribir)
    }
}
