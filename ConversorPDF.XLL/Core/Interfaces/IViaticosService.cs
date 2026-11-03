using System;
using ConversorPDF.XLL.Core.Services;
using ConversorPDF.XLL.Core.Repositories;
using ConversorPDF.XLL.Core.Processors;
using ConversorPDF.XLL.Core.Interfaces;
using ConversorPDF.XLL.Core.Helpers;
using ConversorPDF.XLL.Core.Enums;
using System.Collections.Generic;
using ConversorPDF.XLL.Core.Models;

namespace ConversorPDF.XLL.Core.Interfaces
{
    /// <summary>
    /// Interfaz que orquestra el procesamiento de viáticos en sus diferentes modalidades.
    /// Actúa como el servicio principal para el módulo de Viáticos.
    /// </summary>
    public interface IViaticosService
    {
        /// <summary> Obtiene la configuración actual de viáticos. </summary>
        ViaticosConfig GetConfig(DateTime? overrideInicio = null, DateTime? overrideFin = null, string overrideCdp = null);
        
        /// <summary> Busca y obtiene la información de un empleado por su cédula. </summary>
        Empleado GetEmpleado(string cedula);
        
        /// <summary> Extrae los datos de un viaje específico desde una fila de Excel. </summary>
        Viaje GetViajeFromRow(object[,] data, int dataRow);
        
        /// <summary> Ejecuta el procesamiento de viáticos individuales. </summary>
        int ProcesarIndividual(Action<int, int> progressCallback = null, DateTime? overrideInicio = null, DateTime? overrideFin = null, string overrideCdp = null);
        
        /// <summary> Ejecuta el procesamiento masivo de viáticos. </summary>
        int ProcesarMasivo(Action<int, int> progressCallback = null);
        
        /// <summary> Ejecuta la validación de cumplidos e informes de actividades. </summary>
        int ProcesarCumplidos(Action<int, int> progressCallback = null, Action onComplete = null);
        
        /// <summary> Ejecuta el proceso solo para la fila actualmente seleccionada. </summary>
        int ProcesarFilaSeleccionada(Action<int, int> progressCallback = null);
        
        /// <summary> Renombra archivos PDF de solicitudes basándose en el rango e información extraída. </summary>
        /// <returns> Cantidad de archivos renombrados con éxito, o -1 en caso de error. </returns>
        int RenombrarSolicitudes(int inicio, int fin);

        /// <summary> Renombra solicitudes devolviendo el desglose completo del proceso. </summary>
        ResultadoRenombrado RenombrarSolicitudesConDetalle(int inicio, int fin);

        /// <summary> Obtiene el primer y último registro de la columna B de CONSULTAS. </summary>
        RangoSolicitudes ObtenerRangoSolicitudes();

        /// <summary> Verifica la presencia física de las solicitudes y marca INFORME = SIN SOLICITUD. </summary>
        ResultadoVerificacionSolicitudes VerificarSolicitudesPresentes(int inicio, int fin);
        
        /// <summary> Unifica los expedientes de viáticos en archivos únicos por comisionado. </summary>
        /// <returns> Cantidad de expedientes unificados con éxito, o -1 en caso de error. </returns>
        int UnificarSolicitudes(int inicio, int fin);
        
        /// <summary> Limpia el caché interno de empleados y configuraciones. </summary>
        void ClearCache();

        /// <summary> Verifica si la hoja cumple con el formato estándar de viáticos. </summary>
        bool ValidarFormatoComisionado(dynamic sheet);
        
        /// <summary> Obtiene la lista cargada de CDPs válidos para la vigencia. </summary>
        List<string> GetCdpList();
        
        /// <summary> Recupera la configuración de selección del formulario (Rango de filas y CDP). </summary>
        string[] GetSelectionConfig();
        
        /// <summary> Guarda la configuración de selección del formulario para persistencia. </summary>
        void FilaSeleccionadaSaveConfig(object inicio, object fin, string cdp = null);
        
        /// <summary> Obtiene una variable global de Excel. </summary>
        string GetGlobalVariable(string varName);
        
        /// <summary> Obtiene los datos de configuración avanzada de viáticos. </summary>
        object[,] GetAdvancedConfigRows();
        
        /// <summary> Guarda los datos de configuración avanzada. </summary>
        void SaveAdvancedConfigRows(object[,] data);
        
        /// <summary> Acceso genérico a datos de tablas de Excel configuradas. </summary>
        object[,] GetTableData(string tableName);
        
        /// <summary> Guarda datos en una tabla específica de Excel. </summary>
        void SaveTableData(string tableName, object[,] data);

        /// <summary>
        /// Obtiene la carpeta de informes del libro activo (&lt;carpeta del libro&gt;\Informes).
        /// </summary>
        /// <returns>Ruta absoluta, o cadena vacía si no pudo resolverse.</returns>
        string ObtenerCarpetaInformes();

        /// <summary>
        /// Ejecuta el proceso "Buscar Informes": renombrado al estándar, unificación de parejas
        /// INF + CUM y, opcionalmente, la auditoría de la columna INFORME de CONSULTAS.
        /// </summary>
        /// <param name="inicio">Solicitud inicial del rango (0 = sin filtro).</param>
        /// <param name="fin">Solicitud final del rango (0 = sin filtro).</param>
        /// <param name="unificarParejas">True para unificar parejas INF + CUM.</param>
        /// <param name="auditarConsultas">True para escribir la columna INFORME.</param>
        /// <param name="simular">True para calcular sin modificar archivos ni Excel.</param>
        /// <param name="carpetaRaiz">Carpeta de informes a procesar; si es null se usa la del libro.</param>
        /// <param name="progress">Callback opcional de progreso (porcentaje 0-100).</param>
        /// <returns>Resultado combinado con contadores y detalle.</returns>
        ResultadoBusquedaInformes ProcesarInformes(int inicio, int fin, bool unificarParejas,
                                                   bool auditarConsultas, bool simular,
                                                   string carpetaRaiz = null,
                                                   Action<int, int> progress = null);
    }
}
