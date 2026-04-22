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
        void RenombrarSolicitudes(int inicio, int fin);
        
        /// <summary> Unifica los expedientes de viáticos en archivos únicos por comisionado. </summary>
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
    }
}
