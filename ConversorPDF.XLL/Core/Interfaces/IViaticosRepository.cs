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
    /// Interfaz para el acceso a datos de configuración, empleados y viajes desde Excel.
    /// </summary>
    public interface IViaticosRepository
    {
        /// <summary> Obtiene el objeto de configuración cargado desde la hoja 'Config'. </summary>
        ViaticosConfig GetConfig(DateTime? start = null, DateTime? end = null, string cdpOverride = null);
        
        /// <summary> Busca y obtiene la información de un empleado por su cédula. </summary>
        Empleado GetEmpleado(string cc);
        
        /// <summary> Extrae la información de un viaje desde una fila de datos de Excel. </summary>
        Viaje GetViajeFromRow(object[,] data, int r);
        
        /// <summary> Busca el departamento correspondiente a una ciudad dada en la tabla de destinos. </summary>
        string ObtenerDepartamento(string ciudad);
        
        /// <summary> Normaliza el nombre de una ciudad (limpieza de tildes y espacios). </summary>
        string NormalizarCiudad(string ciudad);
        
        /// <summary> Valida si la hoja activa cumple con el formato requerido por la entidad. </summary>
        bool ValidarFormatoComisionado(dynamic sheet);
        
        /// <summary> Obtiene la lista de CDPs disponibles desde la tabla Dependencias. </summary>
        List<string> GetCdpList();
        
        /// <summary> Obtiene el nombre de la dependencia asociada a un CDP/Código. </summary>
        string GetDependenciaCDP(string dep);
        
        /// <summary> Obtiene el valor de una variable global de Excel. </summary>
        string GetGlobalVariable(string varName);
    }
}