using ConversorPDF.XLL.Core.Enums;
using ConversorPDF.XLL.Core.Services;
using ConversorPDF.XLL.Core.Repositories;
using ConversorPDF.XLL.Core.Processors;
using ConversorPDF.XLL.Core.Models;
using ConversorPDF.XLL.Core.Interfaces;
using ConversorPDF.XLL.Core.Helpers;
namespace ConversorPDF.XLL.Core.Interfaces
{
    /// <summary>
    /// Interfaz para la gestión de usuarios e identidades (firmas, perfiles).
    /// </summary>
    public interface IUserService
    {
        /// <summary> Obtiene el nombre de usuario de Windows de la sesión actual. </summary>
        string GetUsuarioWindows();
        
        /// <summary> Obtiene la representación Base64 de la firma digital de un usuario. </summary>
        string GetFirmaBase64(string username);

        /// <summary> Obtiene los datos de identificación (iniciales y nombre completo) de un usuario. </summary>
        UserIdentification GetIdentification(string username);
    }
}