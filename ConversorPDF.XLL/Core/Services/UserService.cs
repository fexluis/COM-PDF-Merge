using System;
using ConversorPDF.XLL.Core.Services;
using ConversorPDF.XLL.Core.Repositories;
using ConversorPDF.XLL.Core.Processors;
using ConversorPDF.XLL.Core.Models;
using ConversorPDF.XLL.Core.Interfaces;
using ConversorPDF.XLL.Core.Helpers;
using ConversorPDF.XLL.Core.Enums;
using System.Resources;
using System.Security.Principal;
namespace ConversorPDF.XLL.Core.Services
{
    /// <summary>
    /// Servicio que gestiona la identidad del usuario y la recuperación de firmas digitales desde los recursos del ensamblado.
    /// </summary>
    public class UserService : IUserService
    {
        private static readonly ResourceManager _resourceManager = 
            new ResourceManager("ConversorPDF.XLL.Properties.Signatures", typeof(UserService).Assembly);

        /// <summary> Obtiene el nombre completo de la identidad de Windows (ej: DOMINIO\Usuario). </summary>
        public string GetUsuarioWindows()
        {
            try 
            {
                return WindowsIdentity.GetCurrent().Name;
            }
            catch
            {
                return "LUIS.BECHARA"; // Default fallback
            }
        }

        /// <summary>
        /// Recupera la firma en Base64 asociada a un nombre de usuario específico.
        /// Si el usuario no tiene una firma configurada, devuelve la firma por defecto (Luis).
        /// </summary>
        /// <param name="username">Nombre de usuario de red a buscar.</param>
        /// <returns>Cadena Base64 con la imagen de la firma.</returns>
        public string GetFirmaBase64(string username)
        {
            if (string.IsNullOrEmpty(username)) return GetResourceString("FirmaLuis");

            if (username == "ALEXANDER.MONTANO")
                return GetResourceString("FirmaMontano");
            else if (username == "LUIS.BECHARA" || username == "ANTOVILL" || username == "ELIZABETH.CORDOBAO")
                return GetResourceString("FirmaVillegas");
            else if (username == "EDUIN.PALACIOS")
                return GetResourceString("FirmaLuis");
            
            return GetResourceString("FirmaLuis");
        }

        /// <summary> Acceso auxiliar al ResourceManager para obtener cadenas de recursos. </summary>
        private string GetResourceString(string name)
        {
            return _resourceManager.GetString(name) ?? "";
        }

        /// <summary>
        /// Obtiene los datos de identificación (iniciales y nombre completo) asociados a un usuario.
        /// </summary>
        public UserIdentification GetIdentification(string username)
        {
            if (string.IsNullOrEmpty(username)) return null;

            string normalizedUser = username.ToUpper();

            if (normalizedUser == "LUIS.BECHARA" || normalizedUser == "ELIZABETH.CORDOBAO")
            {
                return new UserIdentification("MHavillega", "ANTONIO ALEXANDER VILLEGAS");
            }
            
            if (normalizedUser == "EDUIN.PALACIOS")
            {
                return new UserIdentification("MHlbechara", "LUIS EDUARDO BECHARA POLANIA");
            }

            return null;
        }
    }
}
