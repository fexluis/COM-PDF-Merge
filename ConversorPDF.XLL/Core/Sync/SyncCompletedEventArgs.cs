using System;

namespace ConversorPDF.XLL.Core.Sync
{
    /// <summary>
    /// Datos del evento SyncCompleted que el SyncScheduler dispara
    /// al terminar cada operación de PUSH.
    ///
    /// Permite a la UI (Ribbon, notificaciones) reaccionar al resultado del sync
    /// sin acoplar directamente el servicio a la interfaz gráfica.
    /// </summary>
    internal sealed class SyncCompletedEventArgs : EventArgs
    {
        /// <summary>
        /// true si el PUSH terminó sin errores (puede ser éxito sin cambios).
        /// false si hubo una excepción durante el PUSH.
        /// </summary>
        internal bool Success { get; private set; }

        /// <summary>
        /// Número de bytes escritos al archivo .crsync.
        /// 0 si no había cambios o si hubo un error.
        /// </summary>
        internal long BytesSent { get; private set; }

        /// <summary>
        /// Número de cambios CRDT enviados en el PUSH.
        /// 0 si no había cambios pendientes.
        /// </summary>
        internal int ChangesCount { get; private set; }

        /// <summary>
        /// Excepción que ocurrió durante el PUSH.
        /// null si Success = true.
        /// </summary>
        internal Exception Error { get; private set; }

        /// <summary>
        /// Constructor para resultado de PUSH (exitoso o con error).
        /// </summary>
        internal SyncCompletedEventArgs(bool success, long bytesSent, int changesCount, Exception error)
        {
            Success = success;
            BytesSent = bytesSent;
            ChangesCount = changesCount;
            Error = error;
        }
    }
}
