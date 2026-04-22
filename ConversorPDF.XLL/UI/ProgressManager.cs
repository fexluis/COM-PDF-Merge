using System;
using ConversorPDF.XLL.Core.Services;
using ConversorPDF.XLL.Core.Repositories;
using ConversorPDF.XLL.Core.Processors;
using ConversorPDF.XLL.Core.Models;
using ConversorPDF.XLL.Core.Interfaces;
using ConversorPDF.XLL.Core.Helpers;
using ConversorPDF.XLL.Core.Enums;
using System.Diagnostics;
using System.Threading;
using System.Windows.Forms;
namespace ConversorPDF.XLL.UI
{
    public class ProgressManager : IDisposable
    {
        private ProgressForm _form;
        private Thread _uiThread;
        private Stopwatch _stopwatch;
        private double _max;
        private double _current;
        private readonly ManualResetEvent _formStarted = new ManualResetEvent(false);

        public ProgressManager(string title, double max = 100)
        {
            _max = max;
            _current = 0;
            _stopwatch = new Stopwatch();

            try
            {
                _uiThread = new Thread(() =>
                {
                    try
                    {
                        _form = new ProgressForm(title, max);
                        _form.HandleCreated += (s, e) => _formStarted.Set();
                        Application.Run(_form);
                    }
                    catch (Exception ex)
                    {
                        LoggerHelper.Error("Error en el hilo de UI del progreso", ex);
                        _formStarted.Set();
                    }
                });
                _uiThread.SetApartmentState(ApartmentState.STA);
                _uiThread.Start();

                if (!_formStarted.WaitOne(2000))
                {
                    LoggerHelper.Warning("El formulario de progreso tardó más de 2 segundos en iniciar.");
                }
                _stopwatch.Start();
                LoggerHelper.Info($"ProgressManager iniciado: {title} (max: {max})");
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("Error iniciando ProgressManager", ex);
            }
        }

        public void AddProgress(double inc = 1)
        {
            _current += inc;
            if (_current > _max) _current = _max;

            string etaText = CalculateETA();
            _form?.UpdateProgress(_current, etaText);
        }

        private string CalculateETA()
        {
            if (_current <= 0) return "Iniciando...";

            long elapsedMs = _stopwatch.ElapsedMilliseconds;
            double msPerUnit = (double)elapsedMs / _current;
            double remainingUnits = _max - _current;
            double remainingMs = msPerUnit * remainingUnits;

            TimeSpan t = TimeSpan.FromMilliseconds(remainingMs);
            
            if (_current >= _max) return "Completado";

            if (t.TotalHours >= 1)
                return $"Restante: {(int)t.TotalHours}h {t.Minutes}m {t.Seconds}s";
            if (t.TotalMinutes >= 1)
                return $"Restante: {t.Minutes}m {t.Seconds}s";
            
            return $"Restante: {t.Seconds}s";
        }

        public void Close()
        {
            try
            {
                if (_form != null && !_form.IsDisposed && _form.IsHandleCreated)
                {
                    _form.Invoke(new Action(() => 
                    {
                        try { _form.Close(); } catch { } 
                    }));
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Warning($"Error cerrando el formulario de progreso: {ex.Message}");
            }
        }

        public void Dispose()
        {
            Close();
            _stopwatch.Stop();
        }
    }
}