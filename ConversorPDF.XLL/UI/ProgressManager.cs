using System;
using System.Diagnostics;
using System.Threading;
using System.Windows.Forms;
using ConversorPDF.XLL.Core.Helpers;

namespace ConversorPDF.XLL.UI
{
    public class ProgressManager : IDisposable
    {
        private int _disposed;
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
                        _form.HandleCreated += (s, e) => 
                        {
                            try { _formStarted.Set(); } catch (ObjectDisposedException) { }
                        };
                        Application.Run(_form);
                    }
                    catch (Exception ex)
                    {
                        LoggerHelper.Error("Error en el hilo de UI del progreso", ex);
                    }
                    finally
                    {
                        try { _formStarted.Set(); } catch (ObjectDisposedException) { }
                    }
                });
                _uiThread.SetApartmentState(ApartmentState.STA);
                _uiThread.IsBackground = true;
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
                _formStarted?.Set();
            }
        }

        public void AddProgress(double inc = 1)
        {
            _current += inc;
            if (_current > _max) _current = _max;

            string etaText = CalculateETA();
            try
            {
                if (_form != null && !_form.IsDisposed && _form.IsHandleCreated)
                {
                    double currentSnapshot = _current;
                    string etaSnapshot = etaText;
                    _form.BeginInvoke(new Action(() =>
                    {
                        try { _form.UpdateProgress(currentSnapshot, etaSnapshot); }
                        catch (Exception ex) { LoggerHelper.Warning($"Error en UpdateProgress: {ex.Message}"); }
                    }));
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Warning($"Error actualizando el progreso: {ex.Message}");
            }
        }

        public void ReportProgress(double value)
        {
            _current = value;
            if (_current > _max) _current = _max;

            string etaText = CalculateETA();
            try
            {
                if (_form != null && !_form.IsDisposed && _form.IsHandleCreated)
                {
                    double currentSnapshot = _current;
                    string etaSnapshot = etaText;
                    _form.BeginInvoke(new Action(() =>
                    {
                        try { _form.UpdateProgress(currentSnapshot, etaSnapshot); }
                        catch (Exception ex) { LoggerHelper.Warning($"Error en UpdateProgress: {ex.Message}"); }
                    }));
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Warning($"Error actualizando el progreso: {ex.Message}");
            }
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
                    _form.BeginInvoke(new Action(() => 
                    {
                        try { _form.Close(); }
                        catch (Exception ex) { LoggerHelper.Warning($"Error cerrando form: {ex.Message}"); }
                    }));
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Warning($"Error cerrando el formulario de progreso: {ex.Message}");
            }
        }

        public void CloseAndWait(int timeoutMs = 500)
        {
            try
            {
                if (_form != null && !_form.IsDisposed && _form.IsHandleCreated)
                {
                    _form.BeginInvoke(new Action(() => 
                    {
                        try { _form.Close(); }
                        catch (Exception ex) { LoggerHelper.Warning($"Error cerrando form: {ex.Message}"); }
                    }));
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Warning($"Error en CloseAndWait: {ex.Message}");
            }

            if (_uiThread != null && _uiThread.IsAlive)
            {
                _uiThread.Join(timeoutMs);
            }
        }

        public void Dispose()
        {
            if (Interlocked.Exchange(ref _disposed, 1) == 1) return;
            CloseAndWait(500);
            _stopwatch?.Stop();
            _formStarted?.Dispose();
        }
    }
}