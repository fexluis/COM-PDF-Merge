using System;
using ConversorPDF.XLL.Core.Services;
using ConversorPDF.XLL.Core.Repositories;
using ConversorPDF.XLL.Core.Processors;
using ConversorPDF.XLL.Core.Models;
using ConversorPDF.XLL.Core.Interfaces;
using ConversorPDF.XLL.Core.Helpers;
using ConversorPDF.XLL.Core.Enums;
namespace ConversorPDF.XLL.Core.Helpers
{
    public sealed class ExcelAppStateHelper : IDisposable
    {
        private readonly dynamic _app;
        private readonly bool _hasScreenUpdating;
        private readonly bool _hasDisplayAlerts;
        private readonly bool _hasEnableEvents;
        private readonly bool _hasCalculation;
        private readonly object _screenUpdating;
        private readonly object _displayAlerts;
        private readonly object _enableEvents;
        private readonly object _calculation;

        public ExcelAppStateHelper(dynamic app)
        {
            _app = app;
            try { _screenUpdating = app.ScreenUpdating; _hasScreenUpdating = true; } catch { }
            try { _displayAlerts = app.DisplayAlerts; _hasDisplayAlerts = true; } catch { }
            try { _enableEvents = app.EnableEvents; _hasEnableEvents = true; } catch { }
            try { _calculation = app.Calculation; _hasCalculation = true; } catch { }
        }

        public void ApplyOptimized()
        {
            try { _app.ScreenUpdating = false; } catch { }
            try { _app.DisplayAlerts = false; } catch { }
            try { _app.EnableEvents = false; } catch { }
            try { _app.Calculation = -4135; /*xlCalculationManual*/ } catch { }
        }

        public IDisposable ApplyOptimizedAndReturn()
        {
            ApplyOptimized();
            return this;
        }

        public void Dispose()
        {
            try { if (_hasCalculation) _app.Calculation = _calculation; } catch { }
            try { if (_hasEnableEvents) _app.EnableEvents = _enableEvents; } catch { }
            try { if (_hasDisplayAlerts) _app.DisplayAlerts = _displayAlerts; } catch { }
            try { if (_hasScreenUpdating) _app.ScreenUpdating = _screenUpdating; } catch { }
        }
    }
}