#if !NET5_0_OR_GREATER
using System;

namespace System.Runtime.CompilerServices
{
    /// <summary>
    /// Shim de ModuleInitializerAttribute para .NET Framework 4.8 (Fase 1067).
    /// Permite al compilador C# 9+ emitir inicializadores de módulo (.cctor) en ensamblados net48.
    /// </summary>
    [AttributeUsage(AttributeTargets.Method, Inherited = false)]
    internal sealed class ModuleInitializerAttribute : Attribute
    {
    }
}
#endif
