using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using ConversorPDF.XLL.Academic.Domain.Models.Genetica;
using FluentAssertions;
using Xunit;

namespace ConversorPDF.Tests.Unit
{
    /// <summary>
    /// Test guardián de registro automático de métodos en Excel-DNA 1.8.0 (Fase 1034).
    /// Previene colisiones de nombres globales en funciones registradas en Excel.
    /// </summary>
    public class ExcelDnaRegistrationGuardTests
    {
        private static readonly HashSet<Type> SupportedParameterTypes = new HashSet<Type>
        {
            typeof(double),
            typeof(string),
            typeof(DateTime),
            typeof(double[]),
            typeof(double[,]),
            typeof(object),
            typeof(object[]),
            typeof(object[,]),
            typeof(bool),
            typeof(int),
            typeof(short),
            typeof(ushort),
            typeof(decimal),
            typeof(long)
        };

        private static bool EsTipoRetornoSoportado(Type returnType)
        {
            if (returnType == typeof(void))
            {
                return true;
            }
            return SupportedParameterTypes.Contains(returnType);
        }

        private static bool EsMetodoRegistrableExcelDna180(MethodInfo method)
        {
            if (method == null) return false;
            if (!method.IsPublic || !method.IsStatic) return false;
            if (method.IsAbstract || method.IsGenericMethod) return false;
            if (!EsTipoRetornoSoportado(method.ReturnType)) return false;

            ParameterInfo[] parameters = method.GetParameters();
            for (int i = 0; i < parameters.Length; i++)
            {
                if (!SupportedParameterTypes.Contains(parameters[i].ParameterType)) return false;
            }
            return true;
        }

        private static List<MethodInfo> ObtenerMetodosRegistrables()
        {
            Assembly assembly = typeof(AgClaveCohorte).Assembly;
            Type[] types;
            try
            {
                types = assembly.GetTypes();
            }
            catch (ReflectionTypeLoadException ex)
            {
                types = ex.Types.Where(t => t != null).ToArray();
            }

            var metodosRegistrables = new List<MethodInfo>();
            foreach (Type type in types)
            {
                if (!type.IsClass || !type.IsPublic || type.IsArray) continue;
                if (type.IsGenericType && type.ContainsGenericParameters) continue;
                if (type.Namespace == "My") continue;

                MethodInfo[] methods = type.GetMethods(BindingFlags.Public | BindingFlags.Static);
                foreach (MethodInfo method in methods)
                {
                    if (EsMetodoRegistrableExcelDna180(method))
                    {
                        metodosRegistrables.Add(method);
                    }
                }
            }

            return metodosRegistrables;
        }

        [Fact]
        public void MetodosRegistrables_WhenInspectedUnderExcelDna180Rules_ShouldNotHaveDuplicateNames()
        {
            // Arrange
            List<MethodInfo> metodos = ObtenerMetodosRegistrables();

            // Act
            var duplicados = metodos
                .GroupBy(m => m.Name, StringComparer.OrdinalIgnoreCase)
                .Where(g => g.Count() > 1)
                .Select(g => new
                {
                    Nombre = g.Key,
                    Metodos = string.Join(", ", g.Select(m => m.DeclaringType.FullName + "." + m.Name))
                })
                .ToList();

            // Assert
            duplicados.Should().BeEmpty("Excel-DNA 1.8.0 registra funciones por nombre sin prefijo de clase y arroja advertencia en duplicados");
        }

        [Fact]
        public void MetodosRegistrables_WhenInspectedUnderExcelDna180Rules_ShouldNotContainMethodNamedDe()
        {
            // Arrange
            List<MethodInfo> metodos = ObtenerMetodosRegistrables();

            // Act
            var metodosDe = metodos
                .Where(m => string.Equals(m.Name, "De", StringComparison.OrdinalIgnoreCase))
                .Select(m => m.DeclaringType.FullName + "." + m.Name)
                .ToList();

            // Assert
            metodosDe.Should().BeEmpty("ningún método registrable en Excel debe llamarse 'De' (Fase 1033)");
        }
    }
}
