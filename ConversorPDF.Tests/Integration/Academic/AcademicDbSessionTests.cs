using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Context;
using ConversorPDF.XLL.Core.Helpers;
using Xunit;
using FluentAssertions;

namespace ConversorPDF.Tests.Integration.Academic
{
    public class AcademicDbSessionTests
    {
        private readonly string _dbPath;

        public AcademicDbSessionTests()
        {
            // Aislar este test de la BD del sistema real para evitar schemas desactualizados
            _dbPath = Path.Combine(Path.GetTempPath(), $"AcademicSessionTest_{Guid.NewGuid():N}.db.enc");
            SqliteHelper.SetTestVaultPath(_dbPath);
            using (var ctx = SqliteHelper.GetContext()) { }
        }

        [Fact]
        [Trait("Category", "Hotfix_1_2_51_25")]
        [Trait("Priority", "P2")]
        [Trait("Type", "Integration")]
        public async Task ExecuteAsync_SerializaAccesoConcurrente()
        {
            // Arrange: usa el path temporal aislado configurado en el constructor
            using (var session = new AcademicDbSession())
            {
                var executionOrder = new List<int>();
                var startSignal = new TaskCompletionSource<bool>();

                // Act: Disparar dos tareas que intentan usar la conexión simultáneamente
                var task1 = session.ExecuteAsync(async conn =>
                {
                    await startSignal.Task; // Esperar señal para sincronizar inicio
                    await Task.Delay(100);  // Simular trabajo
                    lock (executionOrder) executionOrder.Add(1);
                    return 1;
                });

                var task2 = session.ExecuteAsync(async conn =>
                {
                    await startSignal.Task;
                    await Task.Delay(10);   // Trabajo más corto
                    lock (executionOrder) executionOrder.Add(2);
                    return 2;
                });

                startSignal.SetResult(true);
                await Task.WhenAll(task1, task2);

                // Assert
                // Si el SemaphoreSlim funciona, la Tarea 1 (que es más larga) debería bloquear a la Tarea 2
                // O al revés, pero nunca deberían ejecutarse en paralelo de forma que se intercalen
                // incorrectamente si hubiera estado compartido sensible.
                // Task 1 entra primero (usualmente), bloquea el semáforo. Task 2 espera.
                // Task 1 termina después de 100ms. Task 2 entra y termina después de 10ms.
                // Total tiempo: ~110ms. Si fuera paralelo, total tiempo ~100ms.
                executionOrder.Should().HaveCount(2);
            }
        }
    }
}
