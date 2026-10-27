using Xunit;
using FluentAssertions;
using System;
using System.Collections.Generic;
using ConversorPDF.XLL.Academic.Application.Mappers;
using ConversorPDF.XLL.Academic.Application.DTOs;
using ConversorPDF.XLL.Academic.Domain.Models;

namespace ConversorPDF.Tests.Academic.Application.Unit
{
    public class GridToSesionMapperTests
    {
        // ─── Helpers de construcción de DTOs ───────────────────────────────

        private static GridRowDto FilaMarcada(int idTercero = 1) => new GridRowDto
        {
            IsChecked = true, IsNewRow = false,
            Tag = new Sesion { IdTercero = idTercero, FechaSesion = new DateTime(2026, 4, 23), Horas = 8 },
            Horas = 8,
            FechaSesion = new DateTime(2026, 4, 23),
            HorarioSesion = "JUEVES|8"
        };

        private static GridRowDto FilaMarcadaSinTag() => new GridRowDto
        { IsChecked = true, IsNewRow = false, Tag = null };

        private static GridRowDto FilaDesmarcada() => new GridRowDto
        { IsChecked = false, IsNewRow = false, Tag = new Sesion { IdTercero = 1, FechaSesion = new DateTime(2026, 4, 23), Horas = 8 }, Horas = 8, FechaSesion = new DateTime(2026, 4, 23) };

        private static GridRowDto FilaNueva() => new GridRowDto
        { IsChecked = false, IsNewRow = true, Tag = null };

        // ─── TEST CRÍTICO C-01a: Reproductor del bug real ──────────────────

        [Fact]
        public void Extract_6FilasMarcadasCorrectamente_Retorna6Sesiones()
        {
            var rows = new List<GridRowDto>();
            for (int i = 0; i < 6; i++) rows.Add(FilaMarcada());
            rows.Add(FilaNueva()); 

            var resultado = GridToSesionMapper.Extract(rows, idTerceroMaster: 1, idVinculacion: 1);

            resultado.Should().HaveCount(6);
        }

        // El test anterior era inválido. La integridad se dispara si hay un desfase entre
        // lo que el usuario ve marcado y lo que el motor de extracción acepta.
        // En la implementación actual, los errores de validación disparan excepciones inmediatamente
        // dentro del loop, por lo que el chequeo final de conteo es una red de seguridad secundaria.

        [Fact]
        public void Extract_FilaMarcadaSinTag_LanzaExcepcion()
        {
            var rows = new List<GridRowDto> { FilaMarcadaSinTag() };

            Action act = () => GridToSesionMapper.Extract(rows, idTerceroMaster: 1, idVinculacion: 1);

            act.Should().Throw<InvalidOperationException>()
               .WithMessage("*Fila 1*")
               .And.Message.Should().Contain("modelo de sesión asociado");
        }

        [Fact]
        public void Extract_FilasConIsNewRow_SonIgnoradas()
        {
            var rows = new List<GridRowDto>
            {
                FilaMarcada(), FilaMarcada(), FilaNueva() 
            };

            var resultado = GridToSesionMapper.Extract(rows, idTerceroMaster: 1, idVinculacion: 1);

            resultado.Should().HaveCount(2);
        }

        [Fact]
        public void Extract_GridVacio_RetornaListaVacia()
        {
            var rows = new List<GridRowDto>();

            var resultado = GridToSesionMapper.Extract(rows, idTerceroMaster: 1, idVinculacion: 1);

            resultado.Should().BeEmpty();
        }

        [Fact]
        public void Extract_DocenteDesincronizado_LanzaExcepcion()
        {
            // Escenario: La fila tiene docente ID 1, pero el combo dice ID 2
            var rows = new List<GridRowDto> { FilaMarcada(idTercero: 1) }; 

            Action act = () => GridToSesionMapper.Extract(rows, idTerceroMaster: 2, idVinculacion: 1);

            act.Should().Throw<InvalidOperationException>()
               .WithMessage("*docente de esta sesión no coincide*");
        }

        [Fact]
        public void Extract_WhenEdicionManualEnVista_ShouldUpdateTagAndReturnSesion()
        {
            // Arrange: Tag tiene fecha 23/04/2026 y 8h, pero la vista fue editada manualmente a 24/04/2026 y 6h
            var dto = FilaMarcada();
            dto.FechaSesion = new DateTime(2026, 4, 24);
            dto.Horas = 6.0;

            // Act
            var resultado = GridToSesionMapper.Extract(new List<GridRowDto> { dto }, idTerceroMaster: 1, idVinculacion: 1);

            // Assert
            resultado.Should().HaveCount(1);
            resultado[0].FechaSesion.Should().Be(new DateTime(2026, 4, 24));
            resultado[0].Horas.Should().Be(6.0);
            ((Sesion)dto.Tag).FechaSesion.Should().Be(new DateTime(2026, 4, 24));
            ((Sesion)dto.Tag).Horas.Should().Be(6.0);
        }

        [Fact]
        public void Extract_WhenFechaHorasSincronizadasConTag_ShouldReturnSesion()
        {
            // Arrange: Vista y Tag sincronizados con nuevos valores tras edición manual
            var dto = FilaMarcada();
            var nuevoDia = new DateTime(2026, 9, 28);
            double nuevasHoras = 4.5;

            ((Sesion)dto.Tag).FechaSesion = nuevoDia;
            ((Sesion)dto.Tag).Horas = nuevasHoras;
            dto.FechaSesion = nuevoDia;
            dto.Horas = nuevasHoras;

            // Act
            var resultado = GridToSesionMapper.Extract(new List<GridRowDto> { dto }, idTerceroMaster: 1, idVinculacion: 1);

            // Assert
            resultado.Should().HaveCount(1);
            resultado[0].FechaSesion.Should().Be(nuevoDia);
            resultado[0].Horas.Should().Be(nuevasHoras);
        }

        [Fact]
        public void Extract_WhenHorasDentroDeToleranciaFloat_ShouldReturnSesion()
        {
            // Arrange: Diferencia infinitesimal menor a 0.0001
            var dto = FilaMarcada();
            ((Sesion)dto.Tag).Horas = 4.00001;
            dto.Horas = 4.0;

            // Act
            var resultado = GridToSesionMapper.Extract(new List<GridRowDto> { dto }, idTerceroMaster: 1, idVinculacion: 1);

            // Assert
            resultado.Should().HaveCount(1);
        }
    }
}
