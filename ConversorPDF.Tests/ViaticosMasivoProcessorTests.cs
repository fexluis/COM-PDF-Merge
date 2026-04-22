using Xunit;
using Moq;
using ConversorPDF.XLL.Core.Services;
using ConversorPDF.XLL.Core.Repositories;
using ConversorPDF.XLL.Core.Processors;
using ConversorPDF.XLL.Core.Interfaces;
using System;
using System.Collections.Generic;

namespace ConversorPDF.Tests
{
    public class ViaticosMasivoProcessorTests
    {
        [Fact]
        public void GetCat1_WhenTableHas1BasedData_ShouldReturnCorrectSet()
        {
            // Arrange
            var mockExcel = new Mock<IExcelRepository>();
            var mockViaticos = new Mock<IViaticosRepository>();
            
            // Excel Interop arrays are 1-based.
            // Create a 2x2 array (base 1)
            // Array[row, col]
            // We use GetUpperBound to determine size, but GetLength works too.
            // Under net48, object[,] from Interop is indeed 1-based.
            // To simulate 1-based, we can create an array with lower bounds 1.
            
            var catData = Array.CreateInstance(typeof(object), new int[] { 2, 1 }, new int[] { 1, 1 }) as object[,];
            catData[1, 1] = "CAT1";
            catData[2, 1] = "CAT2";
            
            mockExcel.Setup(x => x.GetTableData("tblCategoria1")).Returns(catData);

            var sut = new ViaticosMasivoProcessor(mockExcel.Object, mockViaticos.Object);

            // Act
            var result = sut.GetCat1();

            // Assert
            Assert.Contains("CAT1", result);
            Assert.Contains("CAT2", result);
            Assert.Equal(2, result.Count);
        }

        [Fact]
        public void GetCat1_WhenTableIsEmpty_ShouldReturnEmptySet()
        {
            // Arrange
            var mockExcel = new Mock<IExcelRepository>();
            var mockViaticos = new Mock<IViaticosRepository>();
            
            mockExcel.Setup(x => x.GetTableData("tblCategoria1")).Returns((object[,])null);

            var sut = new ViaticosMasivoProcessor(mockExcel.Object, mockViaticos.Object);

            // Act
            var result = sut.GetCat1();

            // Assert
            Assert.Empty(result);
        }
    }
}
