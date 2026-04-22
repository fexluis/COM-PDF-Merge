using Xunit;
using ConversorPDF.XLL.Core.Services;
using System.Collections.Generic;
using System;

namespace ConversorPDF.Tests
{
    public class PdfServiceTests
    {
        [Fact]
        public void CombinarPDFs_WhenListIsEmpty_ShouldThrowException()
        {
            // Arrange
            var sut = new PdfService();
            var emptyList = new List<string>();
            string output = "dummy.pdf";

            // Act & Assert
            Assert.Throws<InvalidOperationException>(() => sut.CombinarPDFs(emptyList, output));
        }
    }
}
