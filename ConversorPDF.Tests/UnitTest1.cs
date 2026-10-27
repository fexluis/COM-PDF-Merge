using Xunit;
using Moq;
using ConversorPDF.XLL.Core.Services;
using ConversorPDF.XLL.Core.Helpers;
using ConversorPDF.XLL.Core.Interfaces;

namespace ConversorPDF.Tests
{
    public class UnitTest1
    {
        [Theory]
        [InlineData("hola mundo", "Hola Mundo")]
        [InlineData("PDF SHARP", "Pdf Sharp")]
        [InlineData("esap", "Esap")]
        public void CapitalizarPalabras_ShouldReturnCorrectFormat(string input, string expected)
        {
            // Act
            var result = App.CapitalizarPalabras(input);

            // Assert
            Assert.Equal(expected, result);
        }

        [Fact]
        public void GetUserName_WhenServiceReturnsUser_ShouldReturnUpperPart()
        {
            // Arrange
            var mockUserService = new Mock<IUserService>();
            var mockPdfService = new Mock<IPdfService>();
            var mockViaticosService = new Mock<IViaticosService>();
            var mockContadorService = new Mock<IContadorService>();
            mockUserService.Setup(s => s.GetUsuarioWindows()).Returns(@"DOMAIN\LUIS.BECHARA");
            var sut = new App(mockUserService.Object, mockPdfService.Object, mockViaticosService.Object, mockContadorService.Object);

            // Act
            var result = sut.GetUserName();

            // Assert
            Assert.Equal("LUIS.BECHARA", result);
        }

        [Fact]
        public void GetUserName_WhenServiceReturnsNull_ShouldReturnFallback()
        {
            // Arrange
            var mockUserService = new Mock<IUserService>();
            var mockPdfService = new Mock<IPdfService>();
            var mockViaticosService = new Mock<IViaticosService>();
            var mockContadorService = new Mock<IContadorService>();
            mockUserService.Setup(s => s.GetUsuarioWindows()).Returns((string)null);
            var sut = new App(mockUserService.Object, mockPdfService.Object, mockViaticosService.Object, mockContadorService.Object);

            // Act
            var result = sut.GetUserName();

            // Assert
            var expectedFallback = System.Environment.UserName?.ToUpper() ?? "UNKNOWN_USER";
            Assert.Equal(expectedFallback, result);
        }
    }
}
