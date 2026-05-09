using ClosedXML.Excel;
using Kevull.MultiHeader.Core;

namespace Kevull.MultiHeader.ClosedXml.Test
{
    /// <summary>
    /// Unit tests for <see cref="ClosedXmlExcelWriter"/>.
    /// </summary>
    public class ClosedXmlExcelWriterTests : IDisposable
    {
        private readonly XLWorkbook _workbook;
        private readonly IXLWorksheet _worksheet;
        private readonly ClosedXmlExcelWriter _writer;

        public ClosedXmlExcelWriterTests()
        {
            _workbook = new XLWorkbook();
            _worksheet = _workbook.Worksheets.Add("TestSheet");
            _writer = new ClosedXmlExcelWriter(_workbook, _worksheet);
        }

        public void Dispose()
        {
            _workbook.Dispose();
        }

        [Fact]
        public void ApplyNativeFormat_NullFormat_ReturnsWithoutError()
        {
            _writer.ApplyNativeFormat(1, 1, null!);
        }

        [Fact]
        public void ApplyNativeFormat_ActionModifiesRange_RangeIsModified()
        {
            _writer.ApplyNativeFormat(2, 3, range =>
            {
                var xlRange = Assert.IsAssignableFrom<IXLRange>(range);
                xlRange.Value = "TestValue";
            });

            Assert.Equal("TestValue", _worksheet.Cell(2, 3).GetValue<string>());
        }

        [Fact]
        public void ApplyNativeFormat_ActionParameterType_IsRange()
        {
            Type? capturedType = null;

            _writer.ApplyNativeFormat(1, 1, range =>
            {
                capturedType = range?.GetType();
            });

            Assert.NotNull(capturedType);
            Assert.True(typeof(IXLRange).IsAssignableFrom(capturedType));
        }

        [Fact]
        public void ApplyFormat_NullFormat_DoesNothing()
        {
            _writer.ApplyFormat(1, 1, 2, 2, null!);

            Assert.False(_worksheet.Cell(1, 1).Style.Font.Bold);
        }

        [Fact]
        public void ApplyFormat_SingleCell_AppliesFormat()
        {
            var format = new CellFormat
            {
                Bold = true,
                FontSize = 14.0f,
                HorizontalAlignment = HorizontalAlignment.Center
            };

            _writer.ApplyFormat(1, 1, 1, 1, format);

            var style = _worksheet.Cell(1, 1).Style;
            Assert.True(style.Font.Bold);
            Assert.Equal(14.0, style.Font.FontSize);
            Assert.Equal(XLAlignmentHorizontalValues.Center, style.Alignment.Horizontal);
        }

        [Fact]
        public void ApplyFormat_MultiCellRange_AppliesFormatToAllCells()
        {
            var format = new CellFormat
            {
                Italic = true,
                FontSize = 12.0f
            };

            _writer.ApplyFormat(1, 1, 3, 3, format);

            Assert.True(_worksheet.Cell(1, 1).Style.Font.Italic);
            Assert.True(_worksheet.Cell(2, 2).Style.Font.Italic);
            Assert.True(_worksheet.Cell(3, 3).Style.Font.Italic);
            Assert.Equal(12.0, _worksheet.Cell(1, 1).Style.Font.FontSize);
            Assert.Equal(12.0, _worksheet.Cell(2, 2).Style.Font.FontSize);
            Assert.Equal(12.0, _worksheet.Cell(3, 3).Style.Font.FontSize);
        }

        [Fact]
        public void ApplyFormat_WithBorders_AppliesBorderStyles()
        {
            var format = new CellFormat
            {
                LeftBorder = BorderStyle.Thin,
                RightBorder = BorderStyle.Thin,
                TopBorder = BorderStyle.Thick,
                BottomBorder = BorderStyle.Thick
            };

            _writer.ApplyFormat(2, 2, 4, 4, format);

            var border = _worksheet.Cell(2, 2).Style.Border;
            Assert.Equal(XLBorderStyleValues.Thin, border.LeftBorder);
            Assert.Equal(XLBorderStyleValues.Thin, border.RightBorder);
            Assert.Equal(XLBorderStyleValues.Thick, border.TopBorder);
            Assert.Equal(XLBorderStyleValues.Thick, border.BottomBorder);
        }

        [Fact]
        public void ApplyFormat_WithTextRotation_AppliesRotation()
        {
            var format = new CellFormat
            {
                TextRotation = 45
            };

            _writer.ApplyFormat(1, 1, 1, 1, format);

            Assert.Equal(45, _worksheet.Cell(1, 1).Style.Alignment.TextRotation);
        }

        [Fact]
        public void ApplyFormat_WithWrapText_AppliesWrapText()
        {
            var format = new CellFormat
            {
                WrapText = true
            };

            _writer.ApplyFormat(1, 1, 3, 3, format);

            Assert.True(_worksheet.Cell(1, 1).Style.Alignment.WrapText);
            Assert.True(_worksheet.Cell(2, 2).Style.Alignment.WrapText);
        }

        [Fact]
        public void ApplyFormat_WithNumberFormat_AppliesNumberFormat()
        {
            var format = new CellFormat
            {
                NumberFormat = "0.00"
            };

            _writer.ApplyFormat(1, 1, 2, 2, format);

            Assert.Equal("0.00", _worksheet.Cell(1, 1).Style.NumberFormat.Format);
        }

        [Fact]
        public void ApplyFormat_WithVerticalAlignment_AppliesAlignment()
        {
            var format = new CellFormat
            {
                VerticalAlignment = VerticalAlignment.Bottom
            };

            _writer.ApplyFormat(1, 1, 2, 2, format);

            Assert.Equal(XLAlignmentVerticalValues.Bottom, _worksheet.Cell(1, 1).Style.Alignment.Vertical);
        }

        [Theory]
        [InlineData("Calibri", 11d, "#FF0000")]
        [InlineData("Arial", 12.5d, "112233")]
        [InlineData("Consolas", 10d, "CC445566")]
        public void ApplyFormat_WithFontFormat_AppliesFontFormat(string fontName, double fontSize, string fontColorHex)
        {
            var expectedColor = new Core.ExcelColor(fontColorHex);
            var format = new CellFormat
            {
                FontName = fontName,
                FontSize = (float)fontSize,
                FontColor = expectedColor
            };

            _writer.ApplyFormat(1, 1, format);

            var style = _worksheet.Cell(1, 1).Style;
            Assert.Equal(fontName, style.Font.FontName);
            Assert.Equal(fontSize, style.Font.FontSize);
            Assert.Equal(expectedColor.Argb, style.Font.FontColor.Color.ToArgb().ToString("X8"));
        }

        [Theory]
        [InlineData("FF55")]
        [InlineData("GGHHII")]
        [InlineData("#12345")]
        [InlineData("#123456789")]
        [InlineData(null)]
        [InlineData("")]
        public void ExcelColor_WithInvalidColor_ThowsException(string? invalidColor)
        {
            if (string.IsNullOrEmpty(invalidColor))
            {
                Assert.Throws<ArgumentNullException>(() => new Core.ExcelColor(invalidColor!));
            }
            else
            {
                Assert.Throws<ArgumentException>(() => new Core.ExcelColor(invalidColor));
            }
        }

        [Fact]
        public void WriteCell_SingleCellRange_SetsValue()
        {
            _writer.WriteCell(5, 3, 5, 3, 42);

            Assert.Equal(42, _worksheet.Cell(5, 3).GetValue<int>());
        }

        [Fact]
        public void WriteCell_NullValue_SetsNullOnAllCells()
        {
            _writer.WriteCell(1, 1, 2, 2, null);

            Assert.True(_worksheet.Cell(1, 1).IsEmpty());
            Assert.True(_worksheet.Cell(1, 2).IsEmpty());
            Assert.True(_worksheet.Cell(2, 1).IsEmpty());
            Assert.True(_worksheet.Cell(2, 2).IsEmpty());
        }

        [Fact]
        public void WriteCell_HorizontalRange_SetsValueOnAllCells()
        {
            _writer.WriteCell(3, 1, 3, 5, "Horizontal");

            for (int col = 1; col <= 5; col++)
            {
                Assert.Equal("Horizontal", _worksheet.Cell(3, col).GetValue<string>());
            }
        }

        [Fact]
        public void WriteCell_VerticalRange_SetsValueOnAllCells()
        {
            _writer.WriteCell(1, 2, 4, 2, "Vertical");

            for (int row = 1; row <= 4; row++)
            {
                Assert.Equal("Vertical", _worksheet.Cell(row, 2).GetValue<string>());
            }
        }

        [Fact]
        public void CreateNamedStyle_NullStyleName_ThrowsArgumentNullException()
        {
            var exception = Assert.Throws<ArgumentNullException>(() => _writer.CreateNamedStyle(null!, new CellFormat()));
            Assert.Equal("styleName", exception.ParamName);
        }

        [Fact]
        public void CreateNamedStyle_NullFormat_ThrowsArgumentNullException()
        {
            var exception = Assert.Throws<ArgumentNullException>(() => _writer.CreateNamedStyle("Style1", null!));
            Assert.Equal("format", exception.ParamName);
        }

        [Fact]
        public void CreateNamedStyle_ValidInputs_CreatesNamedStyle()
        {
            _writer.CreateNamedStyle("ValidStyle", new CellFormat { Bold = true });

            Assert.True(_writer.NamedStyleExists("ValidStyle"));
        }

        [Fact]
        public void ApplyNamedStyle_UnknownStyle_ThrowsKeyNotFoundException()
        {
            Assert.Throws<KeyNotFoundException>(() => _writer.ApplyNamedStyle(1, 1, 1, 1, "Missing"));
        }

        [Fact]
        public void ApplyNamedStyle_ValidStyleName_AppliesStyleToRange()
        {
            const string styleName = "MyStyle";
            _writer.CreateNamedStyle(styleName, new CellFormat { Bold = true });

            _writer.ApplyNamedStyle(1, 1, 2, 2, styleName);

            Assert.True(_worksheet.Cell(1, 1).Style.Font.Bold);
            Assert.True(_worksheet.Cell(2, 2).Style.Font.Bold);
        }
    }
}
