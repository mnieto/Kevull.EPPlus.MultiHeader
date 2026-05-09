using Kevull.MultiHeader.Core;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.XmlAccess;
using System;
using System.Collections.Generic;
using System.Linq;


namespace Kevull.MultiHeader.EPPLus.Test
{
    /// <summary>
    /// Unit tests for EPPlusExcelWriter class
    /// </summary>
    public partial class EPPlusExcelWriterTests : BaseTest, IDisposable
    {
        private readonly ExcelPackage _package;
        private readonly ExcelWorksheet _worksheet;
        private readonly EPPlusExcelWriter _writer;

        public EPPlusExcelWriterTests()
        {
            _package = new ExcelPackage();
            _worksheet = _package.Workbook.Worksheets.Add("TestSheet");
            _writer = new EPPlusExcelWriter(_package, _worksheet);
        }

        public void Dispose()
        {
            _package?.Dispose();
        }

        /// <summary>
                 /// Tests that ApplyNativeFormat with null format parameter returns without error
                 /// Input: row=1, col=1, format=null
                 /// Expected: Method returns without throwing exception
                 /// </summary>
        [Fact]
        public void ApplyNativeFormat_NullFormat_ReturnsWithoutError()
        {
            // Arrange
            var xls = new ExcelPackage();
            xls.Workbook.Worksheets.Add("NullFormatTest");
            var sheet = xls.Workbook.Worksheets["NullFormatTest"];
            var writer = new EPPlusExcelWriter(xls, sheet);

            // Act & Assert - should not throw
            writer.ApplyNativeFormat(1, 1, null!);
        }

        /// <summary>
        /// Tests that ApplyNativeFormat allows the action to modify the cell
        /// Input: row=2, col=3, action that sets cell value to "TestValue"
        /// Expected: Cell C2 contains "TestValue" after action execution
        /// </summary>
        [Fact]
        public void ApplyNativeFormat_ActionModifiesCell_CellIsModified()
        {
            // Arrange
            var xls = new ExcelPackage();
            xls.Workbook.Worksheets.Add("ModifyTest");
            var sheet = xls.Workbook.Worksheets["ModifyTest"];
            var writer = new EPPlusExcelWriter(xls, sheet);

            // Act
            writer.ApplyNativeFormat(2, 3, range =>
            {
                var excelRange = (ExcelRange)range;
                excelRange.Value = "TestValue";
            });

            // Assert
            Assert.Equal("TestValue", sheet.Cells[2, 3].Value);
        }

        /// <summary>
        /// Tests that ApplyNativeFormat passes ExcelRange as object type to the action
        /// Input: row=1, col=1, action expecting object parameter
        /// Expected: Action receives an object that is an ExcelRange instance
        /// </summary>
        [Fact]
        public void ApplyNativeFormat_ActionParameterType_IsObject()
        {
            // Arrange
            var xls = new ExcelPackage();
            xls.Workbook.Worksheets.Add("TypeTest");
            var sheet = xls.Workbook.Worksheets["TypeTest"];
            var writer = new EPPlusExcelWriter(xls, sheet);
            Type? capturedType = null;

            // Act
            writer.ApplyNativeFormat(1, 1, range =>
            {
                capturedType = range?.GetType();
            });

            // Assert
            Assert.NotNull(capturedType);
            Assert.Equal(typeof(ExcelRange), capturedType);
        }

        /// <summary>
        /// Tests that ApplyFormat with null format parameter returns early without throwing exceptions.
        /// Input: null format
        /// Expected: Method returns without error, no formatting applied
        /// </summary>
        [Fact]
        public void ApplyFormat_NullFormat_DoesNothing()
        {
            // Arrange
            var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("TestSheet");
            var writer = new EPPlusExcelWriter(package, sheet);

            // Act
            writer.ApplyFormat(1, 1, 2, 2, null!);

            // Assert
            // No exception should be thrown, and the cells should remain unformatted
            Assert.False(sheet.Cells[1, 1].Style.Font.Bold);
        }

        /// <summary>
        /// Tests that ApplyFormat applies the format to a single cell.
        /// Input: Single cell coordinates with valid format
        /// Expected: Format is applied to the specified cell
        /// </summary>
        [Fact]
        public void ApplyFormat_SingleCell_AppliesFormat()
        {
            // Arrange
            var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("TestSheet");
            var writer = new EPPlusExcelWriter(package, sheet);
            var format = new CellFormat
            {
                Bold = true,
                FontSize = 14.0f,
                HorizontalAlignment = HorizontalAlignment.Center
            };

            // Act
            writer.ApplyFormat(1, 1, 1, 1, format);

            // Assert
            Assert.True(sheet.Cells[1, 1].Style.Font.Bold);
            Assert.Equal(14.0f, sheet.Cells[1, 1].Style.Font.Size);
            Assert.Equal(ExcelHorizontalAlignment.Center, sheet.Cells[1, 1].Style.HorizontalAlignment);
        }

        /// <summary>
        /// Tests that ApplyFormat applies the format to a range of cells.
        /// Input: Multi-cell range coordinates with valid format
        /// Expected: Format is applied to all cells in the range
        /// </summary>
        [Fact]
        public void ApplyFormat_MultiCellRange_AppliesFormatToAllCells()
        {
            // Arrange
            var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("TestSheet");
            var writer = new EPPlusExcelWriter(package, sheet);
            var format = new CellFormat
            {
                Italic = true,
                FontSize = 12.0f
            };

            // Act
            writer.ApplyFormat(1, 1, 3, 3, format);

            // Assert
            Assert.True(sheet.Cells[1, 1].Style.Font.Italic);
            Assert.True(sheet.Cells[2, 2].Style.Font.Italic);
            Assert.True(sheet.Cells[3, 3].Style.Font.Italic);
            Assert.Equal(12.0f, sheet.Cells[1, 1].Style.Font.Size);
            Assert.Equal(12.0f, sheet.Cells[2, 2].Style.Font.Size);
            Assert.Equal(12.0f, sheet.Cells[3, 3].Style.Font.Size);
        }

        /// <summary>
        /// Tests that ApplyFormat applies border styles correctly.
        /// Input: Format with border styles specified
        /// Expected: Border styles are applied to the range
        /// </summary>
        [Fact]
        public void ApplyFormat_WithBorders_AppliesBorderStyles()
        {
            // Arrange
            var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("TestSheet");
            var writer = new EPPlusExcelWriter(package, sheet);
            var format = new CellFormat
            {
                LeftBorder = BorderStyle.Thin,
                RightBorder = BorderStyle.Thin,
                TopBorder = BorderStyle.Thick,
                BottomBorder = BorderStyle.Thick
            };

            // Act
            writer.ApplyFormat(2, 2, 4, 4, format);

            // Assert
            Assert.Equal(ExcelBorderStyle.Thin, sheet.Cells[2, 2].Style.Border.Left.Style);
            Assert.Equal(ExcelBorderStyle.Thin, sheet.Cells[2, 2].Style.Border.Right.Style);
            Assert.Equal(ExcelBorderStyle.Thick, sheet.Cells[2, 2].Style.Border.Top.Style);
            Assert.Equal(ExcelBorderStyle.Thick, sheet.Cells[2, 2].Style.Border.Bottom.Style);
        }

        /// <summary>
        /// Tests that ApplyFormat applies text rotation correctly.
        /// Input: Format with text rotation
        /// Expected: Text rotation is applied to the range
        /// </summary>
        [Fact]
        public void ApplyFormat_WithTextRotation_AppliesRotation()
        {
            // Arrange
            var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("TestSheet");
            var writer = new EPPlusExcelWriter(package, sheet);
            var format = new CellFormat
            {
                TextRotation = 45
            };

            // Act
            writer.ApplyFormat(1, 1, 1, 1, format);

            // Assert
            Assert.Equal(45, sheet.Cells[1, 1].Style.TextRotation);
        }

        /// <summary>
        /// Tests that ApplyFormat applies wrap text setting correctly.
        /// Input: Format with wrap text enabled
        /// Expected: Wrap text is applied to the range
        /// </summary>
        [Fact]
        public void ApplyFormat_WithWrapText_AppliesWrapText()
        {
            // Arrange
            var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("TestSheet");
            var writer = new EPPlusExcelWriter(package, sheet);
            var format = new CellFormat
            {
                WrapText = true
            };

            // Act
            writer.ApplyFormat(1, 1, 3, 3, format);

            // Assert
            Assert.True(sheet.Cells[1, 1].Style.WrapText);
            Assert.True(sheet.Cells[2, 2].Style.WrapText);
        }

        /// <summary>
        /// Tests that ApplyFormat applies number format correctly.
        /// Input: Format with number format string
        /// Expected: Number format is applied to the range
        /// </summary>
        [Fact]
        public void ApplyFormat_WithNumberFormat_AppliesNumberFormat()
        {
            // Arrange
            var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("TestSheet");
            var writer = new EPPlusExcelWriter(package, sheet);
            var format = new CellFormat
            {
                NumberFormat = "0.00"
            };

            // Act
            writer.ApplyFormat(1, 1, 2, 2, format);

            // Assert
            Assert.Equal("0.00", sheet.Cells[1, 1].Style.Numberformat.Format);
        }

        /// <summary>
        /// Tests that ApplyFormat applies vertical alignment correctly.
        /// Input: Format with vertical alignment
        /// Expected: Vertical alignment is applied to the range
        /// </summary>
        [Fact]
        public void ApplyFormat_WithVerticalAlignment_AppliesAlignment()
        {
            // Arrange
            var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("TestSheet");
            var writer = new EPPlusExcelWriter(package, sheet);
            var format = new CellFormat
            {
                VerticalAlignment = VerticalAlignment.Bottom
            };

            // Act
            writer.ApplyFormat(1, 1, 2, 2, format);

            // Assert
            Assert.Equal(ExcelVerticalAlignment.Bottom, sheet.Cells[1, 1].Style.VerticalAlignment);
        }

        /// <summary>
        /// Tests that ApplyFormat with empty format object does not throw exceptions.
        /// Input: Empty CellFormat object with no properties set
        /// Expected: Method executes without error
        /// </summary>
        [Fact]
        public void ApplyFormat_EmptyFormat_DoesNotThrow()
        {
            // Arrange
            var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("TestSheet");
            var writer = new EPPlusExcelWriter(package, sheet);
            var format = new CellFormat();

            // Act & Assert
            writer.ApplyFormat(1, 1, 2, 2, format);
            // No exception expected
        }

        /// <summary>
        /// Tests that ApplyFormat handles null format parameter gracefully by delegating
        /// to the range-based overload which returns early without applying any format.
        /// </summary>
        [Fact]
        public void ApplyFormat_NullFormat_HandlesGracefully()
        {
            // Arrange
            var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add("TestSheet");
            var writer = new EPPlusExcelWriter(package, worksheet);
            int row = 1;
            int col = 1;

            // Act
            writer.ApplyFormat(row, col, null!);

            // Assert
            // When format is null, the overload returns early, so no exception should be thrown
            // and no format should be applied (verify default state)
            var cell = worksheet.Cells[row, col];
            Assert.False(cell.Style.Font.Bold);

            package.Dispose();
        }

        /// <summary>
        /// Tests that ApplyFormat correctly applies various format properties by delegating
        /// to the range-based overload for a single cell.
        /// </summary>
        [Fact]
        public void ApplyFormat_ComplexFormat_AppliesAllProperties()
        {
            // Arrange
            var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add("TestSheet");
            var writer = new EPPlusExcelWriter(package, worksheet);
            int row = 2;
            int col = 3;
            var format = new CellFormat
            {
                Bold = true,
                Italic = true,
                FontSize = 14,
                WrapText = true,
                TextRotation = 45,
                FontColor = Core.ExcelColor.Red
            };

            // Act
            writer.ApplyFormat(row, col, format);

            // Assert
            var cell = worksheet.Cells[row, col];
            Assert.True(cell.Style.Font.Bold);
            Assert.True(cell.Style.Font.Italic);
            Assert.Equal(14, cell.Style.Font.Size);
            Assert.True(cell.Style.WrapText);
            Assert.Equal(45, cell.Style.TextRotation);
            Assert.Equal(cell.Style.Font.Color.Rgb, Core.ExcelColor.Red.Argb);

            package.Dispose();
        }

        /// <summary>
        /// Tests that ApplyFormat applies font family, size and color combinations correctly.
        /// </summary>
        [Theory]
        [InlineData("Calibri", 11d, "#FF0000")]
        [InlineData("Arial", 12.5d, "112233")]
        [InlineData("Consolas", 10d, "CC445566")]
        public void ApplyFormat_WithFontFormat_AppliesFontFormat(string fontName, double fontSize, string fontColorHex)
        {
            // Arrange
            using var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add("TestSheet");
            var writer = new EPPlusExcelWriter(package, worksheet);
            var expectedColor = new Core.ExcelColor(fontColorHex);
            var format = new CellFormat
            {
                FontName = fontName,
                FontSize = (float)fontSize,
                FontColor = expectedColor
            };

            // Act
            writer.ApplyFormat(1, 1, format);

            // Assert
            var cell = worksheet.Cells[1, 1];
            Assert.Equal(fontName, cell.Style.Font.Name);
            Assert.Equal((float)fontSize, cell.Style.Font.Size);
            Assert.Equal(expectedColor.Argb, cell.Style.Font.Color.Rgb);
        }

        /// <summary>
        /// Tests that ApplyFormat with boundary row and column values delegates correctly
        /// without throwing exceptions for typical Excel ranges.
        /// </summary>
        [Theory]
        [InlineData(1, 1)]
        [InlineData(1, 16384)]
        [InlineData(1048576, 1)]
        public void ApplyFormat_TypicalExcelBoundaries_AppliesFormat(int row, int col)
        {
            // Arrange
            var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add("TestSheet");
            var writer = new EPPlusExcelWriter(package, worksheet);
            var format = new CellFormat
            {
                Bold = true
            };

            // Act
            writer.ApplyFormat(row, col, format);

            // Assert
            var cell = worksheet.Cells[row, col];
            Assert.True(cell.Style.Font.Bold);

            package.Dispose();
        }

        [Theory]
        [InlineData("FF55")] //Invalid hex color string (too short)
        [InlineData("GGHHII")] //Invalid hex color string (non-hex characters)
        [InlineData("#12345")] //Invalid hex color string (too short with #)
        [InlineData("#123456789")] //Invalid hex color string (too long with #)
        [InlineData(null)] //Null hex color string
        [InlineData("")] //Empty hex color string
        public void ExcelColor_WithInvalidColor_ThowsException(string? invalidColor)
        {
            // Arrange & Act & Assert
            if (string.IsNullOrEmpty(invalidColor))
            {
                Assert.Throws<ArgumentNullException>(() => new Core.ExcelColor(invalidColor!));
            }
            else
            {
                Assert.Throws<ArgumentException>(() => new Core.ExcelColor(invalidColor));
            }
        }

        /// <summary>
        /// Tests that AutoFitColumn executes successfully with a valid column index.
        /// Input: Valid column index (1)
        /// Expected: Method executes without throwing an exception
        /// </summary>
        [Fact]
        public void AutoFitColumn_ValidColumnIndex_DoesNotThrow()
        {
            // Arrange
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("TestSheet");
            var writer = new EPPlusExcelWriter(package, sheet);
            sheet.Cells[1, 1].Value = "Test Data";

            // Act & Assert
            var exception = Record.Exception(() => writer.AutoFitColumn(1));
            Assert.Null(exception);
        }

        /// <summary>
        /// Tests that AutoFitColumn handles zero column index.
        /// Input: Column index of 0
        /// Expected: Throws ArgumentException
        /// </summary>
        [Fact]
        public void AutoFitColumn_ZeroColumnIndex_ThrowsException()
        {
            // Arrange
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("TestSheet");
            var writer = new EPPlusExcelWriter(package, sheet);

            // Act & Assert
            Assert.Throws<ArgumentException>(() => writer.AutoFitColumn(0));
        }

        /// <summary>
        /// Tests that AutoFitColumn works correctly on empty column.
        /// Input: Valid column index with no data
        /// Expected: Method executes without throwing (AutoFit on empty column is valid)
        /// </summary>
        [Fact]
        public void AutoFitColumn_EmptyColumn_DoesNotThrow()
        {
            // Arrange
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("TestSheet");
            var writer = new EPPlusExcelWriter(package, sheet);

            // Act & Assert
            var exception = Record.Exception(() => writer.AutoFitColumn(1));
            Assert.Null(exception);
        }

        /// <summary>
        /// Tests that ApplyNamedStyle throws ArgumentNullException when styleName is null.
        /// </summary>
        [Fact]
        public void ApplyNamedStyle_NullStyleName_ThrowsArgumentNullException()
        {
            // Arrange
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("TestSheet");
            var writer = new EPPlusExcelWriter(package, sheet);
            const int row = 1;
            const int col = 1;

            // Act & Assert
            var exception = Assert.Throws<ArgumentNullException>(() => writer.ApplyNamedStyle(row, col, null!));
            Assert.Equal("styleName", exception.ParamName);
        }

        /// <summary>
        /// Tests that ApplyNamedStyle throws ArgumentNullException when styleName is empty.
        /// </summary>
        [Fact]
        public void ApplyNamedStyle_EmptyStyleName_ThrowsArgumentNullException()
        {
            // Arrange
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("TestSheet");
            var writer = new EPPlusExcelWriter(package, sheet);
            const int row = 1;
            const int col = 1;

            // Act & Assert
            var exception = Assert.Throws<ArgumentNullException>(() => writer.ApplyNamedStyle(row, col, string.Empty));
            Assert.Equal("styleName", exception.ParamName);
        }

        /// <summary>
        /// Tests that ApplyNativeFormat allows the format action to modify the ExcelRange
        /// properties such as style.
        /// </summary>
        [Fact]
        public void ApplyNativeFormat_ValidFormat_ActionCanModifyRange()
        {
            // Arrange
            var xls = new ExcelPackage();
            xls.Workbook.Worksheets.Add("TestSheet");
            var sheet = xls.Workbook.Worksheets["TestSheet"];
            var writer = new EPPlusExcelWriter(xls, sheet);

            Action<object> formatAction = (range) =>
            {
                var excelRange = (ExcelRange)range;
                excelRange.Style.Font.Bold = true;
            };

            // Act
            writer.ApplyNativeFormat(1, 1, 2, 2, formatAction);

            // Assert
            Assert.True(sheet.Cells[1, 1].Style.Font.Bold);
            Assert.True(sheet.Cells[2, 2].Style.Font.Bold);
        }

        /// <summary>
        /// Tests that NamedStyleExists returns false when the named style does not exist in the workbook.
        /// This test validates the search logic when the style is not present in the collection.
        /// </summary>
        [Fact]
        public void NamedStyleExists_StyleDoesNotExist_ReturnsFalse()
        {
            // Arrange
            var package = new ExcelPackage();
            package.Workbook.Worksheets.Add("TestSheet");
            var sheet = package.Workbook.Worksheets["TestSheet"];
            var writer = new EPPlusExcelWriter(package, sheet);

            // Act
            var result = writer.NamedStyleExists("NonExistentStyle");

            // Assert
            Assert.False(result);
        }

        /// <summary>
        /// Tests that NamedStyleExists returns true when the named style exists in the workbook.
        /// This test validates the search logic when the style is present in the collection.
        /// </summary>
        [Fact]
        public void NamedStyleExists_StyleExists_ReturnsTrue()
        {
            // Arrange
            var package = new ExcelPackage();
            package.Workbook.Worksheets.Add("TestSheet");
            var sheet = package.Workbook.Worksheets["TestSheet"];
            var writer = new EPPlusExcelWriter(package, sheet);
            package.Workbook.Styles.CreateNamedStyle("TestStyle");

            // Act
            var result = writer.NamedStyleExists("TestStyle");

            // Assert
            Assert.True(result);
        }

        /// <summary>
        /// Tests that NamedStyleExists returns false when the style name does not match exactly (case sensitivity).
        /// This test validates that the search is case-sensitive.
        /// </summary>
        [Fact]
        public void NamedStyleExists_StyleExistsWithDifferentCase_ReturnsFalse()
        {
            // Arrange
            var package = new ExcelPackage();
            package.Workbook.Worksheets.Add("TestSheet");
            var sheet = package.Workbook.Worksheets["TestSheet"];
            var writer = new EPPlusExcelWriter(package, sheet);
            package.Workbook.Styles.CreateNamedStyle("TestStyle");

            // Act
            var result = writer.NamedStyleExists("teststyle");

            // Assert
            Assert.False(result);
        }

        /// <summary>
        /// Tests that NamedStyleExists works correctly with multiple named styles in the collection.
        /// This test validates the search logic when multiple styles exist.
        /// </summary>
        [Fact]
        public void NamedStyleExists_MultipleStyles_ReturnsCorrectResult()
        {
            // Arrange
            var package = new ExcelPackage();
            package.Workbook.Worksheets.Add("TestSheet");
            var sheet = package.Workbook.Worksheets["TestSheet"];
            var writer = new EPPlusExcelWriter(package, sheet);
            package.Workbook.Styles.CreateNamedStyle("Style1");
            package.Workbook.Styles.CreateNamedStyle("Style2");
            package.Workbook.Styles.CreateNamedStyle("Style3");

            // Act
            var result1 = writer.NamedStyleExists("Style1");
            var result2 = writer.NamedStyleExists("Style2");
            var result3 = writer.NamedStyleExists("Style3");
            var result4 = writer.NamedStyleExists("Style4");

            // Assert
            Assert.True(result1);
            Assert.True(result2);
            Assert.True(result3);
            Assert.False(result4);
        }

        /// <summary>
        /// Tests that NamedStyleExists handles style names with special characters correctly.
        /// This test validates behavior with non-alphanumeric style names.
        /// </summary>
        [Theory]
        [InlineData("Style-1")]
        [InlineData("Style_2")]
        [InlineData("Style.3")]
        [InlineData("Style@4")]
        [InlineData("Style#5")]
        public void NamedStyleExists_StyleNameWithSpecialCharacters_WorksCorrectly(string styleName)
        {
            // Arrange
            var package = new ExcelPackage();
            package.Workbook.Worksheets.Add("TestSheet");
            var sheet = package.Workbook.Worksheets["TestSheet"];
            var writer = new EPPlusExcelWriter(package, sheet);
            package.Workbook.Styles.CreateNamedStyle(styleName);

            // Act
            var result = writer.NamedStyleExists(styleName);

            // Assert
            Assert.True(result);
        }

        /// <summary>
        /// Tests that WriteCell with a single cell range sets the value correctly.
        /// Input: A single cell range (fromRow == toRow, fromCol == toCol) with an integer value.
        /// Expected: The single cell contains the specified value.
        /// </summary>
        [Fact]
        public void WriteCell_SingleCellRange_SetsValue()
        {
            // Arrange
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("TestSheet");
            var writer = new EPPlusExcelWriter(package, sheet);
            var value = 42;
            int row = 5, col = 3;

            // Act
            writer.WriteCell(row, col, row, col, value);

            // Assert
            Assert.Equal(value, sheet.GetValue<int>(5, 3));
        }

        /// <summary>
        /// Tests that WriteCell with a null value sets null on all cells in the range.
        /// Input: A multi-cell range with null value.
        /// Expected: All cells in the range have null values.
        /// </summary>
        [Fact]
        public void WriteCell_NullValue_SetsNullOnAllCells()
        {
            // Arrange
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("TestSheet");
            var writer = new EPPlusExcelWriter(package, sheet);
            int fromRow = 1, fromCol = 1, toRow = 2, toCol = 2;

            // Act
            writer.WriteCell(fromRow, fromCol, toRow, toCol, null);

            // Assert
            Assert.Null(sheet.GetValue(1, 1));
            Assert.Null(sheet.GetValue(1, 2));
            Assert.Null(sheet.GetValue(2, 1));
            Assert.Null(sheet.GetValue(2, 2));
        }

        /// <summary>
        /// Tests that WriteCell correctly handles various value types.
        /// Input: Different value types (string, int, double, DateTime, bool).
        /// Expected: Each value type is correctly set on the range.
        /// </summary>
        [Theory]
        [InlineData("StringValue")]
        [InlineData(123)]
        [InlineData(45.67)]
        [InlineData(true)]
        [InlineData(0)]
        [InlineData(-999)]
        public void WriteCell_VariousValueTypes_SetsCorrectly(object value)
        {
            // Arrange
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("TestSheet");
            var writer = new EPPlusExcelWriter(package, sheet);
            int fromRow = 1, fromCol = 1, toRow = 1, toCol = 1;

            // Act
            writer.WriteCell(fromRow, fromCol, toRow, toCol, value);

            // Assert
            Assert.Equal(value, sheet.GetValue(1, 1));
        }

        /// <summary>
        /// Tests that WriteCell with a horizontal range sets the value on all cells in the row.
        /// Input: A horizontal range spanning multiple columns in a single row.
        /// Expected: All cells in the horizontal range contain the specified value.
        /// </summary>
        [Fact]
        public void WriteCell_HorizontalRange_SetsValueOnAllCells()
        {
            // Arrange
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("TestSheet");
            var writer = new EPPlusExcelWriter(package, sheet);
            var value = "Horizontal";
            int row = 3, fromCol = 1, toCol = 5;

            // Act
            writer.WriteCell(row, fromCol, row, toCol, value);

            // Assert
            for (int col = fromCol; col <= toCol; col++)
            {
                Assert.Equal(value, sheet.GetValue<string>(row, col));
            }
        }

        /// <summary>
        /// Tests that WriteCell with a vertical range sets the value on all cells in the column.
        /// Input: A vertical range spanning multiple rows in a single column.
        /// Expected: All cells in the vertical range contain the specified value.
        /// </summary>
        [Fact]
        public void WriteCell_VerticalRange_SetsValueOnAllCells()
        {
            // Arrange
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("TestSheet");
            var writer = new EPPlusExcelWriter(package, sheet);
            var value = "Vertical";
            int col = 2, fromRow = 1, toRow = 4;

            // Act
            writer.WriteCell(fromRow, col, toRow, col, value);

            // Assert
            for (int row = fromRow; row <= toRow; row++)
            {
                Assert.Equal(value, sheet.GetValue<string>(row, col));
            }
        }

        /// <summary>
        /// Tests that WriteCell with an empty string sets empty string on all cells.
        /// Input: A range with an empty string value.
        /// Expected: All cells in the range contain an empty string.
        /// </summary>
        [Fact]
        public void WriteCell_EmptyString_SetsEmptyStringOnAllCells()
        {
            // Arrange
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("TestSheet");
            var writer = new EPPlusExcelWriter(package, sheet);
            var value = string.Empty;
            int fromRow = 1, fromCol = 1, toRow = 2, toCol = 2;

            // Act
            writer.WriteCell(fromRow, fromCol, toRow, toCol, value);

            // Assert
            Assert.Equal(value, sheet.GetValue<string>(1, 1));
            Assert.Equal(value, sheet.GetValue<string>(1, 2));
            Assert.Equal(value, sheet.GetValue<string>(2, 1));
            Assert.Equal(value, sheet.GetValue<string>(2, 2));
        }

        /// <summary>
        /// Tests that WriteCell with special numeric values sets them correctly.
        /// Input: Special numeric values including zero, negative values, and decimal values.
        /// Expected: Each special value is correctly set on the cells.
        /// </summary>
        [Theory]
        [InlineData(0.0)]
        [InlineData(-1.5)]
        [InlineData(double.MaxValue)]
        [InlineData(double.MinValue)]
        public void WriteCell_SpecialNumericValues_SetsCorrectly(double value)
        {
            // Arrange
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("TestSheet");
            var writer = new EPPlusExcelWriter(package, sheet);
            int fromRow = 1, fromCol = 1, toRow = 1, toCol = 1;

            // Act
            writer.WriteCell(fromRow, fromCol, toRow, toCol, value);

            // Assert
            Assert.Equal(value, sheet.GetValue<double>(1, 1));
        }

        /// <summary>
        /// Tests that CreateNamedStyle throws ArgumentNullException when styleName is null.
        /// Input: null styleName, valid format.
        /// Expected: ArgumentNullException with parameter name "styleName".
        /// </summary>
        [Fact]
        public void CreateNamedStyle_NullStyleName_ThrowsArgumentNullException()
        {
            // Arrange
            string? styleName = null;
            CellFormat format = new CellFormat();

            // Act & Assert
            ArgumentNullException exception = Assert.Throws<ArgumentNullException>(() => _writer.CreateNamedStyle(styleName!, format));
            Assert.Equal("styleName", exception.ParamName);
        }

        /// <summary>
        /// Tests that CreateNamedStyle throws ArgumentNullException when styleName is empty.
        /// Input: empty string styleName, valid format.
        /// Expected: ArgumentNullException with parameter name "styleName".
        /// </summary>
        [Fact]
        public void CreateNamedStyle_EmptyStyleName_ThrowsArgumentNullException()
        {
            // Arrange
            string styleName = string.Empty;
            CellFormat format = new CellFormat();

            // Act & Assert
            ArgumentNullException exception = Assert.Throws<ArgumentNullException>(() => _writer.CreateNamedStyle(styleName, format));
            Assert.Equal("styleName", exception.ParamName);
        }

        /// <summary>
        /// Tests that CreateNamedStyle throws ArgumentNullException when format is null.
        /// Input: valid styleName, null format.
        /// Expected: ArgumentNullException with parameter name "format".
        /// </summary>
        [Fact]
        public void CreateNamedStyle_NullFormat_ThrowsArgumentNullException()
        {
            // Arrange
            string styleName = "TestStyle";
            CellFormat? format = null;

            // Act & Assert
            ArgumentNullException exception = Assert.Throws<ArgumentNullException>(() => _writer.CreateNamedStyle(styleName, format!));
            Assert.Equal("format", exception.ParamName);
        }

        /// <summary>
        /// Tests that CreateNamedStyle returns early without creating a style when the style already exists.
        /// Input: valid styleName that already exists, valid format.
        /// Expected: Method returns without error, no new style is created.
        /// </summary>
        [Fact]
        public void CreateNamedStyle_StyleAlreadyExists_ReturnsEarly()
        {
            // Arrange
            string styleName = "ExistingStyle";
            CellFormat format = new CellFormat { Bold = true };

            // Create the style first
            _writer.CreateNamedStyle(styleName, format);
            int initialStyleCount = _package.Workbook.Styles.NamedStyles.Count();

            // Act - try to create the same style again
            _writer.CreateNamedStyle(styleName, format);

            // Assert - style count should remain the same
            int finalStyleCount = _package.Workbook.Styles.NamedStyles.Count();
            Assert.Equal(initialStyleCount, finalStyleCount);
        }

        /// <summary>
        /// Tests that CreateNamedStyle successfully creates a named style with valid inputs.
        /// Input: valid styleName, valid format.
        /// Expected: Named style is created and exists in the workbook.
        /// </summary>
        [Fact]
        public void CreateNamedStyle_ValidInputs_CreatesNamedStyle()
        {
            // Arrange
            string styleName = "ValidStyle";
            CellFormat format = new CellFormat { Bold = true };

            // Act
            _writer.CreateNamedStyle(styleName, format);

            // Assert
            Assert.True(_writer.NamedStyleExists(styleName));
            ExcelNamedStyleXml? createdStyle = _package.Workbook.Styles.NamedStyles.FirstOrDefault(s => s.Name == styleName);
            Assert.NotNull(createdStyle);
        }

        /// <summary>
        /// Tests that CreateNamedStyle handles empty CellFormat (no formatting properties set).
        /// Input: valid styleName, empty CellFormat.
        /// Expected: Named style is created without errors.
        /// </summary>
        [Fact]
        public void CreateNamedStyle_EmptyFormat_CreatesStyleWithoutErrors()
        {
            // Arrange
            string styleName = "EmptyFormatStyle";
            CellFormat format = new CellFormat();

            // Act
            _writer.CreateNamedStyle(styleName, format);

            // Assert
            Assert.True(_writer.NamedStyleExists(styleName));
        }

        /// <summary>
        /// Tests that ApplyNamedStyle correctly sets the StyleName property on a cell range.
        /// This test validates the normal operation with a valid style name.
        /// Expected: The StyleName property is set to the provided value
        /// </summary>
        [Fact]
        public void ApplyNamedStyle_ValidStyleName_SetsStyleNameOnRange()
        {
            // Arrange
            using var package = new ExcelPackage();
            package.Workbook.Worksheets.Add("Test");
            var sheet = package.Workbook.Worksheets["Test"];
            var writer = new EPPlusExcelWriter(package, sheet);
            string expectedStyleName = "MyStyle";
            writer.CreateNamedStyle(expectedStyleName, new CellFormat());

            // Act
            writer.ApplyNamedStyle(1, 1, 2, 2, expectedStyleName);

            // Assert
            Assert.Equal(expectedStyleName, sheet.Cells[1, 1, 2, 2].StyleName);
        }


        /// <summary>
        /// Tests that ApplyNamedStyle handles style names with special characters correctly.
        /// This test validates that various valid string values are correctly assigned.
        /// Expected: The StyleName is set to the exact value provided, including special characters
        /// </summary>
        [Theory]
        [InlineData("Style1")]
        [InlineData("My-Style")]
        [InlineData("Style_123")]
        [InlineData("StyleWithNumbers123")]
        [InlineData("Style.Name")]
        public void ApplyNamedStyle_SpecialCharactersInStyleName_SetsStyleNameCorrectly(string styleName)
        {
            // Arrange
            using var package = new ExcelPackage();
            package.Workbook.Worksheets.Add("Test");
            var sheet = package.Workbook.Worksheets["Test"];
            var writer = new EPPlusExcelWriter(package, sheet);
            writer.CreateNamedStyle(styleName, new CellFormat());

            // Act
            writer.ApplyNamedStyle(1, 1, 2, 2, styleName);

            // Assert
            Assert.Equal(styleName, sheet.Cells[1, 1, 2, 2].StyleName);
        }
    }
}