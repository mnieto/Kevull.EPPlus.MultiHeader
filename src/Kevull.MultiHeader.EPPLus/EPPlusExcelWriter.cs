using Kevull.MultiHeader.Core;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Drawing;
using System.Linq;
using CoreExcelColor = Kevull.MultiHeader.Core.ExcelColor;

namespace Kevull.MultiHeader.EPPLus
{
    /// <summary>
    /// EPPlus implementation of <see cref="IExcelWriter"/>
    /// </summary>
    public class EPPlusExcelWriter : IExcelWriter
    {
        private readonly ExcelWorksheet _sheet;
        private readonly ExcelPackage _package;

        /// <summary>
        /// Creates a new instance of EPPlusExcelWriter
        /// </summary>
        /// <param name="package">The ExcelPackage to work with</param>
        /// <param name="sheet">The ExcelWorksheet to work with</param>
        public EPPlusExcelWriter(ExcelPackage package, ExcelWorksheet sheet)
        {
            _package = package ?? throw new ArgumentNullException(nameof(package));
            _sheet = sheet ?? throw new ArgumentNullException(nameof(sheet));
        }

        #region Basic Writing

        /// <inheritdoc/>
        public void WriteCell(int row, int col, object value)
        {
            _sheet.Cells[row, col].Value = value;
        }

        /// <inheritdoc/>
        public void WriteCell(int fromRow, int fromCol, int toRow, int toCol, object value)
        {
            _sheet.Cells[fromRow, fromCol, toRow, toCol].Value = value;
        }

        /// <inheritdoc/>
        public void WriteFormula(int row, int col, string formula)
        {
            _sheet.Cells[row, col].Formula = formula;
        }

        /// <inheritdoc/>
        public void WriteFormula(int fromRow, int fromCol, int toRow, int toCol, string formula)
        {
            _sheet.Cells[fromRow, fromCol, toRow, toCol].Formula = formula;
        }

        #endregion

        #region Formatting

        /// <inheritdoc/>
        public void ApplyFormat(int row, int col, CellFormat format)
        {
            ApplyFormat(row, col, row, col, format);
        }

        /// <inheritdoc/>
        public void ApplyFormat(int fromRow, int fromCol, int toRow, int toCol, CellFormat format)
        {
            if (format == null) return;

            var range = _sheet.Cells[fromRow, fromCol, toRow, toCol];
            ApplyFormatToRange(range.Style, format);
        }

        /// <inheritdoc/>
        public void ApplyNativeFormat(int row, int col, Action<object> format)
        {
            if (format == null) return;
            var range = _sheet.Cells[row, col];
            format(range);
        }

        /// <inheritdoc/>
        public void ApplyNativeFormat(int fromRow, int fromCol, int toRow, int toCol, Action<object> format)
        {
            if (format == null) return;
            var range = _sheet.Cells[fromRow, fromCol, toRow, toCol];
            format(range);
        }

        #endregion

        #region Named Styles

        /// <inheritdoc/>
        public void CreateNamedStyle(string styleName, CellFormat format)
        {
            if (string.IsNullOrWhiteSpace(styleName))
                throw new ArgumentNullException(nameof(styleName));
            if (format == null)
                throw new ArgumentNullException(nameof(format));

            // Check if style already exists
            if (NamedStyleExists(styleName))
                return;

            var namedStyle = _package.Workbook.Styles.CreateNamedStyle(styleName);
            ApplyFormatToRange(namedStyle.Style, format);
        }

        /// <inheritdoc/>
        public void ApplyNamedStyle(int row, int col, string styleName)
        {
            ApplyNamedStyle(row, col, row, col, styleName);
        }

        /// <inheritdoc/>
        public void ApplyNamedStyle(int fromRow, int fromCol, int toRow, int toCol, string styleName)
        {
            if (string.IsNullOrWhiteSpace(styleName))
                throw new ArgumentNullException(nameof(styleName));

            var range = _sheet.Cells[fromRow, fromCol, toRow, toCol];
            range.StyleName = styleName;
        }

        /// <inheritdoc/>
        public bool NamedStyleExists(string styleName)
        {
            if (string.IsNullOrWhiteSpace(styleName))
                return false;

            return _package.Workbook.Styles.NamedStyles.Any(x => x.Name == styleName);
        }

        #endregion

        #region Cell Operations

        /// <inheritdoc/>
        public void Merge(int fromRow, int fromCol, int toRow, int toCol)
        {
            _sheet.Cells[fromRow, fromCol, toRow, toCol].Merge = true;
        }

        #endregion

        #region Column Operations

        /// <inheritdoc/>
        public void AutoFitColumn(int col)
        {
            _sheet.Column(col).AutoFit();
        }

        /// <inheritdoc/>
        public void AutoFitColumn(int col, double minWidth, double maxWidth)
        {
            _sheet.Column(col).AutoFit(minWidth, maxWidth);
        }

        /// <inheritdoc/>
        public void SetColumnWidth(int col, double width)
        {
            _sheet.Column(col).Width = width;
        }

        /// <inheritdoc/>
        public void SetColumnHidden(int col, bool hidden)
        {
            _sheet.Column(col).Hidden = hidden;
        }

        /// <inheritdoc/>
        public void SetAutoFilter(int fromRow, int fromCol, int toRow, int toCol, bool autoFilter)
        {
            _sheet.Cells[fromRow, fromCol, toRow, toCol].AutoFilter = autoFilter;
        }

        /// <inheritdoc/>
        public void FreezePanes(int row, int col)
        {
            _sheet.View.FreezePanes(row, col);
        }

        /// <inheritdoc/>
        public void Recalculate()
        {
            _sheet.Calculate();
        }

        #endregion

        #region Private Helper Methods

        /// <summary>
        /// Applies CellFormat to an ExcelStyle
        /// </summary>
        private void ApplyFormatToRange(ExcelStyle style, CellFormat format)
        {
            // Borders
            if (format.LeftBorder.HasValue)
                style.Border.Left.Style = ConvertBorderStyle(format.LeftBorder.Value);
            if (format.RightBorder.HasValue)
                style.Border.Right.Style = ConvertBorderStyle(format.RightBorder.Value);
            if (format.TopBorder.HasValue)
                style.Border.Top.Style = ConvertBorderStyle(format.TopBorder.Value);
            if (format.BottomBorder.HasValue)
                style.Border.Bottom.Style = ConvertBorderStyle(format.BottomBorder.Value);

            // Alignment
            if (format.VerticalAlignment.HasValue)
                style.VerticalAlignment = ConvertVerticalAlignment(format.VerticalAlignment.Value);
            if (format.HorizontalAlignment.HasValue)
                style.HorizontalAlignment = ConvertHorizontalAlignment(format.HorizontalAlignment.Value);

            // Fill
            if (format.BackgroundColor.HasValue && format.FillStyle.HasValue)
            {
                var color = ConvertColor(format.BackgroundColor.Value);
                var fillStyle = ConvertFillStyle(format.FillStyle.Value);
                style.Fill.SetBackground(color, fillStyle);
            }

            // Font
            if (format.Bold.HasValue)
                style.Font.Bold = format.Bold.Value;
            if (format.Italic.HasValue)
                style.Font.Italic = format.Italic.Value;
            if (format.FontSize.HasValue)
                style.Font.Size = format.FontSize.Value;
            if (format.FontColor.HasValue)
                style.Font.Color.SetColor(ConvertColor(format.FontColor.Value));

            // Number Format
            if (!string.IsNullOrWhiteSpace(format.NumberFormat))
                style.Numberformat.Format = format.NumberFormat;

            // Text Rotation
            if (format.TextRotation.HasValue)
                style.TextRotation = format.TextRotation.Value;

            // Wrap Text
            if (format.WrapText.HasValue)
                style.WrapText = format.WrapText.Value;
        }

        /// <summary>
        /// Converts library-agnostic BorderStyle to EPPlus ExcelBorderStyle
        /// </summary>
        private ExcelBorderStyle ConvertBorderStyle(Core.BorderStyle borderStyle)
        {
            return borderStyle switch
            {
                Core.BorderStyle.None => ExcelBorderStyle.None,
                Core.BorderStyle.Thin => ExcelBorderStyle.Thin,
                Core.BorderStyle.Medium => ExcelBorderStyle.Medium,
                Core.BorderStyle.Thick => ExcelBorderStyle.Thick,
                Core.BorderStyle.Double => ExcelBorderStyle.Double,
                Core.BorderStyle.Dotted => ExcelBorderStyle.Dotted,
                Core.BorderStyle.Dashed => ExcelBorderStyle.Dashed,
                Core.BorderStyle.DashDot => ExcelBorderStyle.DashDot,
                Core.BorderStyle.DashDotDot => ExcelBorderStyle.DashDotDot,
                _ => ExcelBorderStyle.None
            };
        }

        /// <summary>
        /// Converts library-agnostic VerticalAlignment to EPPlus ExcelVerticalAlignment
        /// </summary>
        private ExcelVerticalAlignment ConvertVerticalAlignment(Core.VerticalAlignment alignment)
        {
            return alignment switch
            {
                Core.VerticalAlignment.Top => ExcelVerticalAlignment.Top,
                Core.VerticalAlignment.Center => ExcelVerticalAlignment.Center,
                Core.VerticalAlignment.Bottom => ExcelVerticalAlignment.Bottom,
                Core.VerticalAlignment.Justify => ExcelVerticalAlignment.Justify,
                Core.VerticalAlignment.Distributed => ExcelVerticalAlignment.Distributed,
                _ => ExcelVerticalAlignment.Bottom
            };
        }

        /// <summary>
        /// Converts library-agnostic HorizontalAlignment to EPPlus ExcelHorizontalAlignment
        /// </summary>
        private ExcelHorizontalAlignment ConvertHorizontalAlignment(Core.HorizontalAlignment alignment)
        {
            return alignment switch
            {
                Core.HorizontalAlignment.General => ExcelHorizontalAlignment.General,
                Core.HorizontalAlignment.Left => ExcelHorizontalAlignment.Left,
                Core.HorizontalAlignment.Center => ExcelHorizontalAlignment.Center,
                Core.HorizontalAlignment.Right => ExcelHorizontalAlignment.Right,
                Core.HorizontalAlignment.Fill => ExcelHorizontalAlignment.Fill,
                Core.HorizontalAlignment.Justify => ExcelHorizontalAlignment.Justify,
                Core.HorizontalAlignment.CenterContinuous => ExcelHorizontalAlignment.CenterContinuous,
                Core.HorizontalAlignment.Distributed => ExcelHorizontalAlignment.Distributed,
                _ => ExcelHorizontalAlignment.General
            };
        }

        /// <summary>
        /// Converts library-agnostic FillStyle to EPPlus ExcelFillStyle
        /// </summary>
        private ExcelFillStyle ConvertFillStyle(Core.FillStyle fillStyle)
        {
            return fillStyle switch
            {
                Core.FillStyle.None => ExcelFillStyle.None,
                Core.FillStyle.Solid => ExcelFillStyle.Solid,
                Core.FillStyle.DarkGray => ExcelFillStyle.DarkGray,
                Core.FillStyle.MediumGray => ExcelFillStyle.MediumGray,
                Core.FillStyle.LightGray => ExcelFillStyle.LightGray,
                Core.FillStyle.Gray125 => ExcelFillStyle.Gray125,
                Core.FillStyle.Gray0625 => ExcelFillStyle.Gray0625,
                _ => ExcelFillStyle.None
            };
        }

        /// <summary>
        /// Converts library-agnostic ExcelColor to System.Drawing.Color
        /// </summary>
        private Color ConvertColor(CoreExcelColor excelColor)
        {
            return Color.FromArgb(excelColor.A, excelColor.R, excelColor.G, excelColor.B);
        }

        #endregion
    }
}
