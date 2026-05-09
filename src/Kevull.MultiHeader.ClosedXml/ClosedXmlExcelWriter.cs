using ClosedXML.Excel;
using Kevull.MultiHeader.Core;
using System;
using System.Collections.Generic;
using CoreExcelColor = Kevull.MultiHeader.Core.ExcelColor;

namespace Kevull.MultiHeader.ClosedXml
{
    /// <summary>
    /// ClosedXML implementation of <see cref="IExcelWriter"/>
    /// </summary>
    public class ClosedXmlExcelWriter : IExcelWriter
    {
        private readonly IXLWorksheet _sheet;
        private readonly XLWorkbook _workbook;
        private readonly Dictionary<string, CellFormat> _namedStyles = new Dictionary<string, CellFormat>(StringComparer.Ordinal);

        /// <summary>
        /// Creates a new instance of ClosedXmlExcelWriter.
        /// </summary>
        /// <param name="workbook">The <see cref="XLWorkbook"/> to work with.</param>
        /// <param name="sheet">The <see cref="IXLWorksheet"/> to work with.</param>
        public ClosedXmlExcelWriter(XLWorkbook workbook, IXLWorksheet sheet)
        {
            _workbook = workbook ?? throw new ArgumentNullException(nameof(workbook));
            _sheet = sheet ?? throw new ArgumentNullException(nameof(sheet));
        }

        /// <inheritdoc/>
        public void WriteCell(int row, int col, object? value)
        {
            _sheet.Cell(row, col).Value = XLCellValue.FromObject(value);
        }

        /// <inheritdoc/>
        public void WriteCell(int fromRow, int fromCol, int toRow, int toCol, object? value)
        {
            var cellValue = XLCellValue.FromObject(value);
            var range = _sheet.Range(fromRow, fromCol, toRow, toCol);
            range.SetValue(cellValue);
        }

        /// <inheritdoc/>
        public void WriteCellWithHyperlink(int row, int col, object? value, string url)
        {
            var cell = _sheet.Cell(row, col);
            cell.Value = XLCellValue.FromObject(value);
            if (!string.IsNullOrWhiteSpace(url))
            {
                cell.SetHyperlink(new XLHyperlink(url));
            }
        }

        /// <inheritdoc/>
        public void WriteFormula(int row, int col, string formula)
        {
            _sheet.Cell(row, col).FormulaA1 = formula;
        }

        /// <inheritdoc/>
        public void WriteFormula(int fromRow, int fromCol, int toRow, int toCol, string formula)
        {
            _sheet.Range(fromRow, fromCol, toRow, toCol).FormulaA1 = formula;
        }

        /// <inheritdoc/>
        public void ApplyFormat(int row, int col, CellFormat format)
        {
            ApplyFormat(row, col, row, col, format);
        }

        /// <inheritdoc/>
        public void ApplyFormat(int fromRow, int fromCol, int toRow, int toCol, CellFormat format)
        {
            if (format == null) return;

            var range = _sheet.Range(fromRow, fromCol, toRow, toCol);
            ExcelStyleExtensions.ApplyFormatToStyle(range.Style, format);
        }

        /// <inheritdoc/>
        public void ApplyNativeFormat(int row, int col, Action<object> format)
        {
            if (format == null) return;
            var range = _sheet.Range(row, col, row, col);
            format(range);
        }

        /// <inheritdoc/>
        public void ApplyNativeFormat(int fromRow, int fromCol, int toRow, int toCol, Action<object> format)
        {
            if (format == null) return;
            var range = _sheet.Range(fromRow, fromCol, toRow, toCol);
            format(range);
        }

        /// <inheritdoc/>
        public void CreateNamedStyle(string styleName, CellFormat format)
        {
            if (string.IsNullOrWhiteSpace(styleName))
                throw new ArgumentNullException(nameof(styleName));
            if (format == null)
                throw new ArgumentNullException(nameof(format));

            if (NamedStyleExists(styleName))
                return;

            _namedStyles[styleName] = format;
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

            if (!_namedStyles.TryGetValue(styleName, out var format))
                throw new KeyNotFoundException($"Style '{styleName}' does not exist.");

            var range = _sheet.Range(fromRow, fromCol, toRow, toCol);
            ExcelStyleExtensions.ApplyFormatToStyle(range.Style, format);
        }

        /// <inheritdoc/>
        public bool NamedStyleExists(string styleName)
        {
            if (string.IsNullOrWhiteSpace(styleName))
                return false;

            return _namedStyles.ContainsKey(styleName);
        }

        /// <inheritdoc/>
        public void Merge(int fromRow, int fromCol, int toRow, int toCol)
        {
            _sheet.Range(fromRow, fromCol, toRow, toCol).Merge();
        }

        /// <inheritdoc/>
        public void AutoFitColumn(int col)
        {
            _sheet.Column(col).AdjustToContents();
        }

        /// <inheritdoc/>
        public void AutoFitColumn(int col, double minWidth, double maxWidth)
        {
            _sheet.Column(col).AdjustToContents(1, _sheet.LastRowUsed()?.RowNumber() ?? 1, minWidth, maxWidth);
        }

        /// <inheritdoc/>
        public void SetColumnWidth(int col, double width)
        {
            _sheet.Column(col).Width = width;
        }

        /// <inheritdoc/>
        public void SetColumnHidden(int col, bool hidden)
        {
            var column = _sheet.Column(col);
            if (hidden)
                column.Hide();
            else
                column.Unhide();
        }

        /// <inheritdoc/>
        public void SetAutoFilter(int fromRow, int fromCol, int toRow, int toCol, bool autoFilter)
        {
            var range = _sheet.Range(fromRow, fromCol, toRow, toCol);
            if (autoFilter)
                range.SetAutoFilter();
            else if (_sheet.AutoFilter != null)
                _sheet.AutoFilter.Clear();
        }

        /// <inheritdoc/>
        public void FreezePanes(int row, int col)
        {
            _sheet.SheetView.Freeze(row - 1, col - 1);
        }

        /// <inheritdoc/>
        public void Recalculate()
        {
            _workbook.RecalculateAllFormulas();
        }
    }

    /// <summary>
    /// Extension methods to apply library-agnostic cell format definitions to ClosedXML styles.
    /// </summary>
    public static class ExcelStyleExtensions
    {
        /// <summary>
        /// Applies <see cref="CellFormat"/> to an <see cref="IXLStyle"/>.
        /// </summary>
        public static IXLStyle SetBackground(this CellFormat format, IXLStyle style)
        {
            ApplyFormatToStyle(style, format);
            return style;
        }

        /// <summary>
        /// Applies <see cref="CellFormat"/> to an <see cref="IXLStyle"/>.
        /// </summary>
        internal static void ApplyFormatToStyle(IXLStyle style, CellFormat format)
        {
            if (style == null || format == null)
                return;

            if (format.LeftBorder.HasValue)
                style.Border.LeftBorder = ConvertBorderStyle(format.LeftBorder.Value);
            if (format.RightBorder.HasValue)
                style.Border.RightBorder = ConvertBorderStyle(format.RightBorder.Value);
            if (format.TopBorder.HasValue)
                style.Border.TopBorder = ConvertBorderStyle(format.TopBorder.Value);
            if (format.BottomBorder.HasValue)
                style.Border.BottomBorder = ConvertBorderStyle(format.BottomBorder.Value);

            if (format.VerticalAlignment.HasValue)
                style.Alignment.Vertical = ConvertVerticalAlignment(format.VerticalAlignment.Value);
            if (format.HorizontalAlignment.HasValue)
                style.Alignment.Horizontal = ConvertHorizontalAlignment(format.HorizontalAlignment.Value);

            if (format.BackgroundColor.HasValue && format.FillStyle.HasValue)
            {
                style.Fill.BackgroundColor = ConvertColor(format.BackgroundColor.Value);
                style.Fill.PatternType = ConvertFillStyle(format.FillStyle.Value);
            }

            if (format.Bold.HasValue)
                style.Font.Bold = format.Bold.Value;
            if (format.Italic.HasValue)
                style.Font.Italic = format.Italic.Value;
            if (!string.IsNullOrWhiteSpace(format.FontName))
                style.Font.FontName = format.FontName;
            if (format.FontSize.HasValue)
                style.Font.FontSize = format.FontSize.Value;
            if (format.FontColor.HasValue)
                style.Font.FontColor = ConvertColor(format.FontColor.Value);

            if (!string.IsNullOrWhiteSpace(format.NumberFormat))
                style.NumberFormat.Format = format.NumberFormat;

            if (format.TextRotation.HasValue)
                style.Alignment.TextRotation = format.TextRotation.Value;

            if (format.WrapText.HasValue)
                style.Alignment.WrapText = format.WrapText.Value;
        }

        private static XLBorderStyleValues ConvertBorderStyle(Core.BorderStyle borderStyle)
        {
            return borderStyle switch
            {
                Core.BorderStyle.None => XLBorderStyleValues.None,
                Core.BorderStyle.Thin => XLBorderStyleValues.Thin,
                Core.BorderStyle.Medium => XLBorderStyleValues.Medium,
                Core.BorderStyle.Thick => XLBorderStyleValues.Thick,
                Core.BorderStyle.Double => XLBorderStyleValues.Double,
                Core.BorderStyle.Dotted => XLBorderStyleValues.Dotted,
                Core.BorderStyle.Dashed => XLBorderStyleValues.Dashed,
                Core.BorderStyle.DashDot => XLBorderStyleValues.DashDot,
                Core.BorderStyle.DashDotDot => XLBorderStyleValues.DashDotDot,
                _ => XLBorderStyleValues.None
            };
        }

        private static XLAlignmentVerticalValues ConvertVerticalAlignment(Core.VerticalAlignment alignment)
        {
            return alignment switch
            {
                Core.VerticalAlignment.Top => XLAlignmentVerticalValues.Top,
                Core.VerticalAlignment.Center => XLAlignmentVerticalValues.Center,
                Core.VerticalAlignment.Bottom => XLAlignmentVerticalValues.Bottom,
                Core.VerticalAlignment.Justify => XLAlignmentVerticalValues.Justify,
                Core.VerticalAlignment.Distributed => XLAlignmentVerticalValues.Distributed,
                _ => XLAlignmentVerticalValues.Bottom
            };
        }

        private static XLAlignmentHorizontalValues ConvertHorizontalAlignment(Core.HorizontalAlignment alignment)
        {
            return alignment switch
            {
                Core.HorizontalAlignment.General => XLAlignmentHorizontalValues.General,
                Core.HorizontalAlignment.Left => XLAlignmentHorizontalValues.Left,
                Core.HorizontalAlignment.Center => XLAlignmentHorizontalValues.Center,
                Core.HorizontalAlignment.Right => XLAlignmentHorizontalValues.Right,
                Core.HorizontalAlignment.Fill => XLAlignmentHorizontalValues.Fill,
                Core.HorizontalAlignment.Justify => XLAlignmentHorizontalValues.Justify,
                Core.HorizontalAlignment.CenterContinuous => XLAlignmentHorizontalValues.CenterContinuous,
                Core.HorizontalAlignment.Distributed => XLAlignmentHorizontalValues.Distributed,
                _ => XLAlignmentHorizontalValues.General
            };
        }

        private static XLFillPatternValues ConvertFillStyle(Core.FillStyle fillStyle)
        {
            return fillStyle switch
            {
                Core.FillStyle.None => XLFillPatternValues.None,
                Core.FillStyle.Solid => XLFillPatternValues.Solid,
                Core.FillStyle.DarkGray => XLFillPatternValues.DarkGray,
                Core.FillStyle.MediumGray => XLFillPatternValues.MediumGray,
                Core.FillStyle.LightGray => XLFillPatternValues.LightGray,
                Core.FillStyle.Gray125 => XLFillPatternValues.Gray125,
                Core.FillStyle.Gray0625 => XLFillPatternValues.Gray0625,
                _ => XLFillPatternValues.None
            };
        }

        private static XLColor ConvertColor(CoreExcelColor excelColor)
        {
            return XLColor.FromArgb(excelColor.A, excelColor.R, excelColor.G, excelColor.B);
        }
    }
}
