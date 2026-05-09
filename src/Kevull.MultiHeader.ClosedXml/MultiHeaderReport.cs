using ClosedXML.Excel;
using Kevull.MultiHeader.Core;
using Kevull.MultiHeader.Core.Columns;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace Kevull.MultiHeader.ClosedXml
{
    /// <summary>
    /// Given an <see cref="IEnumerable{T}"/> list of objects it creates an in-memory Excel report.
    /// </summary>
    /// <typeparam name="T">Type of objects.</typeparam>
    public class MultiHeaderReport<T> : IMultiHeaderReport<T>
    {
        private readonly IXLWorksheet _sheet;
        private readonly XLWorkbook _xls;
        private readonly IExcelWriter _writer;

        private int FirstDataRow => (_header == null || !_header.AppendToExistingReport)
            ? _header?.FirstRow + _header?.Height ?? 2
            : (_sheet.LastRowUsed()?.RowNumber() ?? 0) + 1;

        private int row;

        /// <summary>
        /// Internal <see cref="HeaderManager{T}"/>.
        /// </summary>
        protected HeaderManager<T>? _header;

        /// <summary>
        /// Custom styles defined by the user to be applied to headers or columns.
        /// </summary>
        protected Dictionary<string, CellFormat> _namedStyles = [];

        /// <summary>
        /// Object properties associated to the columns.
        /// </summary>
        protected Dictionary<string, PropertyInfo>? Properties { get; private set; }

        /// <summary>
        /// Ctor.
        /// </summary>
        /// <param name="xls">Initialized <see cref="XLWorkbook"/>.</param>
        /// <param name="sheet">Existing worksheet where to generate the report.</param>
        public MultiHeaderReport(XLWorkbook xls, IXLWorksheet sheet)
        {
            _xls = xls ?? throw new ArgumentNullException(nameof(xls));
            _sheet = sheet ?? throw new ArgumentNullException(nameof(sheet));
            _writer = new ClosedXmlExcelWriter(xls, sheet);
        }

        /// <summary>
        /// Ctor.
        /// </summary>
        /// <param name="xls">Initialized <see cref="XLWorkbook"/>.</param>
        /// <param name="sheetName">Worksheet name to be created where generate the report.</param>
        public MultiHeaderReport(XLWorkbook xls, string sheetName) : this(xls, AddSheet(xls, sheetName)) { }

        /// <inheritdoc />
        public IMultiHeaderReport<T> Configure(Action<IConfigurationBuilder<T>> options)
        {
            var builder = new ConfigurationBuilder<T>(_xls);
            options?.Invoke(builder);
            _header = builder.Build();
            _namedStyles = builder.NamedStyles;
            return this;
        }

        /// <inheritdoc />
        public void GenerateReport(IEnumerable<T> data)
        {
            if (_header == null)
            {
                _header = new HeaderManager<T>();
            }
            else
            {
                _header.BuildHeaders();
            }

            Properties = _header.Properties;
            if (!_header.AppendToExistingReport)
                WriteHeaders();

            row = FirstDataRow;
            foreach (var item in data)
            {
                ProcessRow(item);
            }

            DoFormatting();
            CalulateFormulas();
        }

        /// <inheritdoc />
        public void Save(string fileName)
        {
            _xls.SaveAs(fileName);
        }

        private static IXLWorksheet AddSheet(XLWorkbook xls, string sheetName)
        {
            if (!xls.Worksheets.Any(x => x.Name == sheetName))
            {
                xls.AddWorksheet(sheetName);
            }
            return xls.Worksheet(sheetName);
        }

        private void ProcessRow(T item)
        {
            foreach (var columnInfo in _header!.Columns)
            {
                if (columnInfo.HasChildren)
                {
                    ProcessRow(columnInfo.Header!, Properties![columnInfo.Name].GetValue(item));
                }
                else
                {
                    columnInfo.WriteCell(_writer, row, columnInfo.Index, Properties!, item!);
                }
            }
            row++;
        }

        private void ProcessRow(HeaderManager header, object? item)
        {
            if (item == null)
                return;
            if (header.Properties == null)
                throw new ArgumentNullException(nameof(header.Properties));

            foreach (var columnInfo in header.Columns)
            {
                if (columnInfo.HasChildren)
                {
                    ProcessRow(columnInfo.Header!, header.Properties[columnInfo.Name].GetValue(item));
                }
                else
                {
                    columnInfo.WriteCell(_writer, row, columnInfo.Index, header.Properties, item);
                }
            }
        }

        private void WriteHeaders(HeaderManager? header = null, int? topRow = null)
        {
            header ??= _header!;
            int localRow = topRow ?? _header!.FirstRow;
            foreach (var columnInfo in header.Columns)
            {
                columnInfo.WriteHeader(_writer, localRow, columnInfo.Index);
                columnInfo.FormatHeader(_writer, localRow, columnInfo.Index, columnInfo.HasChildren ? 1 : header.Height - (localRow - _header!.FirstRow));
                if (columnInfo.HasChildren)
                {
                    WriteHeaders(columnInfo.Header!, localRow + 1);
                }
            }
        }

        private void DoFormatting()
        {
            if (_header!.AutoFreezePanes)
                _writer.FreezePanes(_header.FirstRow + _header.Height, _header.FirstColumn);

            foreach (var columnInfo in _header.Columns.Where(x => x.Hidden || x.ColumnWidth.Type == WidthType.Hidden))
            {
                _writer.SetColumnHidden(columnInfo.Index, true);
            }

            int lastHeaderRow = _header.FirstRow + _header.Height - 1;
            int lastHeaderColumn = _header.FirstColumn + _header.Width - 1;
            _writer.SetAutoFilter(lastHeaderRow, _header.Columns.Min(x => x.Index), lastHeaderRow, lastHeaderColumn, _header.AutoFilter);

            foreach (var columnInfo in _header.Columns.Where(x => x.ColumnWidth.Type == WidthType.Auto))
            {
                double minWidth = columnInfo.ColumnWidth.MinimumWidth == double.MinValue ? _sheet.ColumnWidth : columnInfo.ColumnWidth.MinimumWidth;
                double maxWidth = columnInfo.ColumnWidth.MaximunWidth;
                _writer.AutoFitColumn(columnInfo.Index, minWidth, maxWidth);
            }
            foreach (var columnInfo in _header.Columns.Where(x => x.ColumnWidth.Type == WidthType.Custom))
            {
                _writer.SetColumnWidth(columnInfo.Index, columnInfo.ColumnWidth.Width!.Value);
            }

            BuildDateStyle();
            BuildTimeStyle();
            BuildColumnStyles();

            if (!_header.AppendToExistingReport)
            {
                BuildDefaultHeaderStyle();
                _writer.ApplyNamedStyle(_header.FirstRow, _header.Columns.Min(x => x.Index), lastHeaderRow, lastHeaderColumn, StyleNames.HeaderStyleName);
            }

            int lastDataRow = _sheet.LastRowUsed()?.RowNumber() ?? (FirstDataRow - 1);
            if (lastDataRow >= FirstDataRow)
            {
                foreach (var columnInfo in _header.Columns.Where(x => x.StyleName != null))
                {
                    _writer.ApplyNamedStyle(FirstDataRow, columnInfo.Index, lastDataRow, columnInfo.Index, columnInfo.StyleName!);
                }
            }
        }

        private void CalulateFormulas()
        {
            bool needsCalculate = false;
            foreach (var columnInfo in _header!.Columns.OfType<ColumnFormula>())
            {
                int lastRow = _sheet.LastRowUsed()?.RowNumber() ?? FirstDataRow;
                if (lastRow >= FirstDataRow)
                {
                    columnInfo.WriteCell(_writer, FirstDataRow, columnInfo.Index, lastRow, columnInfo.Index, Properties!, null);
                    needsCalculate = true;
                }
            }
            if (needsCalculate)
                _writer.Recalculate();
        }

        private void BuildDefaultHeaderStyle()
        {
            if (!_writer.NamedStyleExists(StyleNames.HeaderStyleName))
            {
                var format = new CellFormat
                {
                    LeftBorder = BorderStyle.Thin,
                    RightBorder = BorderStyle.Thin,
                    TopBorder = BorderStyle.Thin,
                    BottomBorder = BorderStyle.Thin,
                    VerticalAlignment = VerticalAlignment.Center,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    BackgroundColor = ExcelColor.LightGray,
                    FillStyle = FillStyle.Solid,
                    Bold = true
                };

                if (_namedStyles.ContainsKey(StyleNames.HeaderStyleName))
                    format.Merge(_namedStyles[StyleNames.HeaderStyleName]);
                _writer.CreateNamedStyle(StyleNames.HeaderStyleName, format);
            }
        }

        private void BuildDateStyle()
        {
            if (!_writer.NamedStyleExists(StyleNames.DateStyleName))
            {
                var format = new CellFormat
                {
                    NumberFormat = StyleNames.DateFormat
                };
                _writer.CreateNamedStyle(StyleNames.DateStyleName, format);
            }
        }

        private void BuildTimeStyle()
        {
            if (!_writer.NamedStyleExists(StyleNames.TimeStyleName))
            {
                var format = new CellFormat
                {
                    NumberFormat = StyleNames.TimeFormat
                };
                _writer.CreateNamedStyle(StyleNames.TimeStyleName, format);
            }
        }

        private void BuildColumnStyles()
        {
            foreach (var style in _namedStyles.Keys)
            {
                if (style != StyleNames.HeaderStyleName && !_writer.NamedStyleExists(style))
                {
                    _writer.CreateNamedStyle(style, _namedStyles[style]);
                }
            }
        }
    }
}
