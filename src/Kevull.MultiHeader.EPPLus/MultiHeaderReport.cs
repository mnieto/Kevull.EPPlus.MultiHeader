using Kevull.MultiHeader.Core;
using Kevull.MultiHeader.EPPLus.Columns;
using OfficeOpenXml;
using System.Linq;
using System.Reflection;

namespace Kevull.MultiHeader.EPPLus
{
    /// <summary>
    /// Given an <see cref="IEnumerable{T}"/> list of objects it creates an in-memory Excel report
    /// </summary>
    /// <typeparam name="T">Type of objects</typeparam>
    public class MultiHeaderReport<T>
    {
        private ExcelWorksheet _sheet;
        private ExcelPackage _xls;
        private IExcelWriter _writer;

        private int FirstDataRow => (_header == null || !_header.AppendToExistingReport) ?
                                    _header?.FirstRow + _header?.Height ?? 2 :
                                    _sheet.Dimension.End.Row + 1;
        private int row;

        /// <summary>
        /// Internal <see cref="HeaderManager{T}"/>
        /// </summary>
        protected HeaderManager<T>? _header;

        internal const string HeaderStyleName = "__Headers__";

        /// <summary>
        /// Object properties associated to the columns
        /// </summary>
        protected Dictionary<string, PropertyInfo>? Properties { get; private set; }

        /// <summary>
        /// Ctor
        /// </summary>
        /// <param name="xls">Initialized <see cref="ExcelPackage"/></param>
        /// <param name="sheet">Existing worksheet where generate the report</param>
        public MultiHeaderReport(ExcelPackage xls, ExcelWorksheet sheet)
        {
            _xls = xls;
            _sheet = sheet;
            _writer = new EPPlusExcelWriter(xls, sheet);
        }

        /// <summary>
        /// Ctor
        /// </summary>
        /// <param name="xls">Initialized <see cref="ExcelPackage"/></param>
        /// <param name="sheetName">Worksheet name to be created where generate the report</param>
        public MultiHeaderReport(ExcelPackage xls, string sheetName): this(xls, AddSheet(xls, sheetName)) { }

        /// <summary>
        /// Customize the columns and formats during the report generation. See <see cref="ConfigurationBuilder{T}"/>.
        /// </summary>
        /// <param name="options">Lambda expresion to configure the report</param>
        /// <returns><see cref="MultiHeaderReport{T}"/>This allows a fluent style to configure and generate the report</returns>
        public MultiHeaderReport<T> Configure(Action<ConfigurationBuilder<T>> options)
        {
            var builder = new ConfigurationBuilder<T>(_xls);
            options?.Invoke(builder);
            _header = builder.Build();
            return this;
        }


        /// <summary>
        /// Generate the report in Excel
        /// </summary>
        /// <param name="data">Data of tyepe <typeparamref name="T"/></param>
        /// <remarks>If there is any configuration, it will generate the report using the default conventions</remarks>
        public void GenerateReport(IEnumerable<T> data)
        {
            //If no configuration is provided, use default simple headers
            if (_header == null)
            {
                _header = new HeaderManager<T>();
            } else
            {
                _header.BuildHeaders();
            }
            Properties = _header.Properties;
            if (!_header.AppendToExistingReport)
                WriteHeaders();

            row = FirstDataRow;
            foreach (T item in data)
            {
                ProcessRow(item);
            }
            DoFormatting();
            CalulateFormulas();
        }

        internal void Save(string fileName)
        {
            _xls.SaveAs(fileName);
        }

        private static ExcelWorksheet AddSheet(ExcelPackage xls, string sheetName)
        {
            if (!xls.Workbook.Worksheets.AsEnumerable().Any(x => x.Name == sheetName))
            {
                xls.Workbook.Worksheets.Add(sheetName);
            }
            return xls.Workbook.Worksheets[sheetName];
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
            foreach(var columnInfo in header.Columns)
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
            header = header ?? _header!;
            int row = topRow ?? _header!.FirstRow;
            foreach (var columnInfo in header.Columns)
            {
                columnInfo.WriteHeader(_writer, row, columnInfo.Index);
                columnInfo.FormatHeader(_writer, row, columnInfo.Index, columnInfo.HasChildren ? 1 : header.Height - (row - _header!.FirstRow));
                if (columnInfo.HasChildren)
                {
                    WriteHeaders(columnInfo.Header!, row + 1);
                }
            }
        }

        private void DoFormatting()
        {
            if (_header!.AutoFreezePanes)
                _writer.FreezePanes(_header.FirstRow + _header!.Height, _header.FirstColumn);

            //Hide columns if needed
            foreach (var columnInfo in _header!.Columns.Where(x => x.Hidden || x.ColumnWidth.Type == WidthType.Hidden ))
            {
                _writer.SetColumnHidden(columnInfo.Index, true);
            }

            //Autofilter
            int lastHeaderRow = _header.FirstRow + _header.Height - 1;
            int lastHeaderColumn = _header.FirstColumn + _header.Width - 1;
            _writer.SetAutoFilter(lastHeaderRow, _header!.Columns.Min(x => x.Index), lastHeaderRow, lastHeaderColumn, _header.AutoFilter);

            //Width
            foreach(var columnInfo in _header!.Columns.Where(x => x.ColumnWidth.Type == WidthType.Auto))
            {
                double minWidth = columnInfo.ColumnWidth.MinimumWidth == double.MinValue ? _sheet.DefaultColWidth : columnInfo.ColumnWidth.MinimumWidth;
                double maxWidth = columnInfo.ColumnWidth.MaximunWidth;
                _writer.AutoFitColumn(columnInfo.Index, minWidth, maxWidth);
            }
            foreach (var columnInfo in _header!.Columns.Where(x => x.ColumnWidth.Type == WidthType.Custom))
            {
                _writer.SetColumnWidth(columnInfo.Index, columnInfo.ColumnWidth.Width!.Value);
            }

            //Styles
            BuildDefaultHeaderStyle();
            BuildDateStyle();
            BuildTimeStyle();

            if (!_header!.AppendToExistingReport)
            {
                _writer.ApplyNamedStyle(_header.FirstRow, _header!.Columns.Min(x => x.Index), lastHeaderRow, lastHeaderColumn, StyleNames.HeaderStyleName);
            }

            foreach (var columnInfo in _header!.Columns.Where(x => x.StyleName != null))
            {
                _writer.ApplyNamedStyle(FirstDataRow, columnInfo.Index, _sheet.Dimension.End.Row, columnInfo.Index, columnInfo.StyleName!);
            }
        }

        private void CalulateFormulas()
        {
            bool NeedsCalculate = false;
            foreach (var columnInfo in _header!.Columns.OfType<ColumnFormula>())
            {
                // Optimized: write formula to entire range at once instead of row by row
                int lastRow = _sheet.Dimension?.End.Row ?? FirstDataRow;
                if (lastRow >= FirstDataRow)
                {
                    columnInfo.WriteCell(_writer, FirstDataRow, columnInfo.Index, lastRow, columnInfo.Index, Properties!, null);
                    NeedsCalculate = true;
                }
            }
            if (NeedsCalculate)
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

    }

    internal class StyleNames
    {
        public const string HeaderStyleName =  "__Headers__";
        public const string DateStyleName = "__date__";
        public const string TimeStyleName = "__time__";

        internal const string TimeFormat = "[$-x-systime]h:mm:ss AM/PM";    //This format depends on local system settings
        internal const string DateFormat = "mm-dd-yy";         //This format depends on local system settings
    }
}