using Kevull.EPPLus.MultiHeader.Columns;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;
using System.Linq;
using System.Reflection;

namespace Kevull.EPPLus.MultiHeader
{
    /// <summary>
    /// Given an <see cref="IEnumerable{T}"/> list of objects it creates an in-memory Excel report
    /// </summary>
    /// <typeparam name="T">Type of objects</typeparam>
    public class MultiHeaderReport<T>
    {
        private ExcelWorksheet _sheet;
        private ExcelPackage _xls;

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
                    columnInfo.WriteCell(_sheet.Cells[row, columnInfo.Index], Properties!, item!);
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
                    columnInfo.WriteCell(_sheet.Cells[row, columnInfo.Index], header.Properties, item);
                }
            }

        }

        private void WriteHeaders(HeaderManager? header = null, int? topRow = null)
        {
            header = header ?? _header!;
            int row = topRow ?? _header!.FirstRow;
            foreach (var columnInfo in header.Columns)
            {
                var cell = _sheet.Cells[row, columnInfo.Index];
                columnInfo.WriteHeader(cell);
                columnInfo.FormatHeader(cell, columnInfo.HasChildren ? 1 : header.Height - (header.FirstRow - row));
                if (columnInfo.HasChildren)
                {
                    WriteHeaders(columnInfo.Header!, row + 1);
                }
            }
        }

        private void DoFormatting()
        {
            //Hide columns if needed
            foreach (var columnInfo in _header!.Columns.Where(x => x.Hidden || x.ColumnWidth.Type == WidthType.Hidden ))
            {
                _sheet.Column(columnInfo.Index).Hidden = true;
            }

            //Autofilter
            int lastHeaderRow = _header.FirstRow + _header.Height - 1;
            _sheet.Cells[lastHeaderRow, _header!.Columns.Min(x => x.Index), lastHeaderRow, _header.Width].AutoFilter = _header.AutoFilter;

            //Width
            foreach(var columnInfo in _header!.Columns.Where(x => x.ColumnWidth.Type == WidthType.Auto))
            {
                double minWidth = columnInfo.ColumnWidth.MinimumWidth == double.MinValue ? _sheet.DefaultColWidth : columnInfo.ColumnWidth.MinimumWidth;
                double maxWidth = columnInfo.ColumnWidth.MaximunWidth;
                _sheet.Column(columnInfo.Index).AutoFit(minWidth, maxWidth);
            }
            foreach (var columnInfo in _header!.Columns.Where(x => x.ColumnWidth.Type == WidthType.Custom))
            {
                _sheet.Column(columnInfo.Index).Width = columnInfo.ColumnWidth.Width!.Value;
            }

            //Styles
            BuildDefaultHeaderStyle();
            BuildDateStyle();
            BuildTimeStyle();

            if (!_header!.AppendToExistingReport)
            {
                var rangeHeader = _sheet.Cells[_header.FirstRow, _header!.Columns.Min(x => x.Index), lastHeaderRow, _header.Width];
                rangeHeader.StyleName = StyleNames.HeaderStyleName;
            }

            foreach (var columnInfo in _header!.Columns.Where(x => x.StyleName != null))
            {
                var range = _sheet.Cells[FirstDataRow, columnInfo.Index, _sheet.Dimension.End.Row, columnInfo.Index];
                range.StyleName = columnInfo.StyleName;
            }
        }

        private void CalulateFormulas()
        {
            bool NeedsCalculate = false;
            foreach (var columnInfo in _header!.Columns.OfType<ColumnFormula>())
            {
                var range = _sheet.Cells[FirstDataRow, columnInfo.Index, _sheet.Dimension.End.Row, columnInfo.Index];
                columnInfo.WriteCell(range, Properties!, null);
                NeedsCalculate = true;
            }
            if (NeedsCalculate)
                _sheet.Calculate();
        }

        private void BuildDefaultHeaderStyle()
        {
            if (_xls.Workbook.Styles.NamedStyles.FirstOrDefault(x => x.Name == StyleNames.HeaderStyleName) == null)
            {
                var namedStyle = _xls.Workbook.Styles.CreateNamedStyle(StyleNames.HeaderStyleName);
                namedStyle.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                namedStyle.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                namedStyle.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                namedStyle.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                namedStyle.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                namedStyle.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                namedStyle.Style.Fill.SetBackground(Color.LightGray, ExcelFillStyle.Solid);
                namedStyle.Style.Font.Bold = true;
            }
        }

        private void BuildDateStyle()
        {
            if (_xls.Workbook.Styles.NamedStyles.FirstOrDefault(x => x.Name == StyleNames.DateStyleName) == null)
            {
                var namedStyle = _xls.Workbook.Styles.CreateNamedStyle(StyleNames.DateStyleName);
                namedStyle.Style.Numberformat.Format = StyleNames.DateFormat;
            }
        }

        private void BuildTimeStyle()
        {
            if (_xls.Workbook.Styles.NamedStyles.FirstOrDefault(x => x.Name == StyleNames.TimeStyleName) == null)
            {
                var namedStyle = _xls.Workbook.Styles.CreateNamedStyle(StyleNames.TimeStyleName);
                namedStyle.Style.Numberformat.Format = StyleNames.TimeFormat;
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