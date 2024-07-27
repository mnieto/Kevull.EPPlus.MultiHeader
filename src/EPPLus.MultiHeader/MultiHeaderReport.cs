using EPPLus.MultiHeader.Columns;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;
using System.Linq;
using System.Reflection;

namespace EPPLus.MultiHeader
{
    /// <summary>
    /// Given an <see cref="IEnumerable{T}"/> list of objects it creates an in-memory Excel report
    /// </summary>
    /// <typeparam name="T">Type of objects</typeparam>
    public class MultiHeaderReport<T>
    {
        private ExcelWorksheet _sheet;
        private ExcelPackage _xls;

        private int FirstDataRow => _header?.Height + 1 ?? 2;
        private int row;
        
        /// <summary>
        /// Internal <see cref="HeaderManager{T}"/>
        /// </summary>
        protected HeaderManager<T>? _header;

        internal const string HeaderStyleName = "Headers";

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
            var builder = new ConfigurationBuilder<T>();
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
            WriteHeaders();

            row = FirstDataRow;
            foreach (T item in data)
            {
                ProcessRow(item);
            }
            DoFormatting();
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

        private void WriteHeaders(HeaderManager? header = null, int row = HeaderManager.FirstRow)
        {
            header = header ?? _header!;
            foreach (var columnInfo in header.Columns)
            {
                var cell = _sheet.Cells[row, columnInfo.Index];
                columnInfo.WriteHeader(cell);
                columnInfo.FormatHeader(cell, columnInfo.HasChildren ? 1 : header.Height - row + 1);
                if (columnInfo.HasChildren)
                {
                    WriteHeaders(columnInfo.Header!, row + 1);
                }
            }
        }

        private void DoFormatting()
        {
            foreach (var columnInfo in _header!.Columns.Where(x => x.Hidden))
            {
                _sheet.Column(columnInfo.Index).Hidden = true;
            }

            int lastHeaderRow = HeaderManager.FirstRow + _header.Height - 1;
            _sheet.Cells[lastHeaderRow, _header!.Columns.Min(x => x.Index), lastHeaderRow, _header.Width].AutoFilter = _header.AutoFilter;

            var rangeHeader = _sheet.Cells[HeaderManager.FirstRow, _header!.Columns.Min(x => x.Index), lastHeaderRow, _header.Width];
            if (_xls.Workbook.Styles.NamedStyles.FirstOrDefault(x => x.Name == HeaderStyleName) == null)
            {
                BuildDefaultStyle();

            }
            rangeHeader.StyleName = HeaderStyleName;

            bool NeedsCalculate = false;
            foreach(var columnInfo in _header!.Columns.OfType<ColumnFormula>())
            {
                var range = _sheet.Cells[FirstDataRow, columnInfo.Order!.Value, _sheet.Dimension.End.Row, columnInfo.Order!.Value];
                columnInfo.WriteCell(range, Properties!, null);
                NeedsCalculate = true;
            }
            if (NeedsCalculate)
                _sheet.Calculate();
        }

        private void BuildDefaultStyle()
        {
            var namedStyle = _xls.Workbook.Styles.CreateNamedStyle(HeaderStyleName);
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
}