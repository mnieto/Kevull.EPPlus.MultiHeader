using OfficeOpenXml;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using System.CodeDom.Compiler;
using System.Reflection;

namespace EPPLus.MultiHeader
{
    public class MultiHeaderReport<T>
    {
        private ExcelWorksheet _sheet;
        private ExcelPackage _xls;

        private int FirstDataRow = 2;
        private int row;
        protected HeaderManager<T>? _header;

        protected Dictionary<string, PropertyInfo>? Properties { get; private set; }

        public MultiHeaderReport(ExcelPackage xls, ExcelWorksheet sheet)
        {
            _xls = xls;
            _sheet = sheet;
        }

        public MultiHeaderReport(ExcelPackage xls, string sheetName): this(xls, AddSheet(xls, sheetName)) { }

        public MultiHeaderReport<T> Configure(Action<ConfigurationBuilder<T>> options)
        {
            var builder = new ConfigurationBuilder<T>();
            options?.Invoke(builder);
            _header = new HeaderManager<T>(builder.Build());
            return this;
        }


        public void GenerateReport(IEnumerable<T> data)
        {
            //If no configuration is provided, use default simple headers
            if (_header == null)
            {
                _header = new HeaderManager<T>();
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
            int col = 1;
            foreach (var columnInfo in _header!.Columns)
            {
                columnInfo.WriteCell(_sheet.Cells[row, col++], Properties!, item!);
            }
            row++;
        }

        private void WriteHeaders()
        {
            int col = 1;
            row = 1;
            foreach (var columnInfo in _header!.Columns)
            {
                _sheet.Cells[row, col++].Value = columnInfo.DisplayName;
            }
            FirstDataRow = row + 1;
        }

        private void DoFormatting()
        {
            foreach (var columnInfo in _header!.Columns.Where(x => x.Hidden))
            {
                _sheet.Column(columnInfo.Order!.Value).Hidden = true;
            }
        }

    }
}