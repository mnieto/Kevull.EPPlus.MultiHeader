using Kevull.MultiHeader.Core;
using Kevull.MultiHeader.Core.Abstractions;
using Kevull.MultiHeader.EPPlus.Adapters;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;

namespace Kevull.MultiHeader.EPPlus;

public class MultiHeaderReportEPPlus<T> : MultiHeaderReport<T>
{
    private readonly ExcelPackage _xls;
    private readonly ExcelWorksheet _sheet;
    
    protected override IExcelWorksheet Worksheet => 
        new EPPlusWorksheetAdapter(_sheet);
    
    protected override IExcelWorkbook Workbook => 
        new EPPlusWorkbookAdapter(_xls);
    
    public MultiHeaderReportEPPlus(ExcelPackage xls, string sheetName)
    {
        _xls = xls;
        _sheet = AddSheet(xls, sheetName); 
    }

    public override MultiHeaderReport<T> Configure(Action<ConfigurationBuilder<T>> options)
    {
        var builder = new EPPlusConfigurationBuilder<T>(_xls);
        options?.Invoke(builder);
        _header = builder.Build();
        return this;
    }
    
    protected override void BuildDefaultHeaderStyle()
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

    protected override void BuildDateStyle()
    {
        if (_xls.Workbook.Styles.NamedStyles.FirstOrDefault(x => x.Name == StyleNames.DateStyleName) == null)
        {
            var namedStyle = _xls.Workbook.Styles.CreateNamedStyle(StyleNames.DateStyleName);
            namedStyle.Style.Numberformat.Format = StyleNames.DateFormat;
        }
    }

    protected override void BuildTimeStyle()
    {
        if (_xls.Workbook.Styles.NamedStyles.FirstOrDefault(x => x.Name == StyleNames.TimeStyleName) == null)
        {
            var namedStyle = _xls.Workbook.Styles.CreateNamedStyle(StyleNames.TimeStyleName);
            namedStyle.Style.Numberformat.Format = StyleNames.TimeFormat;
        }
    }

    private ExcelWorksheet AddSheet(ExcelPackage xls, string sheetName)
    {
        if (!xls.Workbook.Worksheets.AsEnumerable().Any(x => x.Name == sheetName))
        {
            xls.Workbook.Worksheets.Add(sheetName);
        }
        return xls.Workbook.Worksheets[sheetName];
    }
}