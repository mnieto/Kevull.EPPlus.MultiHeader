using Kevull.MultiHeader.Core.Abstractions;
using OfficeOpenXml;

namespace Kevull.MultiHeader.EPPlus.Adapters;

internal class EPPlusWorkbookAdapter : IExcelWorkbook
{
    private readonly ExcelPackage _package;

    public EPPlusWorkbookAdapter(ExcelPackage package)
    {
        _package = package;
    }

    public IExcelWorksheet AddWorksheet(string name)
    {
        var worksheet = _package.Workbook.Worksheets.Add(name);
        return new EPPlusWorksheetAdapter(worksheet);
    }

    public void Calculate()
    {
        _package.Workbook.Calculate();
    }

    public void SaveAs(string filePath)
    {
        _package.SaveAs(filePath);
    }
}
