using ClosedXML.Excel;
using Kevull.MultiHeader.Core;
using Kevull.MultiHeader.ClosedXML.Adapters;
using Kevull.MultiHeader.Core.Abstractions;

namespace Kevull.MultiHeader.ClosedXML;
public class MultiHeaderReportClosedXML<T> : MultiHeaderReport<T>
{
    private readonly XLWorkbook _workbook;
    private readonly IXLWorksheet _sheet;
    
    protected override IExcelWorksheet Worksheet => 
        new ClosedXMLWorksheetAdapter(_sheet);
   
    
    public MultiHeaderReportClosedXML(XLWorkbook workbook, string sheetName)
    {
        _workbook = workbook;
        _sheet = workbook.Worksheets.Add(sheetName);
    }
    
    protected override void BuildDefaultHeaderStyle()
    {
        // Implementación específica de ClosedXML
    }
}