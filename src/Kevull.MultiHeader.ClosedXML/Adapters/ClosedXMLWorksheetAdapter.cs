using ClosedXML.Excel;
using Kevull.MultiHeader.Core.Abstractions;

namespace Kevull.MultiHeader.ClosedXML.Adapters;

internal class ClosedXMLWorksheetAdapter : IExcelWorksheet
{
    private readonly IXLWorksheet _worksheet;
    
    public ClosedXMLWorksheetAdapter(IXLWorksheet worksheet)
    {
        _worksheet = worksheet;
    }
    
    
}