using Kevull.MultiHeader.Core.Abstractions;
using OfficeOpenXml.Style;

namespace Kevull.MultiHeader.EPPlus.Adapters;

internal class EPPlusNumberFormatAdapter : IExcelNumberFormat
{
    private readonly ExcelNumberFormat _numberFormat;

    public EPPlusNumberFormatAdapter(ExcelNumberFormat numberFormat)
    {
        _numberFormat = numberFormat;
    }

    public string Format
    {
        get => _numberFormat.Format;
        set => _numberFormat.Format = value;
    }

}
