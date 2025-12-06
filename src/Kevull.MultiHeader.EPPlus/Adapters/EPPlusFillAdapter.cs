using System.Drawing;
using Kevull.MultiHeader.Core.Abstractions;
using OfficeOpenXml.Style;

namespace Kevull.MultiHeader.EPPlus.Adapters;

internal class EPPlusFillAdapter : IExcelFill
{
    private readonly ExcelFill _fill;

    public EPPlusFillAdapter(ExcelFill fill)
    {
        _fill = fill;
    }

    public object PatternType
    {
        get => _fill.PatternType;
        set => _fill.PatternType = (ExcelFillStyle)value;
    }

    public Color BackgroundColor
    {
        get => _fill.BackgroundColor.ToColor();
        set => _fill.BackgroundColor.SetColor(value);
    }

    public Color PatternColor
    {
        get => _fill.PatternColor.ToColor();
        set => _fill.PatternColor.SetColor(value);
    }
}
