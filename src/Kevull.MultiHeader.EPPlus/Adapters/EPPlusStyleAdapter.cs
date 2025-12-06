using Kevull.MultiHeader.Core.Abstractions;
using OfficeOpenXml.Style;

namespace Kevull.MultiHeader.EPPlus.Adapters;

internal class EPPlusStyleAdapter : IExcelStyle
{
    private readonly ExcelStyle _style;
    private readonly IExcelFont _font;
    private readonly IExcelBorder _border;
    private readonly IExcelFill _fill;
    private readonly IExcelNumberFormat _numberFormat;

    public EPPlusStyleAdapter(ExcelStyle style)
    {
        _style = style;
        _font = new EPPlusFontAdapter(style.Font);
        _border = new EPPlusBorderAdapter(style.Border);
        _fill = new EPPlusFillAdapter(style.Fill);
        _numberFormat = new EPPlusNumberFormatAdapter(style.Numberformat);
    }

    public IExcelFont Font => _font;
    public IExcelBorder Border => _border;
    public IExcelFill Fill => _fill;
    public IExcelNumberFormat Numberformat => _numberFormat;

    public object HorizontalAlignment
    {
        get => _style.HorizontalAlignment;
        set => _style.HorizontalAlignment = (ExcelHorizontalAlignment)value;
    }

    public object VerticalAlignment
    {
        get => _style.VerticalAlignment;
        set => _style.VerticalAlignment = (ExcelVerticalAlignment)value;
    }

    public bool WrapText
    {
        get => _style.WrapText;
        set => _style.WrapText = value;
    }

    public object ReadingOrder
    {
        get => _style.ReadingOrder;
        set => _style.ReadingOrder = (ExcelReadingOrder)value;
    }

    public bool ShrinkToFit
    {
        get => _style.ShrinkToFit;
        set => _style.ShrinkToFit = value;
    }

    public int Indent
    {
        get => _style.Indent;
        set => _style.Indent = value;
    }

    public int TextRotation
    {
        get => _style.TextRotation;
        set => _style.TextRotation = value;
    }

    public bool Locked
    {
        get => _style.Locked;
        set => _style.Locked = value;
    }

    public bool Hidden
    {
        get => _style.Hidden;
        set => _style.Hidden = value;
    }

    public bool QuotePrefix
    {
        get => _style.QuotePrefix;
        set => _style.QuotePrefix = value;
    }
}
