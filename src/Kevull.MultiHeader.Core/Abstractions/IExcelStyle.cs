namespace Kevull.MultiHeader.Core.Abstractions;

public interface IExcelStyle
{
    IExcelFont Font { get; }
    IExcelBorder Border { get; }
    IExcelFill Fill { get; }
    IExcelNumberFormat Numberformat { get; }
    object HorizontalAlignment { get; set; }
    object VerticalAlignment { get; set; }
    bool WrapText { get; set; }
    object ReadingOrder { get; set; }
    bool ShrinkToFit { get; set; }
    int Indent { get; set; }
    int TextRotation { get; set; }
    bool Locked { get; set; }
    bool Hidden { get; set; }
    bool QuotePrefix { get; set; }
}