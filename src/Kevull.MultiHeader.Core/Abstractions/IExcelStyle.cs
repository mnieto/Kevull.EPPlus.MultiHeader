// Kevull.MultiHeader.Core/Abstractions/IExcelStyle.cs
public interface IExcelStyle
{
    IExcelFont Font { get; }
    IExcelBorder Border { get; }
    IExcelFill Fill { get; }
    IExcelNumberFormat NumberFormat { get; }
    ExcelHorizontalAlignment HorizontalAlignment { get; set; }
    ExcelVerticalAlignment VerticalAlignment { get; set; }
}